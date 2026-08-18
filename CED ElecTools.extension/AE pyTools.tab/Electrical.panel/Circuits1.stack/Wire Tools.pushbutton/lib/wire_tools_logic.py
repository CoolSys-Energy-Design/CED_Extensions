# -*- coding: utf-8 -*-
"""Revit-side logic for the modeless multi-scheme Wire Tools command."""

import math

from Autodesk.Revit.UI.Selection import ISelectionFilter
from pyrevit import DB, UI, script

from Snippets import _elecutils as electrical_utils
from Snippets import design_options, revit_helpers
from Snippets.wireutils import (
    are_points_coincident,
    collect_active_view_wires_by_circuit,
    build_local_frame,
    find_previous_connector_for_homerun,
    get_element_connector_from_wire_connector,
    get_wire_type_id,
    is_homerun_wire,
    wire_connected_unconnected_connectors,
)

HOME_RUN_LENGTH = 4.0
GEOMETRY_TOLERANCE = 1e-6
TAG_EXISTING_SKIP = "skip_existing"
TAG_EXISTING_REPLACE = "replace_existing"
TAG_OFFSET_PAPER_INCHES = 1.0 / 4.0
TAG_OFFSET_MINIMUM = 0.50
TAG_OFFSET_MAXIMUM = 2.50
TAG_OFFSET_BOUNDING_BOX_RATIO = 0.20
HOMERUN_DIRECTION_PANEL = "panel"
HOMERUN_DIRECTION_DEVICE = "device"
HOMERUN_SHAPE_STRAIGHT = "straight"
HOMERUN_SHAPE_BEND = "bend"
INTERCONNECT_SCOPE_SELECTED = "selected_only"
INTERCONNECT_SCOPE_CIRCUITS = "selected_circuits"

SCHEME_WIRE_BY_CIRCUIT = "wire_by_circuit"
SCHEME_INTERCONNECT = "interconnect"
SCHEME_INDIVIDUAL_HOMERUN = "individual_homerun"
SCHEME_WIRE_TO_NODE = "wire_to_node"

SCHEME_LABELS = {
    SCHEME_WIRE_BY_CIRCUIT: "Wire by Circuit",
    SCHEME_INTERCONNECT: "Interconnect",
    SCHEME_INDIVIDUAL_HOMERUN: "Individual Homeruns",
    SCHEME_WIRE_TO_NODE: "Wire to Node",
}

SELECTION_RULES = {
    SCHEME_WIRE_BY_CIRCUIT: (
        "Main-model, non-annotation MEP elements and electrical-system "
        "objects are accepted initially. A selected element is valid only "
        "when the shared electrical-system resolver finds at least one "
        "eligible power circuit."
    ),
    SCHEME_INTERCONNECT: (
        "Main-model FamilyInstance elements are accepted only when they have "
        "at least one usable electrical connector: electrical domain when "
        "available, not Logical, and with ElectricalSystemType or electrical "
        "domain. No electrical system is required during selection. With "
        "circuit scope, Revit wires every selected circuit and the tool keeps "
        "one total homerun before connecting the circuit representatives. "
        "With selected-only scope, every selected element is connected into a "
        "spatial nearest-neighbor chain and one selected element receives the "
        "configurable homerun. If the selected set has no usable common connector "
        "type, interconnection is rejected."
    ),
    SCHEME_INDIVIDUAL_HOMERUN: (
        "Main-model FamilyInstance elements are accepted only when they have "
        "at least one usable electrical connector using the same electrical "
        "domain/non-Logical rules. The selected system type is resolved first; "
        "a unique primary connector is used only when several matching "
        "connectors share that type."
    ),
    SCHEME_WIRE_TO_NODE: (
        "Target elements must be main-model FamilyInstance elements with at "
        "least one usable electrical connector using the same electrical "
        "domain/non-Logical rules. The separately selected node must meet the "
        "same rule. No electrical system is required, but each target "
        "connector type must match the node connector type at run time."
    ),
}


def element_id_value(element_id):
    try:
        if not hasattr(element_id, "Value") and not hasattr(element_id, "IntegerValue"):
            return int(element_id)
    except Exception:
        pass
    return revit_helpers.get_elementid_value(element_id)


def element_id_from(value):
    return revit_helpers.elementid_from_value(value)


def safe_element_name(element, fallback="<Unnamed>"):
    if element is None:
        return fallback
    try:
        name = DB.Element.Name.__get__(element)
        if name:
            return str(name)
    except Exception:
        pass
    try:
        name = revit_helpers.get_family_symbol_name(
            element,
            doc=element.Document,
            fallback="",
        )
        if name:
            return str(name)
    except Exception:
        pass
    return fallback


def is_supported_wire_view(view):
    if view is None or not isinstance(view, DB.ViewPlan):
        return False
    floor_plan = getattr(DB.ViewType, "FloorPlan", None)
    ceiling_plan = getattr(DB.ViewType, "CeilingPlan", None)
    if floor_plan is not None and view.ViewType == floor_plan:
        return True
    if ceiling_plan is not None and view.ViewType == ceiling_plan:
        return True
    return False


def supported_view_text(view):
    try:
        return str(view.ViewType)
    except Exception:
        return "This view"


def wiring_type_from_name(name, fallback_name):
    wiring_types = {
        "Arc": DB.Electrical.WiringType.Arc,
        "Chamfer": DB.Electrical.WiringType.Chamfer,
    }
    return wiring_types.get(name, wiring_types[fallback_name])


def wire_type_choices(document):
    choices = []
    collector = DB.FilteredElementCollector(document)
    wire_types = collector.OfClass(DB.Electrical.WireType).ToElements()
    for wire_type in wire_types:
        name_parameter = wire_type.get_Parameter(
            DB.BuiltInParameter.ALL_MODEL_TYPE_NAME
        )
        name_value = name_parameter.AsString() if name_parameter else None
        choices.append({
            "id": element_id_value(wire_type.Id),
            "name": name_value or safe_element_name(wire_type),
        })
    choices.sort(key=lambda item: str(item["name"]).lower())
    return choices


def wire_tag_type_choices(document):
    choices = []
    wire_tag_category = getattr(DB.BuiltInCategory, "OST_WireTags", None)
    if wire_tag_category is not None:
        collector = DB.FilteredElementCollector(document)
        type_elements = collector.OfCategory(wire_tag_category).WhereElementIsElementType()
    else:
        collector = DB.FilteredElementCollector(document)
        type_elements = collector.OfClass(DB.FamilySymbol).WhereElementIsElementType()

    for type_element in type_elements:
        try:
            category_name = str(type_element.Category.Name)
        except Exception:
            category_name = ""
        if wire_tag_category is None and "wire" not in category_name.lower():
            continue
        choices.append({
            "id": element_id_value(type_element.Id),
            "name": safe_element_name(type_element),
        })
    choices.sort(key=lambda item: str(item["name"]).lower())
    return choices


def _connector_manager(element):
    if element is None:
        return None
    try:
        return element.ConnectorManager
    except Exception:
        pass
    try:
        return element.MEPModel.ConnectorManager
    except Exception:
        return None


def _connector_is_electrical(connector):
    if connector is None:
        return False
    domain = getattr(connector, "Domain", None)
    domain_enum = getattr(DB, "Domain", None)
    electrical_domain = getattr(domain_enum, "DomainElectrical", None)
    if domain is not None and electrical_domain is not None:
        if domain != electrical_domain:
            return False
    try:
        connector_type = connector.ConnectorType
        connector_type_enum = getattr(DB, "ConnectorType", None)
        logical_type = getattr(connector_type_enum, "Logical", None)
        if logical_type is not None and connector_type == logical_type:
            return False
    except Exception:
        pass
    try:
        connector.ElectricalSystemType
        return True
    except Exception:
        return (
            electrical_domain is not None
            and domain is not None
            and domain == electrical_domain
        )


def connector_type_key(connector):
    try:
        electrical_type = connector.ElectricalSystemType
        return "ElectricalSystemType:{}".format(str(electrical_type))
    except Exception:
        pass
    try:
        return "Domain:{}".format(str(connector.Domain))
    except Exception:
        return "Electrical"


def _connector_is_primary(connector):
    try:
        return bool(connector.IsPrimary)
    except Exception:
        pass
    owner = getattr(connector, "Owner", None)
    parameter_id = getattr(DB.BuiltInParameter, "RBS_CONNECTOR_ISPRIMARY", None)
    if owner is not None and parameter_id is not None:
        try:
            parameter = owner.get_Parameter(parameter_id)
            if parameter is not None and parameter.HasValue:
                return parameter.AsInteger() == 1
        except Exception:
            pass
    return False


def _connector_sort_key(connector):
    primary_value = 0 if _connector_is_primary(connector) else 1
    connector_id = 0
    try:
        connector_id = int(connector.Id)
    except Exception:
        pass
    try:
        origin = connector.Origin
        return (
            primary_value,
            connector_id,
            round(origin.X, 6),
            round(origin.Y, 6),
            round(origin.Z, 6),
        )
    except Exception:
        return primary_value, connector_id


def electrical_connectors(element):
    manager = _connector_manager(element)
    if manager is None:
        return []
    connectors = []
    try:
        source_connectors = manager.Connectors
    except Exception:
        return connectors
    for connector in source_connectors:
        if _connector_is_electrical(connector):
            connectors.append(connector)
    return connectors


def connector_groups(element):
    grouped = {}
    for connector in electrical_connectors(element):
        key = connector_type_key(connector)
        grouped.setdefault(key, []).append(connector)
    for key in grouped:
        grouped[key] = sorted(grouped[key], key=_connector_sort_key)
    return grouped


def connector_type_name(key):
    """Return a stable, user-facing name for a connector type key."""
    value = str(key or "Electrical")
    prefix = "ElectricalSystemType:"
    if value.startswith(prefix):
        value = value[len(prefix):]
    return value


def system_type_choices(document, element_ids):
    """Return distinct connector types and the number of supporting devices."""
    counts = {}
    for element_id in list(element_ids or []):
        try:
            element = document.GetElement(element_id_from(element_id_value(element_id)))
            groups = connector_groups(element)
        except Exception:
            continue
        for key in groups.keys():
            counts.setdefault(key, set()).add(element_id_value(element.Id))

    keys = sorted(
        counts.keys(),
        key=lambda value: (-len(counts[value]), connector_type_name(value).lower(), value),
    )
    choices = []
    for key in keys:
        count = len(counts[key])
        suffix = "device" if count == 1 else "devices"
        display_name = "{} — {} {}".format(
            connector_type_name(key),
            count,
            suffix,
        )
        choices.append({
            "id": key,
            "name": display_name,
            "type_name": connector_type_name(key),
            "device_count": count,
        })
    return choices


def _default_connector_key(grouped):
    if not grouped:
        return None
    keys = sorted(grouped.keys())
    selected_key = keys[0]
    selected_primary = bool(
        grouped[selected_key]
        and _connector_is_primary(grouped[selected_key][0])
    )
    for key in keys[1:]:
        candidate = grouped[key][0]
        candidate_primary = _connector_is_primary(candidate)
        if candidate_primary and not selected_primary:
            selected_key = key
            selected_primary = True
    return selected_key


def resolve_connector(element, requested_key=None):
    """Resolve one connector without mixing system types or guessing ambiguities."""
    grouped = connector_groups(element)
    if not grouped:
        return None, requested_key, "No usable electrical connector was found."

    selected_key = requested_key
    if selected_key is None:
        selected_key = _default_connector_key(grouped)
    matching = grouped.get(selected_key, [])
    if not matching:
        return (
            None,
            selected_key,
            "No {} connector was found on this device.".format(
                connector_type_name(selected_key),
            ),
        )
    if len(matching) == 1:
        return matching[0], selected_key, None

    primary_matches = [
        connector for connector in matching if _connector_is_primary(connector)
    ]
    if len(primary_matches) == 1:
        return primary_matches[0], selected_key, None
    return (
        None,
        selected_key,
        "Device has multiple matching {} connectors and no unique primary "
        "connector can be identified.".format(connector_type_name(selected_key)),
    )


def primary_connector(element, requested_key=None):
    connector, connector_key, reason = resolve_connector(element, requested_key)
    del reason
    return connector, connector_key


def common_connector_key(elements, requested_key=None):
    if requested_key is not None:
        for element in list(elements or []):
            if requested_key not in connector_groups(element):
                return None
        return requested_key
    common_keys = None
    for element in list(elements or []):
        keys = set(connector_groups(element).keys())
        if not keys:
            return None
        common_keys = keys if common_keys is None else common_keys.intersection(keys)
    if not common_keys:
        return None

    sorted_keys = sorted(common_keys)
    selected_key = sorted_keys[0]
    for key in sorted_keys:
        primary_count = 0
        for element in list(elements or []):
            connector = connector_groups(element).get(key, [None])[0]
            if connector is not None and _connector_is_primary(connector):
                primary_count += 1
        selected_count = 0
        for element in list(elements or []):
            connector = connector_groups(element).get(selected_key, [None])[0]
            if connector is not None and _connector_is_primary(connector):
                selected_count += 1
        if primary_count > selected_count:
            selected_key = key
    return selected_key


def main_model_element(element):
    return design_options.is_main_model_element(element)


def is_annotation_element(element):
    """Return True when an element belongs to an annotation category."""
    if element is None:
        return False
    category = getattr(element, "Category", None)
    if category is None:
        return False
    category_type = getattr(category, "CategoryType", None)
    category_type_enum = getattr(DB, "CategoryType", None)
    annotation_type = getattr(category_type_enum, "Annotation", None)
    if annotation_type is not None and category_type == annotation_type:
        return True
    return False


def is_linked_element(element):
    """Return True for link instances or elements owned by a linked document."""
    if element is None:
        return False
    link_instance_class = getattr(DB, "RevitLinkInstance", None)
    if link_instance_class is not None and isinstance(element, link_instance_class):
        return True
    try:
        return bool(element.Document.IsLinked)
    except Exception:
        return False


def is_electrical_system_element(element):
    electrical_system_class = getattr(DB.Electrical, "ElectricalSystem", None)
    return (
        electrical_system_class is not None
        and element is not None
        and isinstance(element, electrical_system_class)
    )


def has_mep_model(element):
    """Return True when Revit exposes an MEP model for the element."""
    if element is None:
        return False
    try:
        return getattr(element, "MEPModel", None) is not None
    except Exception:
        return False


def is_allowed_device_pick(element, allow_circuit=False):
    """Apply the lightweight restrictions used while picking device elements."""
    if element is None:
        return False
    if is_annotation_element(element) or is_linked_element(element):
        return False
    if not main_model_element(element):
        return False
    element_type_class = getattr(DB, "ElementType", None)
    if element_type_class is not None and isinstance(element, element_type_class):
        return False
    if allow_circuit and not is_electrical_system_element(element):
        return has_mep_model(element)
    return True


def is_valid_device(element, allow_circuit=False):
    if element is None or not main_model_element(element):
        return False
    family_instance_class = getattr(DB, "FamilyInstance", None)
    if allow_circuit:
        element_type_class = getattr(DB, "ElementType", None)
        if element_type_class is not None and isinstance(element, element_type_class):
            return False
        # Wire by Circuit follows the legacy command: selection is broad and
        # circuit resolution happens afterward from the selected elements.
        # Do not require a family class, connector, or resolved system here.
        return True
    if family_instance_class is None or not isinstance(element, family_instance_class):
        return False
    return bool(electrical_connectors(element))


def is_valid_node(element):
    return is_valid_device(element, allow_circuit=False)


def circuit_eligible(circuit):
    return electrical_utils.is_circuit_eligible(circuit)


def _append_selection_step(selection_steps, stage_name, passed, message):
    if selection_steps is not None:
        selection_steps.append({
            "stage": str(stage_name),
            "passed": bool(passed),
            "message": str(message),
        })


def _element_systems(element, selection_steps=None):
    if element is None:
        _append_selection_step(
            selection_steps,
            "element lookup",
            False,
            "Element was None.",
        )
        return []
    selection_resolver = getattr(
        electrical_utils,
        "get_circuits_from_selection",
        None,
    )
    if selection_resolver is not None:
        try:
            resolved_systems = list(selection_resolver([element]) or [])
            if resolved_systems:
                _append_selection_step(
                    selection_steps,
                    "shared circuit resolver",
                    True,
                    "Resolved {} eligible circuit(s).".format(
                        len(resolved_systems)
                    ),
                )
                return resolved_systems
            _append_selection_step(
                selection_steps,
                "shared circuit resolver",
                False,
                "Returned zero eligible circuits.",
            )
        except Exception as error:
            _append_selection_step(
                selection_steps,
                "shared circuit resolver",
                False,
                "Raised {}. Falling back to direct MEPModel lookup.".format(error),
            )
    circuit_class = getattr(DB.Electrical, "ElectricalSystem", None)
    if circuit_class is not None and isinstance(element, circuit_class):
        eligible = circuit_eligible(element)
        _append_selection_step(
            selection_steps,
            "selected electrical system",
            eligible,
            "Selected element is an eligible power circuit."
            if eligible
            else "Selected electrical system is not an eligible power circuit.",
        )
        return [element] if eligible else []
    try:
        mep_model = getattr(element, "MEPModel", None)
    except Exception as error:
        _append_selection_step(
            selection_steps,
            "MEPModel lookup",
            False,
            "MEPModel property raised {}.".format(error),
        )
        return []
    if mep_model is None:
        _append_selection_step(
            selection_steps,
            "MEPModel lookup",
            False,
            "Element has no MEPModel.",
        )
        return []
    getter = getattr(mep_model, "GetElectricalSystems", None)
    if getter is not None:
        try:
            raw_systems = list(getter() or [])
            filtered_systems = electrical_utils.filter_circuits(raw_systems)
            _append_selection_step(
                selection_steps,
                "MEPModel.GetElectricalSystems",
                bool(filtered_systems),
                "Returned {} system(s); {} eligible power circuit(s) remained "
                "after filtering.".format(
                    len(raw_systems),
                    len(filtered_systems),
                ),
            )
            return filtered_systems
        except Exception as error:
            _append_selection_step(
                selection_steps,
                "MEPModel.GetElectricalSystems",
                False,
                "Raised {}. Trying legacy ElectricalSystems property.".format(error),
            )
    try:
        raw_systems = list(mep_model.ElectricalSystems or [])
        filtered_systems = electrical_utils.filter_circuits(raw_systems)
        _append_selection_step(
            selection_steps,
            "MEPModel.ElectricalSystems",
            bool(filtered_systems),
            "Returned {} system(s); {} eligible power circuit(s) remained after "
            "filtering.".format(
                len(raw_systems),
                len(filtered_systems),
            ),
        )
        return filtered_systems
    except Exception as error:
        _append_selection_step(
            selection_steps,
            "MEPModel.ElectricalSystems",
            False,
            "Raised {}.".format(error),
        )
        return []


def circuits_from_elements(document, elements):
    circuits = {}
    for element in list(elements or []):
        for circuit in get_element_circuits(element):
            circuits[element_id_value(circuit.Id)] = circuit
    return list(circuits.values())


def get_element_circuits(element, selection_steps=None):
    """Return eligible circuits associated with one selected host element."""
    return list(_element_systems(element, selection_steps=selection_steps) or [])


def _runtime_type_name(element):
    try:
        return str(element.GetType().FullName)
    except Exception:
        try:
            return str(type(element).__name__)
        except Exception:
            return "<Unknown API type>"


def _category_name(element):
    try:
        return str(element.Category.Name)
    except Exception:
        return "<No category>"


def _selection_outcome(detail, accepted, final_stage, reason, category=""):
    detail["accepted"] = bool(accepted)
    detail["final_stage"] = str(final_stage)
    detail["reason"] = str(reason)
    detail["category"] = str(category or "")
    return detail


def selection_validation_detail(
        document,
        element_id,
        scheme,
        requested_system_type=None):
    """Return the complete selection decision for one raw selection value."""
    detail = {
        "raw_id": str(element_id),
        "normalized_id": None,
        "element": None,
        "name": "<Unresolved>",
        "api_type": "<Unresolved>",
        "category_name": "<Unresolved>",
        "accepted": False,
        "final_stage": "input",
        "reason": "Selection was not evaluated.",
        "category": "",
        "steps": [],
        "resolution_steps": [],
        "circuit_ids": [],
        "circuit_types": [],
        "connector_count": 0,
        "connector_types": [],
        "system_type_key": requested_system_type,
    }

    try:
        normalized_id = element_id_value(element_id)
        detail["normalized_id"] = normalized_id
    except Exception as error:
        return _selection_outcome(
            detail,
            False,
            "ElementId normalization",
            "ElementId normalization raised {}.".format(error),
            "id_normalization",
        )

    _append_selection_step(
        detail["steps"],
        "ElementId normalization",
        normalized_id > 0,
        "Normalized raw value {} to ElementId {}.".format(
            detail["raw_id"],
            normalized_id,
        ),
    )
    if normalized_id <= 0:
        return _selection_outcome(
            detail,
            False,
            "ElementId normalization",
            "The selected value normalized to ElementId {}.".format(normalized_id),
            "id_normalization",
        )

    try:
        element = document.GetElement(element_id_from(normalized_id))
    except Exception as error:
        return _selection_outcome(
            detail,
            False,
            "Document lookup",
            "document.GetElement failed for ElementId {}: {}.".format(
                normalized_id,
                error,
            ),
            "document_lookup",
        )
    if element is None:
        return _selection_outcome(
            detail,
            False,
            "Document lookup",
            "document.GetElement returned None for ElementId {}.".format(
                normalized_id,
            ),
            "document_lookup",
        )

    detail["element"] = element
    detail["name"] = safe_element_name(element)
    detail["api_type"] = _runtime_type_name(element)
    detail["category_name"] = _category_name(element)
    _append_selection_step(
        detail["steps"],
        "Document lookup",
        True,
        "Resolved {} ({}) in category {}.".format(
            detail["name"],
            detail["api_type"],
            detail["category_name"],
        ),
    )

    try:
        main_model_passed = bool(main_model_element(element))
    except Exception as error:
        return _selection_outcome(
            detail,
            False,
            "Main-model check",
            "Main-model/design-option check raised {}.".format(error),
            "main_model",
        )
    _append_selection_step(
        detail["steps"],
        "Main-model check",
        main_model_passed,
        "Element is in the main model."
        if main_model_passed
        else "Element is in a design option or could not be confirmed in the main model.",
    )
    if not main_model_passed:
        return _selection_outcome(
            detail,
            False,
            "Main-model check",
            "Element is not in the main model or its design option could not be confirmed.",
            "main_model",
        )

    element_type_class = getattr(DB, "ElementType", None)
    if element_type_class is not None and isinstance(element, element_type_class):
        _append_selection_step(
            detail["steps"],
            "Element instance check",
            False,
            "Element is an ElementType; model instances are required.",
        )
        return _selection_outcome(
            detail,
            False,
            "Element instance check",
            "Element types cannot be wired; select model instances instead.",
            "element_type",
        )
    _append_selection_step(
        detail["steps"],
        "Element instance check",
        True,
        "Element is a model instance rather than an element type.",
    )

    if scheme == SCHEME_WIRE_BY_CIRCUIT:
        circuit_list = get_element_circuits(
            element,
            selection_steps=detail["resolution_steps"],
        )
        detail["circuit_ids"] = [
            element_id_value(circuit.Id) for circuit in circuit_list
        ]
        detail["circuit_types"] = [
            str(getattr(circuit, "SystemType", "<Unavailable>"))
            for circuit in circuit_list
        ]
        _append_selection_step(
            detail["steps"],
            "Circuit requirement",
            bool(circuit_list),
            "Resolved {} eligible power circuit(s): {}.".format(
                len(circuit_list),
                ", ".join([str(value) for value in detail["circuit_ids"]])
                if circuit_list
                else "none",
            ),
        )
        if not circuit_list:
            return _selection_outcome(
                detail,
                False,
                "Circuit requirement",
                "No eligible power circuit could be resolved from this selected element.",
                "no_circuit",
            )
        return _selection_outcome(
            detail,
            True,
            "Circuit requirement",
            "Accepted for Wire by Circuit.",
        )

    family_instance_class = getattr(DB, "FamilyInstance", None)
    family_instance_passed = (
        family_instance_class is not None
        and isinstance(element, family_instance_class)
    )
    _append_selection_step(
        detail["steps"],
        "FamilyInstance requirement",
        family_instance_passed,
        "Element is a FamilyInstance."
        if family_instance_passed
        else "This scheme only accepts FamilyInstance elements.",
    )
    if not family_instance_passed:
        return _selection_outcome(
            detail,
            False,
            "FamilyInstance requirement",
            "This wiring scheme requires a FamilyInstance host.",
            "family_instance",
        )

    try:
        connector_list = electrical_connectors(element)
        detail["connector_count"] = len(connector_list)
        detail["connector_types"] = sorted(set([
            connector_type_key(connector)
            for connector in connector_list
        ]))
    except Exception as error:
        return _selection_outcome(
            detail,
            False,
            "Electrical connector lookup",
            "Electrical connector lookup raised {}.".format(error),
            "connector_lookup",
        )
    _append_selection_step(
        detail["steps"],
        "Electrical connector requirement",
        detail["connector_count"] > 0,
        "Found {} usable electrical connector(s) in type group(s): {}.".format(
            detail["connector_count"],
            ", ".join(detail["connector_types"])
            if detail["connector_types"]
            else "none",
        ),
    )
    if detail["connector_count"] <= 0:
        return _selection_outcome(
            detail,
            False,
            "Electrical connector requirement",
            "FamilyInstance has no usable electrical connector.",
            "connector_requirement",
        )
    if requested_system_type is not None:
        connector, resolved_key, resolution_reason = resolve_connector(
            element,
            requested_system_type,
        )
        detail["system_type_key"] = resolved_key
        if connector is None:
            type_name = connector_type_name(requested_system_type)
            if not connector_groups(element).get(requested_system_type):
                reason = "This device does not have a matching {} connector.".format(
                    type_name,
                )
                category = "system_type_mismatch"
            else:
                reason = resolution_reason or (
                    "This device has multiple matching {} connectors.".format(
                        type_name,
                    )
                )
                category = "ambiguous_connector"
            _append_selection_step(
                detail["steps"],
                "Selected system type connector resolution",
                False,
                reason,
            )
            return _selection_outcome(
                detail,
                False,
                "Selected system type connector resolution",
                reason,
                category,
            )
        _append_selection_step(
            detail["steps"],
            "Selected system type connector resolution",
            True,
            "Resolved one {} connector for this device.".format(
                connector_type_name(requested_system_type),
            ),
        )
    return _selection_outcome(
        detail,
        True,
        "Electrical connector requirement",
        "Accepted for {}.".format(SCHEME_LABELS.get(scheme, scheme)),
    )


def circuit_member_count(circuit):
    try:
        members = list(circuit.Elements or [])
    except Exception:
        return 0
    count = 0
    for member in members:
        if electrical_connectors(member):
            count += 1
    return count


def _view_wires_for_circuits(document, view_id, circuits):
    circuit_values = set([element_id_value(circuit.Id) for circuit in circuits])
    wire_map = collect_active_view_wires_by_circuit(document, view_id)
    wires = []
    seen_values = set()
    for circuit_value in circuit_values:
        for wire in wire_map.get(circuit_value, []):
            wire_value = element_id_value(wire.Id)
            if wire_value not in seen_values:
                seen_values.add(wire_value)
                wires.append(wire)
    return wires


def _delete_wires(document, wires):
    deleted = 0
    failures = []
    for wire in list(wires or []):
        try:
            document.Delete(wire.Id)
            deleted += 1
        except Exception as error:
            failures.append({
                "id": element_id_value(wire.Id),
                "element": wire,
                "reason": "Existing wire could not be deleted: {}".format(error),
            })
    return deleted, failures


def _wires_connected_to_connector(connector, view, homeruns_only=False):
    wires = []
    try:
        references = connector.AllRefs
    except Exception:
        return wires
    for reference in references:
        owner = getattr(reference, "Owner", None)
        if owner is None or not isinstance(owner, DB.Electrical.Wire):
            continue
        try:
            owner_view_id = owner.OwnerViewId
            if element_id_value(owner_view_id) != element_id_value(view.Id):
                continue
        except Exception:
            continue
        if homeruns_only:
            try:
                if not is_homerun_wire(owner):
                    continue
            except Exception as error:
                script.get_logger().warning(
                    "Could not classify connected wire {} as a homerun: {}".format(
                        element_id_value(owner.Id),
                        error,
                    )
                )
                continue
        wires.append(owner)
    return wires


def _delete_wires_connected_to_records(document, view, connector_records,
                                       homeruns_only=False):
    wires = []
    seen_values = set()
    for element, connector, connector_key in list(connector_records or []):
        del element
        del connector_key
        for wire in _wires_connected_to_connector(
                connector,
                view,
                homeruns_only=homeruns_only):
            wire_value = element_id_value(wire.Id)
            if wire_value in seen_values:
                continue
            seen_values.add(wire_value)
            wires.append(wire)
    return _delete_wires(document, wires)


def _view_normal(view):
    try:
        normal = view.ViewDirection
        if normal.GetLength() > GEOMETRY_TOLERANCE:
            return normal.Normalize()
    except Exception:
        pass
    return DB.XYZ.BasisZ


def _project_direction(vector, view):
    if vector is None:
        return None
    try:
        normal = _view_normal(view)
        projected = vector.Subtract(normal.Multiply(vector.DotProduct(normal)))
        if projected.GetLength() > GEOMETRY_TOLERANCE:
            return projected.Normalize()
    except Exception:
        pass
    return None


def _direction_from_last_device(element, connector, view, fallback_end=None):
    candidates = []
    try:
        candidates.append(element.FacingOrientation)
    except Exception:
        pass
    try:
        candidates.append(element.HandOrientation)
    except Exception:
        pass
    try:
        coordinate_system = connector.CoordinateSystem
        candidates.extend([
            coordinate_system.BasisX,
            coordinate_system.BasisY,
            coordinate_system.BasisZ,
        ])
    except Exception:
        pass
    if fallback_end is not None:
        try:
            candidates.append(fallback_end.Subtract(connector.Origin))
        except Exception:
            pass
    candidates.extend([DB.XYZ.BasisX, DB.XYZ.BasisY])
    for candidate in candidates:
        projected = _project_direction(candidate, view)
        if projected is not None:
            return projected
    return DB.XYZ.BasisX


def _panel_connector_for_element(element, connector_key):
    for system in _element_systems(element):
        base_equipment = getattr(system, "BaseEquipment", None)
        if base_equipment is None:
            continue
        candidate, candidate_key = primary_connector(
            base_equipment,
            connector_key,
        )
        if candidate is not None and candidate_key == connector_key:
            return candidate
    return None


def _panel_direction(element, connector, view, connector_key,
                     fallback_end=None):
    panel_connector = _panel_connector_for_element(element, connector_key)
    if panel_connector is not None:
        projected = _project_direction(
            panel_connector.Origin.Subtract(connector.Origin),
            view,
        )
        if projected is not None:
            return projected
    if fallback_end is not None:
        projected = _project_direction(
            fallback_end.Subtract(connector.Origin),
            view,
        )
        if projected is not None:
            return projected
    return _direction_from_last_device(element, connector, view)


def _homerun_end_point(element, connector, view, length, direction_mode,
                       connector_key, fallback_end=None):
    if direction_mode == HOMERUN_DIRECTION_PANEL:
        direction = _panel_direction(
            element,
            connector,
            view,
            connector_key,
            fallback_end=fallback_end,
        )
    else:
        direction = _direction_from_last_device(
            element,
            connector,
            view,
            fallback_end=fallback_end,
        )
    parsed_length = float(length or HOME_RUN_LENGTH)
    if parsed_length <= GEOMETRY_TOLERANCE:
        parsed_length = HOME_RUN_LENGTH
    return connector.Origin.Add(direction.Multiply(parsed_length))


def _create_wire(document, view, wire_type_id, wiring_type, start_connector,
                 end_connector=None, end_point=None):
    start_point = start_connector.Origin if start_connector is not None else None
    if end_point is None and end_connector is not None:
        end_point = end_connector.Origin
    if start_point is None or end_point is None:
        raise ValueError("Wire endpoints could not be resolved.")
    if start_point.DistanceTo(end_point) <= GEOMETRY_TOLERANCE:
        raise ValueError("Wire endpoints are coincident.")
    midpoint = start_point.Add(end_point.Subtract(start_point).Multiply(0.5))
    points = [start_point, midpoint, end_point]
    return DB.Electrical.Wire.Create(
        document,
        wire_type_id,
        view.Id,
        wiring_type,
        points,
        start_connector,
        end_connector,
    )


def _create_wire_from_points(document, view, wire_type_id, wiring_type,
                             start_connector, points, end_connector=None):
    usable_points = list(points or [])
    if len(usable_points) < 2:
        raise ValueError("At least two wire points are required.")
    if usable_points[0].DistanceTo(usable_points[-1]) <= GEOMETRY_TOLERANCE:
        raise ValueError("Wire endpoints are coincident.")
    if len(usable_points) == 2:
        midpoint = usable_points[0].Add(
            usable_points[1].Subtract(usable_points[0]).Multiply(0.5)
        )
        usable_points.insert(1, midpoint)
    return DB.Electrical.Wire.Create(
        document,
        wire_type_id,
        view.Id,
        wiring_type,
        usable_points,
        start_connector,
        end_connector,
    )


def _vector_length(vector):
    try:
        return vector.GetLength()
    except Exception:
        return 0.0


def _clamp_value(value, minimum, maximum):
    return max(minimum, min(maximum, value))


def _homerun_points(start, end_point, shape, bend_offset,
                     native_vertex=None, previous_connector=None):
    requested_offset = None
    try:
        if bend_offset is not None:
            requested_offset = float(bend_offset)
    except Exception:
        requested_offset = None

    # A zero offset is the straight-path setting.  Keep accepting the shape
    # argument for compatibility with older payloads, but the UI now uses the
    # offset itself to select straight versus bent geometry.
    if (shape == HOMERUN_SHAPE_STRAIGHT
            or (requested_offset is not None
                and abs(requested_offset) <= GEOMETRY_TOLERANCE)):
        midpoint = start.Add(end_point.Subtract(start).Multiply(0.5))
        return [start, midpoint, end_point]

    segment = end_point.Subtract(start)
    segment_length = _vector_length(segment)
    if segment_length <= GEOMETRY_TOLERANCE:
        raise ValueError("Homerun endpoints are coincident.")
    direction = segment.Normalize()
    previous_point = (
        previous_connector.Origin
        if previous_connector is not None
        else None
    )
    unused_along, perpendicular_direction = build_local_frame(
        start,
        previous_point=previous_point,
        fallback_end_point=end_point,
    )
    del unused_along

    perpendicular_length = (
        abs(requested_offset) if requested_offset is not None else 0.0
    )
    if perpendicular_length <= GEOMETRY_TOLERANCE:
        perpendicular_length = max(segment_length * 0.2, 0.15)
    if native_vertex is not None:
        vertex_vector = native_vertex.Subtract(start)
        along_length = vertex_vector.DotProduct(direction)
        projected = direction.Multiply(along_length)
        perpendicular_vector = vertex_vector.Subtract(projected)
        candidate_length = _vector_length(perpendicular_vector)
        if candidate_length > GEOMETRY_TOLERANCE:
            perpendicular_direction = perpendicular_vector.Normalize()
            if requested_offset is None or abs(requested_offset) <= GEOMETRY_TOLERANCE:
                perpendicular_length = candidate_length

    # Preserve the native bend side for positive values and mirror it for a
    # negative value.  The sign must be applied after the native vertex has
    # established the base perpendicular direction.
    if requested_offset is not None and requested_offset < -GEOMETRY_TOLERANCE:
        perpendicular_direction = perpendicular_direction.Multiply(-1.0)

    perpendicular_length = max(
        perpendicular_length,
        max(segment_length * 0.08, 0.05),
    )
    vertex_base = start.Add(direction.Multiply(segment_length * 0.5))
    vertex = vertex_base.Add(
        perpendicular_direction.Multiply(perpendicular_length)
    )
    if (are_points_coincident(start, vertex)
            or are_points_coincident(vertex, end_point)):
        vertex = vertex_base.Add(
            perpendicular_direction.Multiply(
                max(segment_length * 0.08, 0.05)
            )
        )
    return [start, vertex, end_point]


def _replace_homerun_custom(document, view, wire, wire_type_id, wiring_type,
                            homerun_length, direction_mode, shape,
                            bend_offset):
    connected, unconnected = wire_connected_unconnected_connectors(wire)
    if len(connected) != 1 or len(unconnected) != 1:
        raise ValueError("The generated homerun does not have one open endpoint.")
    start_connector = get_element_connector_from_wire_connector(connected[0])
    if start_connector is None:
        raise ValueError("The homerun device connector could not be resolved.")
    owner = getattr(start_connector, "Owner", None)
    if owner is None:
        raise ValueError("The homerun device could not be resolved.")
    native_end = unconnected[0].Origin
    previous_connector = find_previous_connector_for_homerun(
        start_connector,
        homerun_wire_id=wire.Id,
    )
    native_vertex = None
    try:
        if getattr(wire, "NumberOfVertices", 0) > 0:
            native_vertex = wire.GetVertex(0)
    except Exception as error:
        script.get_logger().warning(
            "Could not read the native homerun vertex; using a calculated vertex: {}".format(
                error
            )
        )
    end_point = _homerun_end_point(
        owner,
        start_connector,
        view,
        homerun_length,
        direction_mode,
        connector_type_key(start_connector),
        fallback_end=native_end,
    )
    homerun_points = _homerun_points(
        start_connector.Origin,
        end_point,
        shape,
        bend_offset,
        native_vertex=native_vertex,
        previous_connector=previous_connector,
    )
    wire_type_value = get_wire_type_id(wire) or wire_type_id
    document.Delete(wire.Id)
    return DB.Electrical.Wire.Create(
        document,
        wire_type_value,
        view.Id,
        wiring_type,
        homerun_points,
        start_connector,
        None,
    )


def _homerun_from_wire_set(wire_set):
    for wire in list(wire_set or []):
        if is_homerun_wire(wire):
            return wire
    return None


def _apply_wire_type(wire_set, wire_type_id):
    if wire_type_id is None:
        return
    for wire in list(wire_set or []):
        wire.ChangeTypeId(wire_type_id)


def run_wire_by_circuit(document, view, elements, settings):
    if not settings.get("wire_type_id"):
        raise ValueError("Select a wire type before creating wires.")
    circuits = circuits_from_elements(document, elements)
    if not circuits:
        raise ValueError("No eligible electrical circuits were found from the selected devices.")
    skipped_circuits = []
    if settings.get("skip_single_device"):
        eligible_circuits = []
        for circuit in circuits:
            member_count = circuit_member_count(circuit)
            if member_count <= 1:
                skipped_circuits.append({
                    "id": element_id_value(circuit.Id),
                    "element": circuit,
                    "reason": "Single-device circuit skipped by settings.",
                })
            else:
                eligible_circuits.append(circuit)
        circuits = eligible_circuits
    if not circuits:
        return {
            "created": 0,
            "homeruns": [],
            "deleted": 0,
            "skipped": skipped_circuits,
            "failures": [],
            "scheme": SCHEME_LABELS[SCHEME_WIRE_BY_CIRCUIT],
        }

    wire_type_id = element_id_from(settings["wire_type_id"])
    branch_type = wiring_type_from_name(
        settings.get("branch_wiring_type"),
        "Chamfer",
    )
    homerun_type = wiring_type_from_name(
        settings.get("homerun_wiring_type"),
        "Arc",
    )
    transaction = DB.Transaction(document, "Wire Tools - Wire by Circuit")
    transaction.Start()
    created_count = 0
    homerun_ids = []
    deleted_count = 0
    failures = []
    try:
        if settings.get("redraw_existing_wires", True):
            existing_wires = _view_wires_for_circuits(document, view.Id, circuits)
            deleted_count, deletion_failures = _delete_wires(
                document,
                existing_wires,
            )
            failures.extend(deletion_failures)
        for circuit in circuits:
            subtransaction = DB.SubTransaction(document)
            subtransaction.Start()
            try:
                wire_set = circuit.NewWires(view, branch_type)
                if not wire_set:
                    raise ValueError("ElectricalSystem.NewWires returned no wires.")
                _apply_wire_type(wire_set, wire_type_id)
                generated_homerun = _homerun_from_wire_set(wire_set)
                if generated_homerun is not None:
                    if (settings.get("homerun_direction", HOMERUN_DIRECTION_PANEL)
                            != HOMERUN_DIRECTION_PANEL
                            or settings.get("homerun_shape", HOMERUN_SHAPE_STRAIGHT)
                            != HOMERUN_SHAPE_STRAIGHT):
                        generated_homerun = _replace_homerun_custom(
                            document,
                            view,
                            generated_homerun,
                            wire_type_id,
                            homerun_type,
                            settings.get("homerun_length", HOME_RUN_LENGTH),
                            settings.get(
                                "homerun_direction",
                                HOMERUN_DIRECTION_PANEL,
                            ),
                            settings.get(
                                "homerun_shape",
                                HOMERUN_SHAPE_STRAIGHT,
                            ),
                            settings.get("bend_offset", 1.0),
                        )
                    homerun_ids.append(element_id_value(generated_homerun.Id))
                created_count += len(list(wire_set))
                subtransaction.Commit()
            except Exception as error:
                subtransaction.RollBack()
                failures.append({
                    "id": element_id_value(circuit.Id),
                    "element": circuit,
                    "reason": "Circuit wiring failed: {}".format(error),
                })
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    return {
        "created": created_count,
        "homeruns": homerun_ids,
        "deleted": deleted_count,
        "skipped": skipped_circuits,
        "failures": failures,
        "scheme": SCHEME_LABELS[SCHEME_WIRE_BY_CIRCUIT],
    }


def _resolve_device_connector(element, requested_key=None):
    connector, connector_key, reason = resolve_connector(element, requested_key)
    if connector is None:
        raise ValueError(reason or "No usable electrical connector was found.")
    return connector, connector_key


def _run_direct_wires(document, view, device_elements, settings,
                       node_element=None, individual_homeruns=False):
    if not device_elements:
        raise ValueError("No valid devices were selected.")
    if not settings.get("wire_type_id"):
        raise ValueError("Select a wire type before creating wires.")
    connector_key = settings.get("system_type_key")
    if connector_key is None and not individual_homeruns:
        connector_key = common_connector_key(device_elements)
        if connector_key is None:
            raise ValueError("Selected devices do not share a connector type.")
    if node_element is not None:
        node_connector, node_key = _resolve_device_connector(
            node_element,
            connector_key,
        )
        if connector_key is None:
            connector_key = node_key

    wire_type_id = element_id_from(settings["wire_type_id"])
    branch_type = wiring_type_from_name(
        settings.get("branch_wiring_type"),
        "Chamfer",
    )
    homerun_type = wiring_type_from_name(
        settings.get("homerun_wiring_type"),
        "Arc",
    )
    transaction_name = "Wire Tools - Custom wiring"
    transaction = DB.Transaction(document, transaction_name)
    transaction.Start()
    created_count = 0
    homerun_ids = []
    failures = []
    deleted_count = 0
    try:
        connector_records = []
        for element in device_elements:
            try:
                connector, resolved_key = _resolve_device_connector(
                    element,
                    connector_key,
                )
                if connector_key is None:
                    connector_key = resolved_key
                connector_records.append((element, connector, resolved_key))
            except Exception as error:
                failures.append({
                    "id": element_id_value(element.Id),
                    "element": element,
                    "reason": str(error),
                })

        if settings.get("redraw_existing_wires", True):
            deleted_count, deletion_failures = _delete_wires_connected_to_records(
                document,
                view,
                connector_records,
                # Redraw is intentionally scheme-independent.  A device may
                # currently be connected by a wire created by another scheme,
                # and leaving that wire in place is what produced duplicate
                # homeruns on repeated Individual Homeruns runs.
                homeruns_only=False,
            )
            failures.extend(deletion_failures)

        if node_element is not None:
            if settings.get("redraw_existing_wires", True):
                node_homerun_deleted, node_homerun_failures = (
                    _delete_wires_connected_to_records(
                        document,
                        view,
                        [(node_element, node_connector, node_key)],
                        homeruns_only=True,
                    )
                )
                deleted_count += node_homerun_deleted
                failures.extend(node_homerun_failures)
            for element, connector, resolved_key in connector_records:
                subtransaction = DB.SubTransaction(document)
                subtransaction.Start()
                try:
                    if resolved_key != connector_key:
                        raise ValueError("Device connector type does not match the node connector.")
                    created_wire = _create_wire(
                        document,
                        view,
                        wire_type_id,
                        branch_type,
                        connector,
                        end_connector=node_connector,
                    )
                    created_count += 1
                    subtransaction.Commit()
                except Exception as error:
                    subtransaction.RollBack()
                    failures.append({
                        "id": element_id_value(element.Id),
                        "element": element,
                        "reason": "Wire to node failed: {}".format(error),
                    })
            subtransaction = DB.SubTransaction(document)
            subtransaction.Start()
            try:
                created_homerun = _custom_interconnect_homerun(
                    document,
                    view,
                    wire_type_id,
                    node_element,
                    node_connector,
                    settings,
                )
                created_count += 1
                if created_homerun is not None:
                    homerun_ids.append(element_id_value(created_homerun.Id))
                subtransaction.Commit()
            except Exception as error:
                subtransaction.RollBack()
                failures.append({
                    "id": element_id_value(node_element.Id),
                    "element": node_element,
                    "reason": "Node homerun failed: {}".format(error),
                })
            transaction.Commit()
            return created_count, homerun_ids, failures, deleted_count

        if individual_homeruns:
            for element, connector, resolved_key in connector_records:
                subtransaction = DB.SubTransaction(document)
                subtransaction.Start()
                try:
                    end_point = _homerun_end_point(
                        element,
                        connector,
                        view,
                        settings.get("homerun_length", HOME_RUN_LENGTH),
                        settings.get(
                            "homerun_direction",
                            HOMERUN_DIRECTION_PANEL,
                        ),
                        resolved_key,
                    )
                    homerun_points = _homerun_points(
                        connector.Origin,
                        end_point,
                        settings.get(
                            "homerun_shape",
                            HOMERUN_SHAPE_STRAIGHT,
                        ),
                        settings.get("bend_offset", 1.0),
                    )
                    created_wire = _create_wire_from_points(
                        document,
                        view,
                        wire_type_id,
                        homerun_type,
                        connector,
                        homerun_points,
                    )
                    created_count += 1
                    if created_wire is not None:
                        homerun_ids.append(element_id_value(created_wire.Id))
                    subtransaction.Commit()
                except Exception as error:
                    subtransaction.RollBack()
                    failures.append({
                        "id": element_id_value(element.Id),
                        "element": element,
                        "reason": "Individual homerun failed: {}".format(error),
                    })
        else:
            for index in range(len(connector_records) - 1):
                first_element, first_connector, first_key = connector_records[index]
                second_element, second_connector, second_key = connector_records[index + 1]
                subtransaction = DB.SubTransaction(document)
                subtransaction.Start()
                try:
                    if first_key != second_key:
                        raise ValueError("Adjacent devices do not share a connector type.")
                    _create_wire(
                        document,
                        view,
                        wire_type_id,
                        branch_type,
                        first_connector,
                        end_connector=second_connector,
                    )
                    created_count += 1
                    subtransaction.Commit()
                except Exception as error:
                    subtransaction.RollBack()
                    failures.append({
                        "id": element_id_value(first_element.Id),
                        "element": first_element,
                        "reason": "Interconnect failed to {}: {}".format(
                            element_id_value(second_element.Id),
                            error,
                        ),
                    })
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    return created_count, homerun_ids, failures, deleted_count


def _connector_distance(first_connector, second_connector):
    try:
        return first_connector.Origin.DistanceTo(second_connector.Origin)
    except Exception:
        return float("inf")


def _spatial_element_order(records):
    remaining = list(records or [])
    ordered = []
    if not remaining:
        return ordered
    remaining.sort(
        key=lambda item: (
            item[1].Origin.X,
            item[1].Origin.Y,
            item[1].Origin.Z,
        )
    )
    ordered.append(remaining.pop(0))
    while remaining:
        current = ordered[-1]
        nearest_index = min(
            range(len(remaining)),
            key=lambda index: _connector_distance(
                current[1],
                remaining[index][1],
            ),
        )
        ordered.append(remaining.pop(nearest_index))
    return ordered


def _point_segment_distance_xy(point, start_point, end_point):
    delta_x = end_point.X - start_point.X
    delta_y = end_point.Y - start_point.Y
    length_squared = delta_x * delta_x + delta_y * delta_y
    if length_squared <= GEOMETRY_TOLERANCE:
        return math.sqrt(
            (point.X - start_point.X) ** 2
            + (point.Y - start_point.Y) ** 2
        )
    projection = (
        (point.X - start_point.X) * delta_x
        + (point.Y - start_point.Y) * delta_y
    ) / length_squared
    projection = max(0.0, min(1.0, projection))
    closest_x = start_point.X + delta_x * projection
    closest_y = start_point.Y + delta_y * projection
    return math.sqrt(
        (point.X - closest_x) ** 2
        + (point.Y - closest_y) ** 2
    )


def _spatial_group_edges(group_records):
    group_keys = list(group_records.keys())
    all_records = []
    for records in group_records.values():
        all_records.extend(records)
    edges = []
    for first_index in range(len(group_keys)):
        for second_index in range(first_index + 1, len(group_keys)):
            first_key = group_keys[first_index]
            second_key = group_keys[second_index]
            best_edge = None
            for first_record in group_records[first_key]:
                for second_record in group_records[second_key]:
                    distance = _connector_distance(
                        first_record[1],
                        second_record[1],
                    )
                    start_point = first_record[1].Origin
                    end_point = second_record[1].Origin
                    blocker_count = 0
                    blocker_distance = max(
                        0.5,
                        min(1.5, distance * 0.15),
                    )
                    for blocker_record in all_records:
                        if (blocker_record is first_record
                                or blocker_record is second_record):
                            continue
                        if _point_segment_distance_xy(
                                blocker_record[1].Origin,
                                start_point,
                                end_point) < blocker_distance:
                            blocker_count += 1
                    score = distance + blocker_count * max(
                        10.0,
                        distance * 3.0,
                    )
                    if best_edge is None or score < best_edge[0]:
                        best_edge = (
                            score,
                            first_key,
                            second_key,
                            first_record,
                            second_record,
                        )
            if best_edge is not None:
                edges.append(best_edge)
    edges.sort(key=lambda edge: edge[0])
    return edges


def _spatial_group_mst(group_records):
    parent = {}
    for group_key in group_records.keys():
        parent[group_key] = group_key

    def find_group(group_key):
        while parent[group_key] != group_key:
            parent[group_key] = parent[parent[group_key]]
            group_key = parent[group_key]
        return group_key

    selected_edges = []
    for edge in _spatial_group_edges(group_records):
        first_key = edge[1]
        second_key = edge[2]
        first_root = find_group(first_key)
        second_root = find_group(second_key)
        if first_root == second_root:
            continue
        parent[first_root] = second_root
        selected_edges.append(edge)
        if len(selected_edges) >= max(len(group_records) - 1, 0):
            break
    return selected_edges


def _run_spatial_interconnect(document, view, device_elements, settings):
    connector_key = settings.get("system_type_key")
    if connector_key is None:
        connector_key = common_connector_key(device_elements)
    if connector_key is None:
        raise ValueError("Selected devices do not share a connector type.")

    connector_records = []
    failures = []
    for element in list(device_elements or []):
        try:
            connector, resolved_key = _resolve_device_connector(
                element,
                connector_key,
            )
        except Exception as error:
            failures.append({
                "id": element_id_value(element.Id),
                "element": element,
                "reason": "Connector resolution failed: {}".format(error),
            })
            continue
        if connector is None or resolved_key != connector_key:
            failures.append({
                "id": element_id_value(element.Id),
                "element": element,
                "reason": "No connector matching the common interconnect type.",
            })
            continue
        connector_records.append((element, connector, resolved_key))

    if not connector_records:
        return 0, [], failures, 0

    ordered_records = _spatial_element_order(connector_records)
    pair_records = [
        (ordered_records[index], ordered_records[index + 1])
        for index in range(len(ordered_records) - 1)
    ]

    wire_type_id = element_id_from(settings["wire_type_id"])
    branch_type = wiring_type_from_name(
        settings.get("branch_wiring_type"),
        "Chamfer",
    )
    transaction = DB.Transaction(document, "Wire Tools - Spatial Interconnect")
    transaction.Start()
    created_count = 0
    homerun_ids = []
    deleted_count = 0
    try:
        if settings.get("redraw_existing_wires", True):
            deleted_count, deletion_failures = _delete_wires_connected_to_records(
                document,
                view,
                connector_records,
            )
            failures.extend(deletion_failures)
        for first_record, second_record in pair_records:
            first_element, first_connector, first_key = first_record
            second_element, second_connector, second_key = second_record
            subtransaction = DB.SubTransaction(document)
            subtransaction.Start()
            try:
                if first_key != second_key:
                    raise ValueError("Spatially selected connectors have different types.")
                points = _native_interconnect_points(
                    first_connector.Origin,
                    second_connector.Origin,
                    branch_type,
                    settings.get("bend_offset", 1.0),
                )
                _create_wire_from_points(
                    document,
                    view,
                    wire_type_id,
                    branch_type,
                    first_connector,
                    points,
                    end_connector=second_connector,
                )
                created_count += 1
                subtransaction.Commit()
            except Exception as error:
                subtransaction.RollBack()
                failures.append({
                    "id": element_id_value(first_element.Id),
                    "element": first_element,
                    "reason": "Spatial interconnect to {} failed: {}".format(
                        element_id_value(second_element.Id),
                        error,
                    ),
                })

        homerun_element, homerun_connector, homerun_key = ordered_records[0]
        del homerun_key
        subtransaction = DB.SubTransaction(document)
        subtransaction.Start()
        try:
            created_homerun = _custom_interconnect_homerun(
                document,
                view,
                wire_type_id,
                homerun_element,
                homerun_connector,
                settings,
            )
            created_count += 1
            homerun_ids.append(element_id_value(created_homerun.Id))
            subtransaction.Commit()
        except Exception as error:
            subtransaction.RollBack()
            failures.append({
                "id": element_id_value(homerun_element.Id),
                "element": homerun_element,
                "reason": "Selected-device homerun failed: {}".format(error),
            })
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    return created_count, homerun_ids, failures, deleted_count


def _wire_vertex_offset(wire):
    try:
        if int(getattr(wire, "NumberOfVertices", 0) or 0) <= 0:
            return 0.0
        connected, unconnected = wire_connected_unconnected_connectors(wire)
        if len(connected) != 1 or len(unconnected) != 1:
            return 0.0
        start_connector = get_element_connector_from_wire_connector(connected[0])
        if start_connector is None:
            return 0.0
        start_point = start_connector.Origin
        end_point = unconnected[0].Origin
        direction_vector = end_point.Subtract(start_point)
        if direction_vector.GetLength() <= GEOMETRY_TOLERANCE:
            return 0.0
        direction = direction_vector.Normalize()
        vertex = wire.GetVertex(0)
        vertex_vector = vertex.Subtract(start_point)
        along_length = vertex_vector.DotProduct(direction)
        projected = direction.Multiply(along_length)
        perpendicular = vertex_vector.Subtract(projected)
        return _vector_length(perpendicular)
    except Exception:
        return 0.0


def _native_homerun_record(circuit, wire):
    connected, unconnected = wire_connected_unconnected_connectors(wire)
    if len(connected) != 1 or len(unconnected) != 1:
        return None
    device_connector = get_element_connector_from_wire_connector(connected[0])
    if device_connector is None:
        return None
    owner = getattr(device_connector, "Owner", None)
    if owner is None:
        return None
    connector_key = connector_type_key(device_connector)
    panel_connector = _panel_connector_for_element(owner, connector_key)
    if panel_connector is not None:
        panel_distance = unconnected[0].Origin.DistanceTo(panel_connector.Origin)
    else:
        panel_distance = unconnected[0].Origin.DistanceTo(device_connector.Origin)
    return {
        "circuit": circuit,
        "wire": wire,
        "device_connector": device_connector,
        "open_connector": unconnected[0],
        "owner": owner,
        "connector_key": connector_key,
        "panel_distance": panel_distance,
    }


def _circuit_connector_records(circuit, connector_key):
    """Return all usable member connectors for spatial circuit bridging."""
    records = []
    seen_values = set()
    base_equipment = getattr(circuit, "BaseEquipment", None)
    base_value = element_id_value(getattr(base_equipment, "Id", None))
    try:
        members = list(circuit.Elements or [])
    except Exception:
        members = []
    for member in members:
        member_value = element_id_value(getattr(member, "Id", None))
        if member_value == base_value or member_value in seen_values:
            continue
        if not main_model_element(member):
            continue
        try:
            connector, resolved_key = _resolve_device_connector(
                member,
                connector_key,
            )
        except Exception:
            continue
        if connector is None or resolved_key != connector_key:
            continue
        seen_values.add(member_value)
        records.append((member, connector, resolved_key))
    return records


def _native_interconnect_points(start_point, end_point, branch_type,
                                bend_offset):
    arc_type = getattr(DB.Electrical.WiringType, "Arc", None)
    shape = HOMERUN_SHAPE_BEND if branch_type == arc_type else HOMERUN_SHAPE_STRAIGHT
    return _homerun_points(
        start_point,
        end_point,
        shape,
        bend_offset,
    )


def _custom_interconnect_homerun(document, view, wire_type_id,
                                 element, connector, settings,
                                 fallback_end=None):
    end_point = _homerun_end_point(
        element,
        connector,
        view,
        settings.get("homerun_length", HOME_RUN_LENGTH),
        settings.get("homerun_direction", HOMERUN_DIRECTION_PANEL),
        connector_type_key(connector),
        fallback_end=fallback_end,
    )
    homerun_points = _homerun_points(
        connector.Origin,
        end_point,
        settings.get("homerun_shape", HOMERUN_SHAPE_STRAIGHT),
        settings.get("bend_offset", 1.0),
    )
    homerun_type = wiring_type_from_name(
        settings.get("homerun_wiring_type"),
        "Arc",
    )
    return _create_wire_from_points(
        document,
        view,
        wire_type_id,
        homerun_type,
        connector,
        homerun_points,
    )


def _run_native_interconnect(document, view, device_elements, settings):
    circuits = circuits_from_elements(document, device_elements)
    if not circuits:
        return _run_spatial_interconnect(
            document,
            view,
            device_elements,
            settings,
        )
    if not settings.get("wire_type_id"):
        raise ValueError("Select a wire type before creating wires.")

    branch_type = wiring_type_from_name(
        settings.get("branch_wiring_type"),
        "Chamfer",
    )
    wire_type_id = element_id_from(settings["wire_type_id"])
    transaction = DB.Transaction(document, "Wire Tools - Native Interconnect")
    transaction.Start()
    created_count = 0
    deleted_count = 0
    failures = []
    homerun_records = []
    successful_circuit_values = set()
    reference_bend_offset = 0.0
    try:
        if settings.get("redraw_existing_wires", True):
            existing_wires = _view_wires_for_circuits(document, view.Id, circuits)
            deleted_count, deletion_failures = _delete_wires(
                document,
                existing_wires,
            )
            failures.extend(deletion_failures)

        for circuit in circuits:
            subtransaction = DB.SubTransaction(document)
            subtransaction.Start()
            try:
                wire_set = circuit.NewWires(view, branch_type)
                if not wire_set:
                    raise ValueError("ElectricalSystem.NewWires returned no wires.")
                wire_list = list(wire_set)
                _apply_wire_type(wire_list, wire_type_id)
                circuit_homerun_found = False
                for wire in wire_list:
                    if is_homerun_wire(wire):
                        record = _native_homerun_record(circuit, wire)
                        if record is not None:
                            homerun_records.append(record)
                            circuit_homerun_found = True
                    elif reference_bend_offset <= GEOMETRY_TOLERANCE:
                        candidate_offset = _wire_vertex_offset(wire)
                        if candidate_offset > GEOMETRY_TOLERANCE:
                            # Keep only the scalar geometry reference.  The
                            # native wire object is not retained across the
                            # later deletion and custom-creation steps.
                            reference_bend_offset = candidate_offset
                if not circuit_homerun_found:
                    raise ValueError(
                        "Revit created no resolvable homerun for the circuit."
                    )
                successful_circuit_values.add(element_id_value(circuit.Id))
                created_count += len(wire_list)
                subtransaction.Commit()
            except Exception as error:
                subtransaction.RollBack()
                failures.append({
                    "id": element_id_value(circuit.Id),
                    "element": circuit,
                    "reason": "Native circuit wiring failed: {}".format(error),
                })

        if not homerun_records:
            transaction.Commit()
            return created_count, [], failures, deleted_count

        keeper = min(
            homerun_records,
            key=lambda record: record["panel_distance"],
        )
        removed_records = [
            record for record in homerun_records if record is not keeper
        ]
        removed_wires = [record["wire"] for record in removed_records]
        removed_count, removed_failures = _delete_wires(document, removed_wires)
        deleted_count += removed_count
        created_count = max(created_count - removed_count, 0)
        failures.extend(removed_failures)

        homerun_id = element_id_value(keeper["wire"].Id)
        selected_homerun_record = None
        for selected_element in list(device_elements or []):
            try:
                selected_connector, selected_key = _resolve_device_connector(
                    selected_element,
                    keeper["connector_key"],
                )
            except Exception:
                continue
            if selected_connector is not None and selected_key == keeper["connector_key"]:
                selected_homerun_record = (
                    selected_element,
                    selected_connector,
                    selected_key,
                )
                break

        if selected_homerun_record is not None:
            selected_element, selected_connector, selected_key = selected_homerun_record
            native_open_point = keeper["open_connector"].Origin
            subtransaction = DB.SubTransaction(document)
            subtransaction.Start()
            try:
                document.Delete(keeper["wire"].Id)
                custom_homerun = _custom_interconnect_homerun(
                    document,
                    view,
                    wire_type_id,
                    selected_element,
                    selected_connector,
                    settings,
                    fallback_end=native_open_point,
                )
                homerun_id = element_id_value(custom_homerun.Id)
                subtransaction.Commit()
            except Exception as error:
                subtransaction.RollBack()
                failures.append({
                    "id": element_id_value(selected_element.Id),
                    "element": selected_element,
                    "reason": "Selected-device homerun replacement failed: {}".format(error),
                })

        # Use the native circuit network as each circuit's group.  The
        # selected-device homerun remains connected to its native circuit
        # network; every additional bridge is deliberately device-connector
        # to device-connector.  A wire connector is never a bridge endpoint.
        spatial_groups = {}
        circuit_keys = {}
        for record in homerun_records:
            circuit_value = element_id_value(record["circuit"].Id)
            circuit_keys[circuit_value] = record["connector_key"]
        for circuit in circuits:
            circuit_value = element_id_value(circuit.Id)
            if circuit_value not in successful_circuit_values:
                continue
            circuit_records = _circuit_connector_records(
                circuit,
                circuit_keys.get(circuit_value, keeper["connector_key"]),
            )
            if circuit_records:
                spatial_groups["circuit:{}".format(circuit_value)] = circuit_records
            else:
                failures.append({
                    "id": circuit_value,
                    "element": circuit,
                    "reason": "No usable member connector was found for spatial circuit bridging.",
                })
        bridge_edges = _spatial_group_mst(spatial_groups)

        bend_offset = reference_bend_offset
        for edge in bridge_edges:
            start_record = edge[3]
            end_record = edge[4]
            element = start_record[0]
            connector = start_record[1]
            previous_connector = end_record[1]
            first_owner = getattr(connector, "Owner", None)
            second_owner = getattr(previous_connector, "Owner", None)
            start_point = connector.Origin
            end_point = previous_connector.Origin
            subtransaction = DB.SubTransaction(document)
            subtransaction.Start()
            try:
                if (isinstance(first_owner, DB.Electrical.Wire)
                        or isinstance(second_owner, DB.Electrical.Wire)):
                    raise ValueError(
                        "Interconnect bridge endpoints must both be device connectors."
                    )
                points = _native_interconnect_points(
                    start_point,
                    end_point,
                    branch_type,
                    bend_offset,
                )
                _create_wire_from_points(
                    document,
                    view,
                    wire_type_id,
                    branch_type,
                    connector,
                    points,
                    end_connector=previous_connector,
                )
                created_count += 1
                subtransaction.Commit()
            except Exception as error:
                subtransaction.RollBack()
                failures.append({
                    "id": element_id_value(element.Id),
                    "element": element,
                    "reason": "Device-to-device interconnect failed: {}".format(error),
                })

        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    return created_count, [homerun_id], failures, deleted_count


def run_interconnect(document, view, elements, settings):
    if settings.get("interconnect_scope", INTERCONNECT_SCOPE_CIRCUITS) == INTERCONNECT_SCOPE_CIRCUITS:
        created_count, homerun_ids, failures, deleted_count = _run_native_interconnect(
            document,
            view,
            elements,
            settings,
        )
    else:
        created_count, homerun_ids, failures, deleted_count = _run_spatial_interconnect(
            document,
            view,
            elements,
            settings,
        )
    return {
        "created": created_count,
        "homeruns": homerun_ids,
        "deleted": deleted_count,
        "skipped": [],
        "failures": failures,
        "scheme": SCHEME_LABELS[SCHEME_INTERCONNECT],
    }


def run_individual_homeruns(document, view, elements, settings):
    created_count, homerun_ids, failures, deleted_count = _run_direct_wires(
        document,
        view,
        elements,
        settings,
        individual_homeruns=True,
    )
    return {
        "created": created_count,
        "homeruns": homerun_ids,
        "deleted": deleted_count,
        "skipped": [],
        "failures": failures,
        "scheme": SCHEME_LABELS[SCHEME_INDIVIDUAL_HOMERUN],
    }


def run_wire_to_node(document, view, elements, node_element, settings):
    if node_element is None or not is_valid_node(node_element):
        raise ValueError("Select a valid node with an electrical connector first.")
    created_count, homerun_ids, failures, deleted_count = _run_direct_wires(
        document,
        view,
        elements,
        settings,
        node_element=node_element,
    )
    return {
        "created": created_count,
        "homeruns": homerun_ids,
        "deleted": deleted_count,
        "skipped": [],
        "failures": failures,
        "scheme": SCHEME_LABELS[SCHEME_WIRE_TO_NODE],
    }


def run_scheme(document, view, scheme, device_elements, node_element, settings):
    if scheme == SCHEME_WIRE_BY_CIRCUIT:
        return run_wire_by_circuit(document, view, device_elements, settings)
    if scheme == SCHEME_INTERCONNECT:
        return run_interconnect(document, view, device_elements, settings)
    if scheme == SCHEME_INDIVIDUAL_HOMERUN:
        return run_individual_homeruns(document, view, device_elements, settings)
    if scheme == SCHEME_WIRE_TO_NODE:
        return run_wire_to_node(document, view, device_elements, node_element, settings)
    raise ValueError("Unknown wiring scheme: {}".format(scheme))


def _wire_open_connector(wire):
    connected, unconnected = wire_connected_unconnected_connectors(wire)
    del connected
    if len(unconnected) != 1:
        return None
    return unconnected[0]


def _tag_offset_distance(view):
    """Return a modest model-space clearance for a no-leader wire tag.

    The value is based on paper size so the visual gap remains useful at
    different view scales.  The clamps prevent an impractically small gap in
    very detailed views and an excessive gap in coarse views.
    """
    try:
        view_scale = float(view.Scale)
    except Exception:
        view_scale = 100.0
    model_distance = view_scale * TAG_OFFSET_PAPER_INCHES / 12.0
    return max(
        TAG_OFFSET_MINIMUM,
        min(TAG_OFFSET_MAXIMUM, model_distance),
    )


def _wire_tag_outward_direction(wire, open_connector, view):
    """Find a view-plane direction pointing away from the wire endpoint."""
    open_point = open_connector.Origin
    candidate_points = []

    try:
        for wire_connector in wire.ConnectorManager.Connectors:
            candidate_point = wire_connector.Origin
            if candidate_point.DistanceTo(open_point) > GEOMETRY_TOLERANCE:
                candidate_points.append(candidate_point)
    except Exception as error:
        script.get_logger().warning(
            "Could not inspect wire connector geometry for tag offset: {}".format(error)
        )

    try:
        vertex_count = int(getattr(wire, "NumberOfVertices", 0) or 0)
        for vertex_index in range(vertex_count):
            candidate_point = wire.GetVertex(vertex_index)
            if candidate_point.DistanceTo(open_point) > GEOMETRY_TOLERANCE:
                candidate_points.append(candidate_point)
    except Exception as error:
        if getattr(wire, "NumberOfVertices", 0):
            script.get_logger().warning(
                "Could not inspect wire vertex geometry for tag offset: {}".format(error)
            )

    if candidate_points:
        nearest_point = min(
            candidate_points,
            key=lambda point: point.DistanceTo(open_point),
        )
        direction = _project_direction(
            open_point.Subtract(nearest_point),
            view,
        )
        if direction is not None:
            return direction

    try:
        direction = _project_direction(view.RightDirection, view)
        if direction is not None:
            return direction
    except Exception:
        pass
    return DB.XYZ.BasisX


def _tag_bbox_points(bounding_box):
    for x_value in (bounding_box.Min.X, bounding_box.Max.X):
        for y_value in (bounding_box.Min.Y, bounding_box.Max.Y):
            for z_value in (bounding_box.Min.Z, bounding_box.Max.Z):
                yield DB.XYZ(x_value, y_value, z_value)


def _place_no_leader_tag(document, view, wire, open_connector, tag):
    """Move a no-leader tag clear of the open wire endpoint.

    The initial scale-aware offset is always applied.  After regeneration, the
    tag bounding box is used to add any extra clearance needed to keep the
    annotation body on the open-end side of the wire.  Bounding-box support is
    not identical across older Revit versions, so the initial offset remains a
    safe fallback when refinement is unavailable.
    """
    open_point = open_connector.Origin
    direction = _wire_tag_outward_direction(wire, open_connector, view)
    base_distance = _tag_offset_distance(view)
    try:
        tag.TagHeadPosition = open_point.Add(direction.Multiply(base_distance))
    except Exception as error:
        script.get_logger().warning(
            "No-leader wire tag could not be moved from the endpoint; "
            "the tag was left at its creation point: {}".format(error)
        )
        return

    try:
        document.Regenerate()
        bounding_box = tag.get_BoundingBox(view)
        if bounding_box is None:
            return

        required_gap = max(
            0.08,
            base_distance * TAG_OFFSET_BOUNDING_BOX_RATIO,
        )
        minimum_projection = min(
            point.Subtract(open_point).DotProduct(direction)
            for point in _tag_bbox_points(bounding_box)
        )
        if minimum_projection < required_gap:
            correction = required_gap - minimum_projection
            tag.TagHeadPosition = tag.TagHeadPosition.Add(
                direction.Multiply(correction)
            )
    except Exception as error:
        script.get_logger().warning(
            "No-leader wire tag clearance refinement was unavailable; "
            "the scale-aware offset was retained: {}".format(error)
        )


def _create_wire_tag(document, view, wire, tag_type_id, add_leader):
    open_connector = _wire_open_connector(wire)
    if open_connector is None:
        raise ValueError("Homerun does not have exactly one open connector.")
    point = open_connector.Origin
    if not add_leader:
        direction = _wire_tag_outward_direction(wire, open_connector, view)
        point = point.Add(direction.Multiply(_tag_offset_distance(view)))
    reference = DB.Reference(wire)
    creator = DB.IndependentTag.Create
    orientation = DB.TagOrientation.Horizontal
    errors = []
    try:
        created_tag = creator(
            document,
            tag_type_id,
            view.Id,
            reference,
            bool(add_leader),
            orientation,
            point,
        )
        if not add_leader:
            _place_no_leader_tag(document, view, wire, open_connector, created_tag)
        return created_tag
    except Exception as error:
        errors.append(error)
    try:
        created_tag = creator(
            document,
            view.Id,
            reference,
            bool(add_leader),
            DB.TagMode.TM_ADDBY_CATEGORY,
            orientation,
            point,
        )
        created_tag.ChangeTypeId(tag_type_id)
        if not add_leader:
            _place_no_leader_tag(document, view, wire, open_connector, created_tag)
        return created_tag
    except Exception as error:
        errors.append(error)
    raise RuntimeError(
        "Wire tag creation failed: {}".format(
            " | ".join([str(error) for error in errors])
        )
    )


def _tagged_local_element_values(tag):
    """Return local element ids referenced by an annotation tag."""
    values = []
    sources = []

    for method_name in ("GetTaggedLocalElementIds", "GetTaggedElementIds"):
        try:
            method = getattr(tag, method_name)
            sources.extend(list(method() or []))
        except Exception:
            pass

    for property_name in ("TaggedLocalElementId", "TaggedElementId"):
        try:
            sources.append(getattr(tag, property_name))
        except Exception:
            pass

    for source in sources:
        if source is None:
            continue
        try:
            local_id = getattr(source, "ElementId")
        except Exception:
            local_id = source
        value = element_id_value(local_id)
        if value > 0 and value not in values:
            values.append(value)
    return values


def wire_tag_index(document, view):
    """Index independent wire tags once by referenced local element ids."""
    index = {}
    collector = DB.FilteredElementCollector(document, view.Id)
    for tag in collector.OfClass(DB.IndependentTag).ToElements():
        tagged_values = _tagged_local_element_values(tag)
        for tagged_value in tagged_values:
            index.setdefault(tagged_value, []).append(tag)
    return index


def active_view_homerun_ids(document, view, skip_tagged=False):
    """Collect open-ended homerun wires from the active view."""
    wire_ids = []
    tagged_values = set()
    if skip_tagged:
        tagged_values = set(wire_tag_index(document, view).keys())
    collector = DB.FilteredElementCollector(document, view.Id)
    for wire in collector.OfClass(DB.Electrical.Wire).WhereElementIsNotElementType():
        try:
            wire_value = element_id_value(wire.Id)
            if is_homerun_wire(wire) and wire_value not in tagged_values:
                wire_ids.append(wire_value)
        except Exception as error:
            script.get_logger().warning(
                "Could not evaluate wire {} as a homerun: {}".format(
                    element_id_value(getattr(wire, "Id", None)),
                    error,
                )
            )
    return wire_ids


def tag_homeruns(document, view, wire_ids, tag_type_id, add_leader,
                 existing_behavior=TAG_EXISTING_SKIP):
    if not tag_type_id:
        raise ValueError("Select a Wire Tag type first.")
    existing_behavior = existing_behavior or TAG_EXISTING_SKIP
    transaction = DB.Transaction(document, "Wire Tools - Tag Homeruns")
    transaction.Start()
    created_count = 0
    deleted_count = 0
    skipped = []
    failures = []
    existing_index = wire_tag_index(document, view)
    try:
        for wire_value in list(wire_ids or []):
            wire = document.GetElement(element_id_from(wire_value))
            if wire is None or not isinstance(wire, DB.Electrical.Wire):
                failures.append({
                    "id": wire_value,
                    "element": wire,
                    "reason": "Selected element is not a wire.",
                })
                continue
            subtransaction = DB.SubTransaction(document)
            subtransaction.Start()
            try:
                if not is_homerun_wire(wire):
                    raise ValueError("Wire is not a homerun.")
                existing_tags = list(existing_index.get(element_id_value(wire.Id), []))
                if existing_tags and existing_behavior == TAG_EXISTING_SKIP:
                    skipped.append({
                        "id": wire_value,
                        "element": wire,
                        "reason": "Wire already has an existing tag.",
                    })
                    subtransaction.RollBack()
                    continue
                if existing_tags and existing_behavior == TAG_EXISTING_REPLACE:
                    unsafe_tags = [
                        existing_tag for existing_tag in existing_tags
                        if len(_tagged_local_element_values(existing_tag)) != 1
                    ]
                    if unsafe_tags:
                        raise ValueError(
                            "An existing multi-reference tag was found and was "
                            "not safely replaced."
                        )
                    for existing_tag in existing_tags:
                        document.Delete(existing_tag.Id)
                        deleted_count += 1
                _create_wire_tag(
                    document,
                    view,
                    wire,
                    element_id_from(tag_type_id),
                    add_leader,
                )
                created_count += 1
                subtransaction.Commit()
            except Exception as error:
                subtransaction.RollBack()
                failures.append({
                    "id": wire_value,
                    "element": wire,
                    "reason": "Homerun tag creation failed: {}".format(error),
                })
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    return {
        "created": created_count,
        "deleted": deleted_count,
        "skipped": skipped,
        "failures": failures,
    }


class DeviceSelectionFilter(ISelectionFilter):
    def __init__(self, scheme, excluded_element_id=None):
        self.scheme = scheme
        self.excluded_element_value = None
        if excluded_element_id is not None:
            try:
                self.excluded_element_value = element_id_value(excluded_element_id)
            except Exception:
                self.excluded_element_value = None

    def AllowElement(self, element):
        if self.excluded_element_value is not None:
            try:
                if element_id_value(element.Id) == self.excluded_element_value:
                    return False
            except Exception:
                pass
        if not is_allowed_device_pick(
            element,
            allow_circuit=self.scheme == SCHEME_WIRE_BY_CIRCUIT,
        ):
            return False
        return is_valid_device(
            element,
            allow_circuit=self.scheme == SCHEME_WIRE_BY_CIRCUIT,
        )

    def AllowReference(self, reference, position):
        del reference
        del position
        return False


class NodeSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return is_valid_node(element)

    def AllowReference(self, reference, position):
        del reference
        del position
        return False


def valid_device_ids(
        document,
        element_ids,
        scheme,
        diagnostics=None,
        requested_system_type=None):
    elements = []
    invalid = []
    seen_values = set()
    for element_id in list(element_ids or []):
        element_value = element_id_value(element_id)
        if element_value in seen_values:
            continue
        seen_values.add(element_value)
        detail = selection_validation_detail(
            document,
            element_id,
            scheme,
            requested_system_type,
        )
        if diagnostics is not None:
            diagnostics.append(detail)
        element = detail.get("element")
        candidate_valid = bool(detail.get("accepted"))
        if candidate_valid:
            elements.append(element)
            continue
        invalid.append({
            "id": detail.get("normalized_id", element_value),
            "element": element,
            "category": detail.get("category", ""),
            "reason": detail.get("reason", "Element failed selection validation."),
            "detail": detail,
        })
    return elements, invalid


def valid_homerun_ids(document, element_ids, view=None):
    valid_ids = []
    invalid = []
    seen_values = set()
    for element_id in list(element_ids or []):
        element_value = element_id_value(element_id)
        if element_value in seen_values:
            continue
        seen_values.add(element_value)
        element = document.GetElement(element_id_from(element_value))
        same_view = True
        if view is not None and element is not None:
            try:
                same_view = element_id_value(element.OwnerViewId) == element_id_value(
                    view.Id
                )
            except Exception:
                same_view = False
        if (isinstance(element, DB.Electrical.Wire)
                and same_view
                and is_homerun_wire(element)):
            valid_ids.append(element_value)
        else:
            invalid.append({
                "id": element_value,
                "element": element,
                "reason": (
                    "Element is not a valid homerun wire in the active view."
                    if view is not None
                    else "Element is not a valid homerun wire."
                ),
            })
    return valid_ids, invalid
