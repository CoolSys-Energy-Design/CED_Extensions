# -*- coding: utf-8 -*-
"""Model-wide electrical QC checks for one-shot review tools."""

import math
import re

from pyrevit import DB

from CEDElectrical.Domain import settings_manager
from CEDElectrical.Infrastructure.Revit.repositories import distribution_equipment_repository
from CEDElectrical.Infrastructure.Revit.repositories import panel_schedule_repository
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Model.alerts import get_alert_definition
from CEDElectrical.part_types import PART_TYPE_TRANSFORMER
from Snippets import categories as category_utils
from Snippets import revit_helpers


def _idval(item):
    return revit_helpers.get_elementid_value(item)


def _to_text(value, fallback=""):
    if value is None:
        return fallback
    try:
        return str(value)
    except Exception:
        return fallback


def _safe_float(value):
    if value is None:
        return None
    try:
        return float(value)
    except Exception:
        pass
    text = _to_text(value, "").strip().replace(",", "")
    if not text:
        return None
    match = re.search(r"[-+]?[0-9]*\.?[0-9]+", text)
    if not match:
        return None
    try:
        return float(match.group(0))
    except Exception:
        return None


def _as_unit(value, unit_names):
    number = _safe_float(value)
    if number is None:
        return None
    for unit_name in list(unit_names or []):
        try:
            unit = getattr(DB.UnitTypeId, unit_name)
            return float(DB.UnitUtils.ConvertFromInternalUnits(number, unit))
        except Exception:
            continue
    return number


def _amps(value):
    return _as_unit(value, ("Amperes", "Amps"))


def _volts(value):
    return _as_unit(value, ("Volts",))


def _feet(value):
    return _as_unit(value, ("Feet",))


def _kva(value):
    number = _safe_float(value)
    if number is None:
        return None
    try:
        return float(
            DB.UnitUtils.ConvertFromInternalUnits(
                number,
                DB.UnitTypeId.KilovoltAmperes,
            )
        )
    except Exception:
        return None


def _format_number(value, decimals=1):
    number = _safe_float(value)
    if number is None:
        return "-"
    if abs(number - round(number)) < 0.0001:
        return str(int(round(number)))
    return ("{0:." + str(int(decimals or 1)) + "f}").format(number)


def _format_amp(value):
    return "{} A".format(_format_number(value, 1))


def _format_kva(value):
    return "{} kVA".format(_format_number(value, 1))


def _format_ft(value):
    return "{} ft".format(_format_number(value, 1))


def _format_percent(value):
    return "{}%".format(_format_number(value, 1))


def _qc_device_categories(doc=None):
    return category_utils.get_fixture_device_categories(doc=doc)


def _format_id_array(values):
    ids = []
    for value in list(values or []):
        try:
            numeric = int(value or 0)
        except Exception:
            numeric = 0
        if numeric > 0 and numeric not in ids:
            ids.append(numeric)
    return "[{}]".format(", ".join([str(value) for value in ids]))


def _bip_text(element, bip):
    if element is None:
        return ""
    try:
        param = element.get_Parameter(bip)
        if param is None:
            return ""
        value = param.AsString()
        if value is None:
            value = param.AsValueString()
        return _to_text(value, "")
    except Exception:
        return ""


def _element_name(element):
    if element is None:
        return ""
    for attr in ("Name",):
        try:
            value = getattr(element, attr, None)
            if value:
                return _to_text(value, "")
        except Exception:
            pass
    try:
        return revit_helpers.get_family_symbol_name(element, fallback="")
    except Exception:
        return ""


def _element_label(element, fallback="-"):
    if element is None:
        return fallback
    return _element_name(element) or fallback


def _element_id_list(elements):
    ids = []
    for element in list(elements or []):
        element_id = _idval(getattr(element, "Id", None))
        if element_id > 0 and element_id not in ids:
            ids.append(element_id)
    return ids


def _connected_load_category_name(element):
    try:
        category = getattr(element, "Category", None)
        name = getattr(category, "Name", None)
        if name:
            return _to_text(name, "").strip()
    except Exception:
        pass
    return "Other Devices"


def _connected_load_type_identity(element):
    type_element = None
    try:
        type_element = revit_helpers.get_type_element(element)
    except Exception:
        type_element = None

    type_id = _idval(getattr(type_element, "Id", None))
    type_name = _element_name(type_element) or _element_name(element) or "Unknown Type"
    family_name = ""
    for candidate in (type_element, element):
        if candidate is None:
            continue
        try:
            family_name = _to_text(getattr(candidate, "FamilyName", None), "").strip()
        except Exception:
            family_name = ""
        if not family_name:
            try:
                family = getattr(candidate, "Family", None)
                family_name = _to_text(getattr(family, "Name", None), "").strip()
            except Exception:
                family_name = ""
        if family_name:
            break

    label = type_name
    if family_name and family_name.lower() != type_name.lower():
        label = "{}: {}".format(family_name, type_name)
    if type_id > 0:
        return ("id", type_id), label
    category_name = _connected_load_category_name(element)
    return ("label", category_name.lower(), label.lower()), label


def _format_counted_groups(groups):
    values = list(groups.values())
    values.sort(key=lambda item: (-item["count"], item["label"].lower()))
    return "; ".join("({}) {}".format(item["count"], item["label"]) for item in values)


def _elements_label(elements, fallback="-"):
    items = []
    seen_ids = set()
    for item in list(elements or []):
        if item is None:
            continue
        element_id = _idval(getattr(item, "Id", None))
        if element_id > 0:
            if element_id in seen_ids:
                continue
            seen_ids.add(element_id)
        items.append(item)
    if not items:
        return fallback

    type_groups = {}
    for item in items:
        type_key, type_label = _connected_load_type_identity(item)
        group = type_groups.setdefault(type_key, {"label": type_label, "count": 0})
        group["count"] += 1
    if len(type_groups) <= 3:
        return _format_counted_groups(type_groups) or fallback

    category_groups = {}
    for item in items:
        category_label = _connected_load_category_name(item)
        category_key = category_label.lower()
        group = category_groups.setdefault(category_key, {"label": category_label, "count": 0})
        group["count"] += 1
    return _format_counted_groups(category_groups) or fallback


def _display_component_label(kind, label, ids=None):
    text = _to_text(label, "").strip()
    id_count = len(list(ids or []))
    if not text or text == "-":
        return "-"
    if kind == "supply":
        return "Panel: {}".format(text)
    if kind == "circuit":
        return "Circuit: {}".format(text)
    if kind == "load":
        if id_count > 1:
            return "Devices: {}".format(text)
        return "Equipment/Device: {}".format(text)
    return text


def _select_primary_secondary_components(
    supply_label,
    supply_ids,
    circuit_label,
    circuit_ids,
    load_label,
    load_ids,
    highlight_supply=False,
    highlight_circuit=False,
    highlight_load=False,
):
    supply = {
        "kind": "supply",
        "label": _display_component_label("supply", supply_label, supply_ids),
        "ids": list(supply_ids or []),
        "highlight": bool(highlight_supply),
    }
    circuit = {
        "kind": "circuit",
        "label": _display_component_label("circuit", circuit_label, circuit_ids),
        "ids": list(circuit_ids or []),
        "highlight": bool(highlight_circuit),
    }
    load = {
        "kind": "load",
        "label": _display_component_label("load", load_label, load_ids),
        "ids": list(load_ids or []),
        "highlight": bool(highlight_load),
    }

    def _available(component):
        return bool(component.get("ids")) and _to_text(component.get("label"), "").strip() not in ("", "-")

    highlighted = [component for component in (supply, circuit, load) if component.get("highlight") and _available(component)]
    kinds = set([component.get("kind") for component in highlighted])
    if {"supply", "circuit", "load"}.issubset(kinds):
        ordered = [load, supply]
    elif {"supply", "circuit"}.issubset(kinds):
        ordered = [circuit, supply]
    elif {"circuit", "load"}.issubset(kinds):
        ordered = [circuit, load]
    elif {"supply", "load"}.issubset(kinds):
        ordered = [load, supply]
    else:
        ordered = highlighted
    if len(ordered) < 2:
        for component in (load, circuit, supply):
            if _available(component) and component not in ordered:
                ordered.append(component)
            if len(ordered) >= 2:
                break
    while len(ordered) < 2:
        ordered.append({"label": "-", "ids": []})
    return ordered[0], ordered[1]


def _extract_load_va_from_electrical_data(text, load_va_pattern):
    source = _to_text(text, "")
    if not source:
        return 0.0
    values = []
    for match in load_va_pattern.finditer(source):
        number = _safe_float(match.group(1))
        if number is None:
            continue
        unit = _to_text(match.group(2), "").upper()
        if unit == "KVA":
            number *= 1000.0
        values.append(number)
    if not values:
        return 0.0
    return max(values)


def _iter_circuit_elements(circuit):
    if circuit is None:
        return
    try:
        for element in list(getattr(circuit, "Elements", None) or []):
            if element is not None:
                yield element
        return
    except Exception:
        pass
    try:
        for element in list(circuit.GetCircuitElements() or []):
            if element is not None:
                yield element
    except Exception:
        return


def _fed_equipment(circuit, equipment_ids=None):
    ids = set([int(x) for x in list(equipment_ids or []) if int(x) > 0])
    found = []
    for element in _iter_circuit_elements(circuit):
        try:
            if _idval(element.Category.Id) != int(DB.BuiltInCategory.OST_ElectricalEquipment):
                continue
        except Exception:
            continue
        element_id = _idval(getattr(element, "Id", None))
        if ids and element_id not in ids:
            continue
        found.append(element)
    return found


def _current_from_kva(kva, voltage, poles):
    kva = _safe_float(kva)
    voltage = _safe_float(voltage)
    if kva is None or voltage is None or voltage <= 0:
        return None
    try:
        pole_count = int(poles or 1)
    except Exception:
        pole_count = 1
    divisor = voltage * (math.sqrt(3.0) if pole_count >= 3 else 1.0)
    if divisor <= 0:
        return None
    return (kva * 1000.0) / divisor


def _definition(definition_id):
    return get_alert_definition(definition_id)


def _definition_message(definition_id, fallback):
    definition = _definition(definition_id)
    if definition is None:
        return fallback
    try:
        return definition.GetDescriptionText()
    except Exception:
        return fallback


def _definition_generic_message(definition_id, fallback):
    text = _definition_message(definition_id, fallback)
    text = re.sub(r"\s*\([^)]*\{[^)]*\}[^)]*\)", "", _to_text(text, ""))
    text = re.sub(r":\s*\{[^}]+\}", "", text)
    text = re.sub(r"\{[^}]+\}", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace(" .", ".").replace(" ,", ",")
    return text or fallback


def _definition_severity(definition_id, fallback="MEDIUM"):
    definition = _definition(definition_id)
    if definition is None:
        return fallback
    try:
        return _to_text(definition.GetSeverity(), fallback).upper()
    except Exception:
        return fallback


def _definition_category(definition_id, fallback=None):
    definition = _definition(definition_id)
    if definition is None:
        return fallback
    try:
        category = _to_text(getattr(definition, "category", None), "").strip()
    except Exception:
        category = ""
    return category or fallback


def _category_from_definition(definition_id, fallback="General"):
    category = _definition_category(definition_id, None)
    if category:
        return category
    text = _to_text(definition_id, "")
    if text.startswith("QC."):
        parts = text.split(".")
        if len(parts) > 1 and parts[1]:
            area = parts[1]
            if area == "Circuit":
                return "Circuits"
            if area == "Device":
                return "Devices"
            if area == "Transformer":
                return "Transformers"
            if area in ("Equipment", "Panel"):
                return "Panels / Switchgear"
    return fallback


def _format_branch_type(value):
    text = _to_text(value, "").strip().upper()
    if text == "BRANCH":
        return "Branch"
    if text == "FEEDER":
        return "Feeder"
    if text == "XFMR PRI":
        return "XFMR Pri"
    if text == "XFMR SEC":
        return "XFMR Sec"
    if text == "CONDUIT ONLY":
        return "Conduit Only"
    if text in ("SPARE", "SPACE", "N/A"):
        return text
    return text.title() if text else "Unknown"


def _circuit_issue_category(branch=None):
    return "Circuits"


def _notice_definition_id(definition):
    if definition is None:
        return ""
    try:
        return _to_text(definition.GetId(), "")
    except Exception:
        return ""


def _notice_is_persistent(definition):
    if definition is None:
        return True
    try:
        return bool(getattr(definition, "persistent", True))
    except Exception:
        return True


def _branch_notice_observed(branch, alert_id):
    try:
        rating = getattr(branch, "rating", None)
    except Exception:
        rating = None
    try:
        current = getattr(branch, "circuit_load_current", None)
    except Exception:
        current = None
    if alert_id == "Design.UndersizedOCP":
        return "{} breaker / {} circuit load current".format(_format_amp(rating), _format_amp(current))
    if alert_id == "Design.NearOCPRating":
        percent = ""
        try:
            if rating and float(rating) > 0:
                percent = " = {:.1f}%".format(float(current or 0.0) / float(rating) * 100.0)
        except Exception:
            percent = ""
        return "{} circuit load current / {} breaker{}".format(_format_amp(current), _format_amp(rating), percent)
    if alert_id == "Design.CircuitLoadsNull":
        return "0 A circuit load current"
    if alert_id == "Design.CircuitPanelsNull":
        return "{} circuit load current / no panel".format(_format_amp(current))
    if alert_id == "Design.NonStandardOCPRating":
        return "{} breaker".format(_format_amp(rating))
    return ""


def _branch_notice_limit(alert_id):
    if alert_id == "Design.UndersizedOCP":
        return "Breaker rating >= circuit load current"
    if alert_id == "Design.NearOCPRating":
        return "Circuit load current < 90% of breaker rating"
    if alert_id == "Design.CircuitLoadsNull":
        return "Load > 0 A or intentional spare/space"
    if alert_id == "Design.CircuitPanelsNull":
        return "Circuit should have a supply panel"
    if alert_id == "Design.NonStandardOCPRating":
        return "Standard OCP rating"
    return ""


def _branch_notice_action(alert_id):
    if alert_id == "Design.UndersizedOCP":
        return "Upsize breaker or reduce connected load."
    if alert_id == "Design.NearOCPRating":
        return "Review breaker/loading. Upsize breaker or reduce load if margin is not acceptable."
    if alert_id == "Design.CircuitLoadsNull":
        return "Verify connected device loads."
    if alert_id == "Design.CircuitPanelsNull":
        return "Assign the circuit to the correct panel."
    if alert_id == "Design.NonStandardOCPRating":
        return "Revise breaker to a valid standard size or confirm special condition."
    return "Review circuit alert and correct the circuit properties."


class QCResultRow(object):
    """One row in the model-wide electrical QC report."""

    def __init__(self, **kwargs):
        severity = _to_text(kwargs.get("severity") or "MEDIUM", "MEDIUM").upper()
        self.category = kwargs.get("category") or "General"
        self.check_id = kwargs.get("check_id") or ""
        self.severity = severity
        self.supply_equipment = kwargs.get("supply_equipment") or "-"
        self.supply_equipment_id = int(kwargs.get("supply_equipment_id") or 0)
        self.supply_equipment_ids = list(kwargs.get("supply_equipment_ids") or [])
        self.circuit = kwargs.get("circuit") or "-"
        self.circuit_id = int(kwargs.get("circuit_id") or 0)
        self.circuit_ids = list(kwargs.get("circuit_ids") or [])
        self.load_device = kwargs.get("load_device") or "-"
        self.load_device_id = int(kwargs.get("load_device_id") or 0)
        self.load_device_ids = list(kwargs.get("load_device_ids") or [])
        self.issue = kwargs.get("issue") or ""
        self.generic_issue = kwargs.get("generic_issue") or self.issue
        self.observed = kwargs.get("observed") or ""
        self.limit_target = kwargs.get("limit_target") or ""
        self.recommended_action = kwargs.get("recommended_action") or ""
        self.supply_highlight_level = severity if bool(kwargs.get("highlight_supply")) else ""
        self.circuit_highlight_level = severity if bool(kwargs.get("highlight_circuit")) else ""
        self.load_highlight_level = severity if bool(kwargs.get("highlight_load")) else ""
        primary, secondary = _select_primary_secondary_components(
            self.supply_equipment,
            self.supply_equipment_ids,
            self.circuit,
            self.circuit_ids,
            self.load_device,
            self.load_device_ids,
            highlight_supply=bool(kwargs.get("highlight_supply")),
            highlight_circuit=bool(kwargs.get("highlight_circuit")),
            highlight_load=bool(kwargs.get("highlight_load")),
        )
        self.primary_element = primary.get("label") or "-"
        self.primary_element_ids = list(primary.get("ids") or [])
        self.secondary_element = secondary.get("label") or "-"
        self.secondary_element_ids = list(secondary.get("ids") or [])

    def export_values(self):
        return {
            "category": self.category,
            "check_id": self.check_id,
            "severity": self.severity,
            "primary_element": self.primary_element,
            "primary_element_ids": _format_id_array(self.primary_element_ids),
            "secondary_element": self.secondary_element,
            "secondary_element_ids": _format_id_array(self.secondary_element_ids),
            "supply_equipment": self.supply_equipment,
            "supply_equipment_id": self.supply_equipment_id,
            "supply_equipment_ids": _format_id_array(self.supply_equipment_ids),
            "circuit": self.circuit,
            "circuit_id": self.circuit_id,
            "circuit_ids": _format_id_array(self.circuit_ids),
            "load_device": self.load_device,
            "load_device_id": self.load_device_id,
            "load_device_ids": _format_id_array(self.load_device_ids),
            "issue": self.issue,
            "generic_issue": self.generic_issue,
            "observed": self.observed,
            "limit_target": self.limit_target,
            "recommended_action": self.recommended_action,
            "supply_highlight_level": self.supply_highlight_level,
            "circuit_highlight_level": self.circuit_highlight_level,
            "load_highlight_level": self.load_highlight_level,
        }


class ElectricalQCScanner(object):
    """Runs best-effort electrical QC checks against the active model."""

    def __init__(self, logger=None, load_va_pattern=None):
        self.logger = logger
        self.rows = []
        self.doc = None
        self.equipment_by_id = {}
        self.model_by_id = {}
        self.schedule_by_panel_id = {}
        self.circuit_by_id = {}
        self.branch_by_id = {}
        self.circuit_settings = None
        self.design_option_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
        pattern_text = load_va_pattern or r"([0-9][0-9,]*(?:\.[0-9]+)?)\s*(KVA|VA)\b"
        self.load_va_pattern = re.compile(pattern_text, re.IGNORECASE)

    def _log_debug(self, message):
        if self.logger is None:
            return
        try:
            self.logger.debug(message)
        except Exception:
            pass

    def _add(
        self,
        check_id,
        supply_equipment=None,
        circuit=None,
        load_device=None,
        load_devices=None,
        observed="",
        limit_target="",
        recommended_action="",
        category=None,
        severity=None,
        issue=None,
        highlight_supply=False,
        highlight_circuit=False,
        highlight_load=False,
    ):
        load_device_list = list(load_devices or [])
        load_device_ids = _element_id_list(load_device_list)
        load_device_label = _elements_label(load_device_list) if load_device_list else _element_label(load_device)
        if not load_device_ids:
            single_load_id = _idval(getattr(load_device, "Id", None))
            if single_load_id > 0:
                load_device_ids = [single_load_id]
        supply_id = _idval(getattr(supply_equipment, "Id", None))
        circuit_id = _idval(getattr(circuit, "Id", None))
        self.rows.append(
            QCResultRow(
                category=category or _category_from_definition(check_id),
                check_id=check_id,
                severity=severity or _definition_severity(check_id),
                supply_equipment=_element_label(supply_equipment),
                supply_equipment_id=supply_id,
                supply_equipment_ids=[supply_id] if supply_id > 0 else [],
                circuit=self._circuit_label(circuit),
                circuit_id=circuit_id,
                circuit_ids=[circuit_id] if circuit_id > 0 else [],
                load_device=load_device_label,
                load_device_id=load_device_ids[0] if load_device_ids else 0,
                load_device_ids=load_device_ids,
                issue=issue or _definition_message(check_id, check_id),
                generic_issue=_definition_generic_message(check_id, issue or check_id),
                observed=observed,
                limit_target=limit_target,
                recommended_action=recommended_action,
                highlight_supply=highlight_supply,
                highlight_circuit=highlight_circuit,
                highlight_load=highlight_load,
            )
        )

    def scan(self, doc):
        self.doc = doc
        self.rows = []
        if doc is None:
            return self._snapshot("No active document")
        try:
            self.circuit_settings = settings_manager.load_circuit_settings(doc)
        except Exception as ex:
            self._log_debug("QC circuit settings load failed: {}".format(ex))
            self.circuit_settings = None

        equipment = self._collect_equipment()
        circuits = self._collect_circuits()
        devices = self._collect_devices()
        self.circuit_by_id = {}
        self.branch_by_id = {}
        for circuit in circuits:
            cid = _idval(getattr(circuit, "Id", None))
            if cid > 0:
                self.circuit_by_id[cid] = circuit

        self._prepare_equipment_models(equipment)
        self._check_equipment(equipment)
        self._check_devices(devices)
        self._check_circuits(circuits)
        self._check_transformers()
        self._check_panel_slots()
        self._check_supplies_larger_equipment(circuits)

        return self._snapshot("ok")

    def _snapshot(self, status):
        title = "-"
        try:
            title = self.doc.Title if self.doc is not None else "-"
        except Exception:
            title = "-"
        return {
            "status": status,
            "doc_title": title,
            "rows": list(self.rows or []),
            "count": len(self.rows or []),
        }

    def _collect_equipment(self):
        collector = (
            DB.FilteredElementCollector(self.doc)
            .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)
            .WhereElementIsNotElementType()
            .WherePasses(self.design_option_filter)
        )
        return list(collector.ToElements())

    def _collect_devices(self):
        devices = []
        for category in _qc_device_categories(doc=self.doc):
            try:
                collector = (
                    DB.FilteredElementCollector(self.doc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .WherePasses(self.design_option_filter)
                )
                devices.extend(list(collector.ToElements()))
            except Exception:
                continue
        return devices

    def _collect_circuits(self):
        collector = (
            DB.FilteredElementCollector(self.doc)
            .OfCategory(DB.BuiltInCategory.OST_ElectricalCircuit)
            .WhereElementIsNotElementType()
            .WherePasses(self.design_option_filter)
        )
        return list(collector.ToElements())

    def _prepare_equipment_models(self, equipment):
        self.equipment_by_id = {}
        self.model_by_id = {}
        try:
            self.schedule_by_panel_id = panel_schedule_repository.map_panel_schedule_views(self.doc, panels=equipment)
        except Exception:
            self.schedule_by_panel_id = {}
        for item in list(equipment or []):
            eid = _idval(getattr(item, "Id", None))
            if eid <= 0:
                continue
            self.equipment_by_id[eid] = item
            schedule = self.schedule_by_panel_id.get(eid)
            try:
                model = distribution_equipment_repository.build_distribution_equipment(
                    self.doc,
                    item,
                    schedule_view=schedule,
                )
            except Exception as ex:
                self._log_debug("QC equipment model failed for {}: {}".format(eid, ex))
                model = None
            if model is not None:
                self.model_by_id[eid] = model

    def _branch_for_circuit(self, circuit):
        circuit_id = _idval(getattr(circuit, "Id", None))
        if circuit_id <= 0:
            return None
        if circuit_id in self.branch_by_id:
            return self.branch_by_id.get(circuit_id)
        branch = None
        try:
            branch = CircuitBranch(circuit, settings=self.circuit_settings)
        except Exception as ex:
            self._log_debug("QC CircuitBranch failed for {}: {}".format(circuit_id, ex))
        self.branch_by_id[circuit_id] = branch
        return branch

    def _circuit_label(self, circuit):
        branch = self._branch_for_circuit(circuit)
        if branch is None:
            return "-"
        panel = _to_text(getattr(branch, "panel", ""), "").strip() or "-"
        number = _to_text(getattr(branch, "circuit_number", ""), "").strip() or "-"
        label = "{}/{}".format(panel, number)
        name = _to_text(getattr(branch, "load_name", ""), "").strip()
        if name:
            label = "{} - {}".format(label, name)
        return label

    def _is_regular_power_branch(self, branch):
        if branch is None:
            return False
        try:
            return bool(branch.is_power_circuit) and not branch.is_space and not branch.is_spare
        except Exception:
            return False

    def _circuit_rating_value(self, circuit):
        branch = self._branch_for_circuit(circuit)
        if branch is not None:
            try:
                return _amps(getattr(branch, "rating", None))
            except Exception:
                pass
        return None

    def _circuit_length_value(self, circuit):
        branch = self._branch_for_circuit(circuit)
        if branch is not None:
            try:
                return _feet(getattr(branch, "length", None))
            except Exception:
                pass
        return None

    def _circuit_poles_value(self, circuit):
        branch = self._branch_for_circuit(circuit)
        if branch is None:
            return None
        try:
            return int(getattr(branch, "poles", None) or 0)
        except Exception:
            return None

    def _check_equipment(self, equipment):
        for item in list(equipment or []):
            eid = _idval(getattr(item, "Id", None))
            model = self.model_by_id.get(eid)
            if model is None:
                continue
            if getattr(model, "part_type", None) == PART_TYPE_TRANSFORMER:
                continue
            mains = _amps(getattr(model, "mains_rating", None))
            mcb = _amps(getattr(model, "ocp_rating", None))
            demand = _amps(getattr(model, "current_demand_total", None))
            supply_circuits = distribution_equipment_repository.supply_circuits_for_model(self.doc, model)
            supply = supply_circuits[0] if supply_circuits else None
            has_mcb = mcb is not None and not bool(getattr(model, "is_mlo", False))

            if mains is not None and demand is not None and mains < demand:
                self._add(
                    "QC.Equipment.MainsBelowDemand",
                    supply_equipment=item,
                    circuit=supply,
                    observed="{} mains / {} demand".format(_format_amp(mains), _format_amp(demand)),
                    limit_target="Mains rating >= total demand current",
                    recommended_action="Review equipment mains rating or reduce connected demand.",
                    highlight_supply=True,
                )
            if has_mcb and demand is not None and mcb < demand:
                self._add(
                    "QC.Equipment.McbBelowDemand",
                    supply_equipment=item,
                    circuit=supply,
                    observed="{} MCB / {} demand".format(_format_amp(mcb), _format_amp(demand)),
                    limit_target="MCB rating >= total demand current",
                    recommended_action="Review equipment MCB rating or reduce connected demand.",
                    highlight_supply=True,
                )
            if has_mcb and mains is not None and mcb > mains:
                self._add(
                    "QC.Equipment.McbExceedsMains",
                    supply_equipment=item,
                    circuit=supply,
                    observed="{} MCB / {} mains".format(_format_amp(mcb), _format_amp(mains)),
                    limit_target="MCB rating <= mains rating",
                    recommended_action="Correct equipment mains or MCB rating.",
                    highlight_supply=True,
                )
            for supply in list(supply_circuits or []):
                supply_rating = self._circuit_rating_value(supply)
                if supply_rating is None:
                    continue
                if mains is not None and supply_rating > mains:
                    self._add(
                        "QC.Equipment.SupplyBreakerExceedsMains",
                        supply_equipment=item,
                        circuit=supply,
                        observed="{} feeder breaker / {} mains".format(_format_amp(supply_rating), _format_amp(mains)),
                        limit_target="Supply breaker <= equipment mains rating",
                        recommended_action="Review feeder OCP and downstream equipment mains (bus) rating. This is only acceptable if downstream equipment has a MCB less than or equal to the bus rating.",
                        highlight_supply=True,
                        highlight_circuit=True,
                    )
                if has_mcb and supply_rating > mcb:
                    self._add(
                        "QC.Equipment.SupplyBreakerExceedsMcb",
                        supply_equipment=item,
                        circuit=supply,
                        observed="{} feeder breaker / {} MCB".format(_format_amp(supply_rating), _format_amp(mcb)),
                        limit_target="Supply breaker <= downstream MCB rating",
                        recommended_action="Review feeder OCP coordination or correct MCB rating. This is acceptable if intentional.",
                        highlight_supply=True,
                        highlight_circuit=True,
                    )
                if has_mcb and mcb > supply_rating:
                    self._add(
                        "QC.Equipment.McbExceedsSupplyBreaker",
                        supply_equipment=item,
                        circuit=supply,
                        observed="{} MCB / {} feeder breaker".format(_format_amp(mcb), _format_amp(supply_rating)),
                        limit_target="MCB <= upstream breaker unless intentionally coordinated",
                        recommended_action="Review MCB and feeder breaker sizing.",
                        highlight_supply=True,
                        highlight_circuit=True,
                    )

    def _check_devices(self, devices):
        for device in list(devices or []):
            data = _bip_text(device, DB.BuiltInParameter.RBS_ELECTRICAL_DATA)
            load_va = _extract_load_va_from_electrical_data(data, self.load_va_pattern)
            if load_va <= 0:
                continue
            circuit_number = _bip_text(device, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
            panel_name = _bip_text(device, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
            observed = "{} VA".format(_format_number(load_va, 1))
            if not circuit_number and not panel_name:
                self._add(
                    "QC.Device.LoadNotConnected",
                    load_device=device,
                    observed=observed,
                    limit_target="Loaded devices should be assigned to a circuit",
                    recommended_action="Create or assign an electrical circuit for this device.",
                    highlight_load=True,
                )
            elif circuit_number == "<unnamed>" and not panel_name:
                device_circuits = distribution_equipment_repository.electrical_systems_for_element(device)
                device_circuit = device_circuits[0] if device_circuits else None
                self._add(
                    "QC.Device.LoadNoPanel",
                    circuit=device_circuit,
                    load_device=device,
                    observed="{} / circuit {}".format(observed, circuit_number),
                    limit_target="Circuit should have a supply panel",
                    recommended_action="Assign the circuit to the correct panel.",
                    highlight_load=True,
                )

    def _check_circuits(self, circuits):
        for circuit in list(circuits or []):
            branch = self._branch_for_circuit(circuit)
            if not self._is_regular_power_branch(branch):
                continue
            base = None
            try:
                base = getattr(circuit, "BaseEquipment", None)
            except Exception:
                base = None

            try:
                notices = branch.collect_qc_notices() if branch is not None else []
            except Exception:
                notices = []
            for definition, severity, group, message in list(notices or []):
                alert_id = _notice_definition_id(definition)
                if not alert_id or not _notice_is_persistent(definition):
                    continue
                connected_loads = []
                if alert_id in (
                    "Design.CircuitLoadsNull",
                    "Design.CircuitPanelsNull",
                    "Design.NearOCPRating",
                    "Design.UndersizedOCP",
                ):
                    connected_loads = list(_iter_circuit_elements(circuit) or [])
                self._add(
                    alert_id,
                    supply_equipment=base,
                    circuit=circuit,
                    load_devices=connected_loads,
                    observed=_branch_notice_observed(branch, alert_id),
                    limit_target=_branch_notice_limit(alert_id),
                    recommended_action=_branch_notice_action(alert_id),
                    category=_category_from_definition(alert_id, _circuit_issue_category(branch)),
                    severity=severity,
                    issue=message,
                    highlight_circuit=True,
                    highlight_load=bool(connected_loads),
                )

    def _check_transformers(self):
        for eid, model in list((self.model_by_id or {}).items()):
            if getattr(model, "part_type", None) != PART_TYPE_TRANSFORMER:
                continue
            equipment = self.equipment_by_id.get(eid)
            supply = distribution_equipment_repository.primary_supply_circuit_for_model(self.doc, model)
            supply_base = None
            try:
                supply_base = getattr(supply, "BaseEquipment", None)
            except Exception:
                supply_base = None
            rating_kva = _kva(getattr(model, "xfmr_rating", None))
            demand_kva = _kva(getattr(model, "power_demand_total", None))

            if rating_kva is None:
                self._add(
                    "QC.Transformer.RatingMissing",
                    load_device=equipment,
                    observed="Transformer Rating_CED blank",
                    limit_target="Transformer Rating_CED populated",
                    recommended_action="Add Transformer Rating_CED to the family/type or verify this is not a standard transformer family.",
                    highlight_load=True,
                )
            elif demand_kva is not None and demand_kva > rating_kva:
                self._add(
                    "QC.Transformer.DemandExceedsRating",
                    supply_equipment=supply_base,
                    circuit=supply,
                    load_device=equipment,
                    observed="{} demand / {} rating".format(_format_kva(demand_kva), _format_kva(rating_kva)),
                    limit_target="Demand kVA <= transformer kVA rating",
                    recommended_action="Upsize transformer or reduce downstream demand.",
                    highlight_load=True,
                )

            # Disabled for now: transformer primary OCP is commonly oversized for inrush,
            # so demand near breaker is not a useful QC warning.
            for circuit_id in list(getattr(model, "branch_circuits", []) or []):
                circuit = self.circuit_by_id.get(int(circuit_id or 0))
                length = self._circuit_length_value(circuit)
                if circuit is not None and length is not None and length > 25.0:
                    self._add(
                        "QC.Transformer.SecondaryLengthOver25Ft",
                        supply_equipment=equipment,
                        circuit=circuit,
                        observed=_format_ft(length),
                        limit_target="<= 25 ft",
                        recommended_action="Verify a remote secondary OCP device is specified within 25 ft of the transformer. May be ignored if this is a utility transformer.",
                        highlight_supply=True,
                        highlight_circuit=True,
                    )

    def _check_panel_slots(self):
        try:
            options = panel_schedule_repository.collect_panel_equipment_options(
                self.doc,
                panels=list(self.equipment_by_id.values()),
                include_without_schedule=False,
            )
        except Exception:
            options = []
        for option in list(options or []):
            panel_id = int(option.get("panel_id", 0) or 0)
            panel = self.equipment_by_id.get(panel_id)
            if panel is None:
                continue
            try:
                rows = panel_schedule_repository.build_panel_rows(
                    self.doc,
                    option,
                    panel_id_set=self.equipment_by_id.keys(),
                    all_circuits=list(self.circuit_by_id.values()),
                )
            except Exception:
                continue
            open_slots = []
            for row in list(rows or []):
                if not bool(row.get("is_valid_slot", True)):
                    continue
                kind = _to_text(row.get("kind", ""), "").strip().lower()
                if kind == "empty":
                    open_slots.append(row)
            if open_slots:
                self._add(
                    "QC.Panel.OpenSlots",
                    supply_equipment=panel,
                    observed="{} open slot(s)".format(len(open_slots)),
                    limit_target="Informational",
                    recommended_action="Review whether open slots should be left open, filled with spaces, or reserved. Use custom 'Add Spares/Spaces' tool to quickly resolve.",
                    highlight_supply=True,
                )

    def _check_supplies_larger_equipment(self, circuits):
        equipment_ids = set(self.equipment_by_id.keys())
        for circuit in list(circuits or []):
            source = None
            try:
                source = getattr(circuit, "BaseEquipment", None)
            except Exception:
                source = None
            source_id = _idval(getattr(source, "Id", None))
            source_model = self.model_by_id.get(source_id)
            if source_model is None:
                continue
            if not bool(getattr(source_model, "is_panel_or_switchgear", False)):
                continue
            source_rating = _amps(getattr(source_model, "mains_rating", None))
            source_mcb = _amps(getattr(source_model, "ocp_rating", None))
            if source_rating is None or (source_mcb is not None and source_mcb > source_rating):
                source_rating = source_mcb
            if source_rating is None:
                continue
            for downstream in _fed_equipment(circuit, equipment_ids):
                downstream_id = _idval(getattr(downstream, "Id", None))
                if downstream_id == source_id:
                    continue
                downstream_model = self.model_by_id.get(downstream_id)
                if downstream_model is None:
                    continue
                downstream_rating = _amps(getattr(downstream_model, "mains_rating", None))
                downstream_mcb = _amps(getattr(downstream_model, "ocp_rating", None))
                if downstream_rating is None or (downstream_mcb is not None and downstream_mcb > downstream_rating):
                    downstream_rating = downstream_mcb
                if downstream_rating is None or downstream_rating <= source_rating:
                    continue
                self._add(
                    "QC.Panel.SuppliesLargerEquipment",
                    supply_equipment=source,
                    circuit=circuit,
                    load_device=downstream,
                    observed="{} source rating / {} supplied equipment rating".format(
                        _format_amp(source_rating),
                        _format_amp(downstream_rating),
                    ),
                    limit_target="Supplied equipment rating <= source panel rating",
                    recommended_action="Verify source panel rating, feeder, and downstream equipment data.",
                    highlight_supply=True,
                    highlight_circuit=True,
                    highlight_load=True,
                )


def scan_document(doc, logger=None, standard_breaker_sizes=None, load_va_pattern=None):
    """Run all electrical QC checks and return a snapshot dictionary."""
    scanner = ElectricalQCScanner(
        logger=logger,
        load_va_pattern=load_va_pattern,
    )
    return scanner.scan(doc)


def _default_export_columns():
    return [
        ("category", "Category"),
        ("check_id", "Check ID"),
        ("severity", "Severity"),
        ("issue", "Issue"),
        ("recommended_action", "Recommended Action"),
        ("observed", "Observed"),
        ("primary_element", "Primary Element"),
        ("primary_element_ids", "Primary Element IDs"),
        ("secondary_element", "Secondary Element"),
        ("secondary_element_ids", "Secondary Element IDs"),
    ]


def _csv_escape(value):
    text = _to_text(value, "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if '"' in text:
        text = text.replace('"', '""')
    if "," in text or '"' in text or "\n" in text:
        return '"{}"'.format(text)
    return text


def rows_to_csv(rows, columns=None):
    """Serialize QC rows to CSV text, including hidden UI columns."""
    export_columns = list(columns or _default_export_columns())
    lines = []
    lines.append(",".join([_csv_escape(label) for _, label in export_columns]))
    for row in list(rows or []):
        data = row.export_values() if hasattr(row, "export_values") else dict(row or {})
        lines.append(",".join([_csv_escape(data.get(key, "")) for key, _ in export_columns]))
    return "\r\n".join(lines) + "\r\n"
