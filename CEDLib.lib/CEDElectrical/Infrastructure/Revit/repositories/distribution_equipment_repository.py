# -*- coding: utf-8 -*-
"""Revit mappers for DistributionEquipment domain models."""

import Autodesk.Revit.DB.Electrical as DBE
from System import Guid
from pyrevit import DB

from CEDElectrical.Model.distribution_equipment import DistributionEquipment, PowerBus, Transformer
from CEDElectrical.part_types import (
    PART_TYPE_MAP,
    PART_TYPE_OTHER_PANEL,
    PART_TYPE_PANELBOARD,
    PART_TYPE_SWITCHBOARD,
    PART_TYPE_TRANSFORMER,
)
from Snippets import _elecutils as eu
from Snippets import design_options, revit_helpers

PSTYPE_UNKNOWN = DBE.PanelScheduleType.Unknown
PSTYPE_BRANCH = DBE.PanelScheduleType.Branch
PSTYPE_SWITCHBOARD = DBE.PanelScheduleType.Switchboard
PSTYPE_DATA = DBE.PanelScheduleType.Data

PART_TYPE_TO_PANEL_SCHEDULE_TYPE = {
    PART_TYPE_PANELBOARD: PSTYPE_BRANCH,
    PART_TYPE_SWITCHBOARD: PSTYPE_SWITCHBOARD,
    PART_TYPE_OTHER_PANEL: PSTYPE_DATA,
}


def _idval(item):
    """Return numeric value for ElementId-like objects."""
    return revit_helpers.get_elementid_value(item)


def _to_text(value, fallback=""):
    """Return safe string conversion."""
    if value is None:
        return fallback
    try:
        return str(value)
    except Exception:
        return fallback


BIP_ELEC_PANEL_CONFIGURATION = DB.BuiltInParameter.RBS_ELEC_PANEL_CONFIGURATION_PARAM
BIP_FAMILY_CONTENT_PART_TYPE = DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE
BIP_FAMILY_DIST_SYSTEM = DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM
BIP_FAMILY_SECONDARY_DIST_SYSTEM = DB.BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS
BIP_SYMBOL_NAME = DB.BuiltInParameter.SYMBOL_NAME_PARAM
BIP_VOLTAGE_TYPE_VOLTAGE = DB.BuiltInParameter.RBS_VOLTAGETYPE_VOLTAGE_PARAM
BIP_PANEL_TOTAL_LOAD = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM
BIP_PANEL_TOTAL_CONNECTED_LOAD = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM
BIP_PANEL_TOTAL_CURRENT = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM
BIP_PANEL_TOTAL_CONNECTED_CURRENT = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM
BIP_PANEL_TOTAL_DEMAND_LOAD = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALESTLOAD_PARAM
BIP_PANEL_TOTAL_DEMAND_CURRENT = DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM
BIP_PANEL_CONNECTED_CURRENT_PHASEA = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_CURRENT_PHASEA
BIP_PANEL_CONNECTED_CURRENT_PHASEB = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_CURRENT_PHASEB
BIP_PANEL_CONNECTED_CURRENT_PHASEC = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_CURRENT_PHASEC
BIP_PANEL_CONNECTED_LOAD_PHASEA = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_APPARENT_LOAD_PHASEA
BIP_PANEL_CONNECTED_LOAD_PHASEB = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_APPARENT_LOAD_PHASEB
BIP_PANEL_CONNECTED_LOAD_PHASEC = DB.BuiltInParameter.RBS_ELEC_PANEL_BRANCH_CIRCUIT_APPARENT_LOAD_PHASEC
BIP_PANEL_NAME =DB.BuiltInParameter.RBS_ELEC_PANEL_NAME
BIP_PANEL_MCB_RATING = DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM
BIP_PANEL_MAINS_RATING = DB.BuiltInParameter.RBS_ELEC_MAINS
BIP_PANEL_MODIFICATIONS = DB.BuiltInParameter.RBS_ELEC_MODIFICATIONS
BIP_PANEL_MOUNTING = DB.BuiltInParameter.RBS_ELEC_MOUNTING
BIP_PANEL_ENCLOSURE = DB.BuiltInParameter.RBS_ELEC_ENCLOSURE
BIP_PANEL_FEED_THRU_LUGS = DB.BuiltInParameter.RBS_ELEC_PANEL_FEED_THRU_LUGS_PARAM
BIP_PANEL_FEED = DB.BuiltInParameter.RBS_ELEC_PANEL_FEED_PARAM
BIP_PANEL_MAX_BREAKERS = DB.BuiltInParameter.RBS_ELEC_MAX_POLE_BREAKERS
BIP_PANEL_MAX_CIRCUITS = DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_CIRCUITS
BIP_PANEL_SHORT_CIRCUIT_RATING = DB.BuiltInParameter.RBS_ELEC_SHORT_CIRCUIT_RATING
BIP_PANEL_MAINS_TYPE = getattr(DB.BuiltInParameter, "RBS_ELEC_PANEL_MAINS_TYPE_PARAM", None)
BIP_PANEL_NEUTRAL_BUS = getattr(DB.BuiltInParameter, "RBS_ELEC_PANEL_NEUTRAL_BUS_PARAM", None)


EQUIPMENT_PARAMETER_DEFINITIONS = {
    "isolated_ground_bus": {
        "shared": {
            "name": "Isolated Ground Bus_CED",
            "guid": "fce623e5-837b-448f-a4a9-0e6bb9b3fbe9",
            "include_type": False,
        },
        "builtins": [],
        "value_type": "yesno",
    },
    "main_breaker_rating": {
        "shared": {
            "name": "Main Breaker Rating_CED",
            "guid": "fac2a3cf-802d-4ccf-8854-91da0b41d091",
            "include_type": False,
        },
        "builtins": [{"bip": BIP_PANEL_MCB_RATING, "include_type": False}],
        "value_type": "current",
    },
    "mains_rating": {
        "shared": {
            "name": "Mains Rating_CED",
            "guid": "7e6dfe1f-1be5-4493-aa4e-012e9e3802fe",
            "include_type": True,
        },
        "builtins": [{"bip": BIP_PANEL_MAINS_RATING, "include_type": False}],
        "value_type": "current",
    },
    "mains_type": {
        "shared": {
            "name": "Mains Type_CEDT",
            "guid": "cd8f0de1-db57-4638-946c-caf8d2594aa4",
            "include_type": True,
        },
        "builtins": [{"bip": BIP_PANEL_MAINS_TYPE, "include_type": True}],
        "value_type": "text",
    },
    "neutral_bus": {
        "shared": {
            "name": "Neutral Bus_CED",
            "guid": "65a0822e-c08c-43f8-980c-3ebd59d66a27",
            "include_type": True,
        },
        "builtins": [{"bip": BIP_PANEL_NEUTRAL_BUS, "include_type": True}],
        "value_type": "yesno",
    },
    "panel_name": {
        "shared": {
            "name": "Panel Name_CEDT",
            "guid": "43088534-de5e-4b2a-a001-4190a054bb94",
            "include_type": False,
        },
        "builtins": [{"bip": BIP_PANEL_NAME, "include_type": False}],
        "value_type": "text",
    },
    "short_circuit_rating": {
        "shared": {
            "name": "Short Circuit Rating_CEDT",
            "guid": "fd6628b7-d134-4eec-87b0-6f1e50242ffe",
            "include_type": False,
        },
        "builtins": [{"bip": BIP_PANEL_SHORT_CIRCUIT_RATING, "include_type": False}],
        "value_type": "text",
    },
    "transformer_impedance_percent": {
        "shared": {
            "name": "Transformer %Z_CED",
            "guid": "76b07259-7a21-45a3-8d23-5cb68af8fd7d",
            "include_type": True,
        },
        "builtins": [],
        "value_type": "number",
    },
    "transformer_rating": {
        "shared": {
            "name": "Transformer Rating_CED",
            "guid": "1f76eb82-5f63-41c4-89d0-dd5697ae5dd6",
            "include_type": True,
        },
        "builtins": [],
        "value_type": "apparent_power",
    },
    "distribution_system": {
        "shared": None,
        "builtins": [{"bip": BIP_FAMILY_DIST_SYSTEM, "include_type": False}],
        "value_type": "elementid",
    },
    "secondary_distribution_system": {
        "shared": None,
        "builtins": [{"bip": BIP_FAMILY_SECONDARY_DIST_SYSTEM, "include_type": False}],
        "value_type": "elementid",
    },
    "feed_thru_lugs": {
        "shared": None,
        "builtins": [{"bip": BIP_PANEL_FEED_THRU_LUGS, "include_type": False}],
        "value_type": "yesno",
    },
    "max_single_pole_breakers": {
        "shared": None,
        "builtins": [{"bip": BIP_PANEL_MAX_BREAKERS, "include_type": False}],
        "value_type": "integer",
    },
    "max_circuits": {
        "shared": None,
        "builtins": [{"bip": BIP_PANEL_MAX_CIRCUITS, "include_type": False}],
        "value_type": "integer",
    },
}


def _param_has_value(param):
    if param is None:
        return False
    try:
        return bool(param.HasValue)
    except Exception:
        return True


def _value_is_present(value):
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    return True


def _param_value(param, default=None):
    """Return native parameter value."""
    return revit_helpers.get_parameter_value(param, default=default)


def _param_from_guid(element, guid_text, include_type=False):
    """Return shared parameter by GUID, optionally falling back to type."""
    if element is None or not guid_text:
        return None

    def _lookup(owner):
        if owner is None:
            return None
        try:
            param = owner.get_Parameter(Guid(str(guid_text)))
        except Exception:
            param = None
        if _param_has_value(param):
            return param
        return None

    param = _lookup(element)
    if param is not None or not bool(include_type):
        return param
    return _lookup(revit_helpers.get_type_element(element))


def _param_from_bips(element, bips, include_type=False):
    """Return first non-empty built-in parameter value."""
    owners = [element]
    if bool(include_type):
        owners.append(revit_helpers.get_type_element(element))
    for bip in list(bips or []):
        if bip is None:
            continue
        for owner in owners:
            if owner is None:
                continue
            try:
                param = owner.get_Parameter(bip)
            except Exception:
                param = None
            if not _param_has_value(param):
                continue
            value = _param_value(param, default=None)
            if _value_is_present(value):
                return value
    return None


def _param_record_from_bip(element, bip, include_type=False):
    """Return parameter record for an approved built-in fallback."""
    if bip is None:
        return None
    owners = [("builtin", element)]
    if bool(include_type):
        owners.append(("builtin_type", revit_helpers.get_type_element(element)))
    for source_kind, owner in owners:
        if owner is None:
            continue
        try:
            param = owner.get_Parameter(bip)
        except Exception:
            param = None
        if not _param_has_value(param):
            continue
        value = _param_value(param, default=None)
        if not _value_is_present(value):
            continue
        return {
            "value": value,
            "source": source_kind,
            "builtin": str(bip),
            "name": "",
            "guid": "",
        }
    return None


def _equipment_parameter_record(equipment, key):
    """Resolve one standardized equipment parameter from exact shared/BIP rules."""
    definition = EQUIPMENT_PARAMETER_DEFINITIONS.get(key) or {}
    shared = definition.get("shared")
    if shared:
        param = _param_from_guid(
            equipment,
            shared.get("guid"),
            include_type=bool(shared.get("include_type", False)),
        )
        if param is not None:
            value = _param_value(param, default=None)
            if _value_is_present(value):
                return {
                    "value": value,
                    "source": "shared",
                    "builtin": "",
                    "name": shared.get("name", ""),
                    "guid": shared.get("guid", ""),
                }

    for item in list(definition.get("builtins") or []):
        record = _param_record_from_bip(
            equipment,
            item.get("bip"),
            include_type=bool(item.get("include_type", False)),
        )
        if record is not None:
            return record
    return {"value": None, "source": "", "builtin": "", "name": "", "guid": ""}


def _equipment_parameter_value(equipment, key, default=None):
    record = _equipment_parameter_record(equipment, key)
    value = record.get("value")
    return default if value is None else value


def _equipment_parameter_snapshot(equipment):
    """Return values and source metadata for standardized equipment parameters."""
    values = {}
    sources = {}
    for key in sorted(EQUIPMENT_PARAMETER_DEFINITIONS.keys()):
        record = _equipment_parameter_record(equipment, key)
        values[key] = record.get("value")
        sources[key] = {
            "source": record.get("source", ""),
            "name": record.get("name", ""),
            "guid": record.get("guid", ""),
            "builtin": record.get("builtin", ""),
        }
    return values, sources


def _enum_equals(value, target):
    """Return True when two enum-like values represent the same member."""
    try:
        if value == target:
            return True
    except Exception:
        pass
    try:
        return int(value) == int(target)
    except Exception:
        pass
    return False


PCFG_ONE_COLUMN = DBE.PanelConfiguration.OneColumn
PCFG_TWO_COLUMNS_ACROSS = DBE.PanelConfiguration.TwoColumnsCircuitsAcross
PCFG_TWO_COLUMNS_DOWN = DBE.PanelConfiguration.TwoColumnsCircuitsDown


def _normalize_panel_configuration(value):
    """Return canonical DBE.PanelConfiguration member when possible."""
    if value is None:
        return None
    for candidate in (PCFG_ONE_COLUMN, PCFG_TWO_COLUMNS_ACROSS, PCFG_TWO_COLUMNS_DOWN):
        if candidate is None:
            continue
        if _enum_equals(value, candidate):
            return candidate
    return None


def _family_parameter_from_bip(equipment, bip):
    """Return built-in parameter from Family definition of an instance."""
    if equipment is None:
        return None
    if not isinstance(equipment, DB.FamilyInstance):
        return None
    try:
        symbol = equipment.Symbol
    except Exception:
        symbol = None
    try:
        family = symbol.Family if symbol is not None else None
    except Exception:
        family = None
    if family is None:
        return None
    try:
        param = family.get_Parameter(bip)
    except Exception:
        param = None
    if param is None:
        return None
    try:
        if not bool(param.HasValue):
            return None
    except Exception:
        pass
    return param


def _panel_configuration_for_equipment(equipment, part_type):
    """Return panel configuration using part-type rules and family fallback."""
    if part_type in (PART_TYPE_SWITCHBOARD, PART_TYPE_OTHER_PANEL):
        return PCFG_ONE_COLUMN
    if part_type != PART_TYPE_PANELBOARD:
        return PCFG_ONE_COLUMN
    param = _family_parameter_from_bip(equipment, BIP_ELEC_PANEL_CONFIGURATION)
    value = _param_value(param, default=None)
    config = _normalize_panel_configuration(value)
    if config is not None:
        return config
    return PCFG_ONE_COLUMN


def _param_bool(value):
    """Normalize mixed parameter value types into bool/None."""
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    try:
        return bool(int(value))
    except Exception:
        pass
    return None


def _volts_from_internal(value):
    """Convert Revit internal electrical value to volts when possible."""
    if value is None:
        return None
    try:
        return DB.UnitUtils.ConvertFromInternalUnits(float(value), DB.UnitTypeId.Volts)
    except Exception:
        return value


def get_family_part_type(equipment):
    """Return FAMILY_CONTENT_PART_TYPE integer from family definition."""
    if equipment is None or not isinstance(equipment, DB.FamilyInstance):
        return None
    try:
        symbol = equipment.Symbol
        family = symbol.Family if symbol else None
        param = family.get_Parameter(BIP_FAMILY_CONTENT_PART_TYPE) if family else None
        if param and param.HasValue and param.StorageType == DB.StorageType.Integer:
            return int(param.AsInteger())
    except Exception:
        pass
    return None


def equipment_type_from_part_type(part_type):
    """Return equipment type label from family part type."""
    return PART_TYPE_MAP.get(part_type, "Unknown")


def expected_panel_schedule_type_for_equipment(equipment):
    """Return expected DBE.PanelScheduleType from equipment part type."""
    part_type = get_family_part_type(equipment)
    if part_type in PART_TYPE_TO_PANEL_SCHEDULE_TYPE:
        return PART_TYPE_TO_PANEL_SCHEDULE_TYPE.get(part_type, PSTYPE_BRANCH)
    return PSTYPE_UNKNOWN


def _distribution_system_snapshot(doc, dist_system_id):
    """Return distribution system snapshot from DistributionSysType element id."""
    result = {
        "id": 0,
        "name": "",
        "phase": None,
        "wire_count": None,
        "lg_voltage": None,
        "ll_voltage": None,
    }
    if dist_system_id is None:
        return result
    try:
        dist_id_val = _idval(dist_system_id)
    except Exception:
        dist_id_val = 0
    if dist_id_val <= 0:
        return result
    result["id"] = dist_id_val
    dist = doc.GetElement(dist_system_id)
    if dist is None:
        return result
    try:
        name_param = dist.get_Parameter(BIP_SYMBOL_NAME)
        if name_param and name_param.HasValue:
            result["name"] = _to_text(name_param.AsString(), "")
    except Exception:
        pass
    try:
        result["phase"] = dist.ElectricalPhase
    except Exception:
        pass
    try:
        value = dist.NumWires
        if value is not None:
            wires = int(value)
            if wires > 0:
                result["wire_count"] = wires
    except Exception:
        pass
    try:
        lg = dist.VoltageLineToGround
        if lg is not None:
            lg_param = lg.get_Parameter(BIP_VOLTAGE_TYPE_VOLTAGE)
            if lg_param and lg_param.HasValue:
                result["lg_voltage"] = _volts_from_internal(lg_param.AsDouble())
    except Exception:
        pass
    try:
        ll = dist.VoltageLineToLine
        if ll is not None:
            ll_param = ll.get_Parameter(BIP_VOLTAGE_TYPE_VOLTAGE)
            if ll_param and ll_param.HasValue:
                result["ll_voltage"] = _volts_from_internal(ll_param.AsDouble())
    except Exception:
        pass
    return result


def _branch_circuit_options(primary_profile, secondary_profile=None):
    """Build branch-circuit voltage/pole options from distribution profiles."""
    options = []

    def _append_options(profile, source):
        if not profile:
            return
        lg = profile.get("lg_voltage")
        ll = profile.get("ll_voltage")
        phase = profile.get("phase")
        is_single_phase = False
        try:
            if phase == DBE.ElectricalPhase.SinglePhase:
                is_single_phase = True
        except Exception:
            pass
        if lg is not None:
            options.append({"source": source, "poles": 1, "voltage": lg})
        if ll is not None:
            options.append({"source": source, "poles": 2, "voltage": ll})
            if not is_single_phase:
                options.append({"source": source, "poles": 3, "voltage": ll})

    _append_options(primary_profile, "primary")
    _append_options(secondary_profile, "secondary")
    deduped = []
    seen = set()
    for item in options:
        key = (int(item.get("poles", 0) or 0), _to_text(item.get("voltage"), ""))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)
    return deduped


def _system_ids(systems):
    """Return sorted circuit id values from ElectricalSystem collections."""
    ids = []
    for system in list(systems or []):
        try:
            cid = _idval(system.Id)
        except Exception:
            cid = 0
        if cid > 0:
            ids.append(cid)
    ids.sort()
    return ids


def _connector_bool_attr(connector, names):
    for name in list(names or []):
        try:
            value = getattr(connector, name)
        except Exception:
            continue
        try:
            if callable(value):
                value = value()
        except Exception:
            continue
        if value is not None:
            return bool(value)
    return None


def _connector_system_ids(connector):
    ids = []

    def _add_system(system):
        if not design_options.is_main_model_element(system):
            return
        try:
            sid = _idval(system.Id)
        except Exception:
            sid = 0
        if sid > 0 and sid not in ids:
            ids.append(sid)

    for attr_name in ("MEPSystem", "ElectricalSystem"):
        try:
            system = getattr(connector, attr_name, None)
        except Exception:
            system = None
        if system is not None:
            _add_system(system)

    for method_name in ("GetMEPSystems", "GetElectricalSystems"):
        try:
            method = getattr(connector, method_name, None)
            systems = list(method() or []) if callable(method) else []
        except Exception:
            systems = []
        for system in systems:
            _add_system(system)

    return ids


def _equipment_connectors(equipment):
    connectors = []
    try:
        mep = equipment.MEPModel
    except Exception:
        mep = None
    try:
        connector_manager = getattr(mep, "ConnectorManager", None)
    except Exception:
        connector_manager = None
    if connector_manager is None:
        return connectors
    for attr_name in ("Connectors", "UnusedConnectors"):
        try:
            for connector in list(getattr(connector_manager, attr_name, None) or []):
                if connector is not None and connector not in connectors:
                    connectors.append(connector)
        except Exception:
            continue
    return connectors


def _supply_connection_records(equipment, supply_systems):
    """Return connector-aware supply records for distribution equipment."""
    records = []
    by_id = {}
    for system in list(supply_systems or []):
        try:
            circuit_id = _idval(system.Id)
        except Exception:
            circuit_id = 0
        if circuit_id <= 0 or circuit_id in by_id:
            continue
        record = {
            "circuit_id": circuit_id,
            "is_primary": None,
            "connector_id": 0,
        }
        by_id[circuit_id] = record
        records.append(record)

    if not records:
        return records

    for connector in _equipment_connectors(equipment):
        is_primary = _connector_bool_attr(connector, ("IsPrimary", "isPrimary", "isprimary"))
        connector_id = _idval(getattr(connector, "Id", None))
        for system_id in _connector_system_ids(connector):
            record = by_id.get(system_id)
            if record is None:
                continue
            if is_primary is not None:
                record["is_primary"] = bool(is_primary)
            if connector_id > 0:
                record["connector_id"] = connector_id

    primary_records = [x for x in records if bool(x.get("is_primary"))]
    if not primary_records and records:
        records[0]["is_primary"] = True
    for record in records:
        if record.get("is_primary") is None:
            record["is_primary"] = False
    records.sort(key=lambda x: (0 if bool(x.get("is_primary")) else 1, int(x.get("circuit_id") or 0)))
    return records


def electrical_systems_for_element(element):
    """Return unique electrical systems assigned to an MEP-backed element."""
    systems = []
    seen = set()
    if element is None:
        return systems
    try:
        mep_model = getattr(element, "MEPModel", None)
    except Exception:
        mep_model = None
    if mep_model is None:
        return systems

    candidates = []
    for method_name in ("GetAssignedElectricalSystems", "GetElectricalSystems"):
        try:
            method = getattr(mep_model, method_name, None)
        except Exception:
            method = None
        if method is None:
            continue
        try:
            candidates.extend(list(method() or []))
        except Exception:
            pass
    try:
        candidates.extend(list(getattr(mep_model, "ElectricalSystems", None) or []))
    except Exception:
        pass

    for system in candidates:
        system_id = _idval(getattr(system, "Id", None))
        if system_id <= 0 or system_id in seen:
            continue
        if not eu.is_circuit_eligible(system):
            continue
        seen.add(system_id)
        systems.append(system)
    return systems


def supply_circuits_for_model(doc, model, primary_only=False):
    """Resolve a distribution equipment model's supply records to Revit circuits."""
    circuits = []
    if doc is None or model is None:
        return circuits
    if bool(primary_only):
        primary = getattr(model, "primary_supply", None)
        records = [primary] if primary else []
    else:
        records = list(getattr(model, "supply_connections", []) or [])
        if not records:
            records = [
                {"circuit_id": int(circuit_id or 0), "is_primary": False}
                for circuit_id in list(getattr(model, "supply_circuits", []) or [])
            ]

    for record in records:
        try:
            circuit_id = int(record.get("circuit_id") or 0)
        except Exception:
            circuit_id = 0
        if circuit_id <= 0:
            continue
        try:
            circuit = doc.GetElement(revit_helpers.elementid_from_value(circuit_id))
        except Exception:
            circuit = None
        if eu.is_circuit_eligible(circuit):
            circuits.append(circuit)
    return circuits


def primary_supply_circuit_for_model(doc, model):
    """Resolve the primary supply circuit for a distribution equipment model."""
    for circuit in supply_circuits_for_model(doc, model, primary_only=True):
        return circuit
    for circuit in supply_circuits_for_model(doc, model, primary_only=False):
        return circuit
    return None


def _schedule_slot_count(schedule_view):
    """Return schedule slot count from PanelScheduleView."""
    if schedule_view is None:
        return 0
    try:
        table = schedule_view.GetTableData()
        return int(table.NumberOfSlots or 0)
    except Exception:
        return 0


def _total_power_current_snapshot(equipment):
    """Return connected/demand panel load totals when available."""
    return {
        "power_connected_total": _param_from_bips(
            equipment,
            [
                BIP_PANEL_TOTAL_LOAD,
                BIP_PANEL_TOTAL_CONNECTED_LOAD,
            ],
        ),
        "current_connected_total": _param_from_bips(
            equipment,
            [
                BIP_PANEL_TOTAL_CURRENT,
                BIP_PANEL_TOTAL_CONNECTED_CURRENT,
            ],
        ),
        "power_demand_total": _param_from_bips(
            equipment,
            [BIP_PANEL_TOTAL_DEMAND_LOAD],
        ),
        "current_demand_total": _param_from_bips(
            equipment,
            [BIP_PANEL_TOTAL_DEMAND_CURRENT],
        ),
        "branch_current_phase_a": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_CURRENT_PHASEA],
        ),
        "branch_current_phase_b": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_CURRENT_PHASEB],
        ),
        "branch_current_phase_c": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_CURRENT_PHASEC],
        ),
        "branch_load_phase_a": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_LOAD_PHASEA],
        ),
        "branch_load_phase_b": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_LOAD_PHASEB],
        ),
        "branch_load_phase_c": _param_from_bips(
            equipment,
            [BIP_PANEL_CONNECTED_LOAD_PHASEC],
        ),
    }


def build_distribution_equipment(doc, equipment, schedule_view=None):
    """Map a Revit electrical equipment instance into a domain model object."""
    if equipment is None or not design_options.is_main_model_element(equipment):
        return None

    part_type = get_family_part_type(equipment)
    equipment_type = equipment_type_from_part_type(part_type)

    parameter_values, parameter_sources = _equipment_parameter_snapshot(equipment)

    primary_dist_id = parameter_values.get("distribution_system")
    secondary_dist_id = parameter_values.get("secondary_distribution_system")
    primary_profile = _distribution_system_snapshot(doc, primary_dist_id)
    secondary_profile = _distribution_system_snapshot(doc, secondary_dist_id)

    mep = None
    try:
        mep = equipment.MEPModel
    except Exception:
        mep = None

    all_systems = []
    assigned_systems = []
    if mep is not None:
        try:
            all_systems = eu.filter_circuits(mep.GetElectricalSystems() or [])
        except Exception:
            all_systems = []
        try:
            assigned_systems = eu.filter_circuits(mep.GetAssignedElectricalSystems() or [])
        except Exception:
            assigned_systems = []
    assigned_ids = set(_system_ids(assigned_systems))
    supply_systems = []
    for system in list(all_systems or []):
        try:
            sid = _idval(system.Id)
        except Exception:
            sid = 0
        if sid > 0 and sid not in assigned_ids:
            supply_systems.append(system)

    mains_rating = parameter_values.get("mains_rating")
    mains_type = parameter_values.get("mains_type")
    ocp_rating = parameter_values.get("main_breaker_rating")
    short_circuit_rating = parameter_values.get("short_circuit_rating")
    panel_name = parameter_values.get("panel_name")
    has_feed_thru_lugs = _param_bool(parameter_values.get("feed_thru_lugs"))
    has_neutral_bus = _param_bool(parameter_values.get("neutral_bus"))
    has_ground_bus = None
    has_isolated_ground_bus = _param_bool(parameter_values.get("isolated_ground_bus"))

    totals = _total_power_current_snapshot(equipment)
    options = _branch_circuit_options(primary_profile, secondary_profile)

    voltage = primary_profile.get("ll_voltage") or primary_profile.get("lg_voltage")
    poles = None
    if options:
        poles = max([int(x.get("poles", 0) or 0) for x in options if int(x.get("poles", 0) or 0) > 0] or [None])

    supply_connections = _supply_connection_records(equipment, supply_systems)

    max_poles = None
    if part_type in (PART_TYPE_PANELBOARD, PART_TYPE_TRANSFORMER, PART_TYPE_OTHER_PANEL):
        max_poles = parameter_values.get("max_single_pole_breakers")
        try:
            max_poles = int(max_poles or 0)
        except Exception:
            max_poles = 0
        if max_poles <= 0:
            max_poles = _schedule_slot_count(schedule_view)
    elif part_type == PART_TYPE_SWITCHBOARD:
        max_poles = 0
        mep_model = getattr(equipment, "MEPModel", None)
        if mep_model is not None:
            for attr in ("MaxNumberOfCircuits", "maxNumberOfCircuits"):
                try:
                    value = int(getattr(mep_model, attr, 0) or 0)
                except Exception:
                    value = 0
                if value > 0:
                    max_poles = int(value)
                    break
        if max_poles <= 0:
            value = parameter_values.get("max_circuits")
            try:
                max_poles = int(value or 0)
            except Exception:
                max_poles = 0
        if max_poles <= 0:
            max_poles = _schedule_slot_count(schedule_view)

    equipment_name = None
    try:
        equipment_name = equipment.Name
    except Exception:
        equipment_name = None

    base_kwargs = {
        "id": _idval(equipment.Id),
        "name": _to_text(panel_name, None) or _to_text(equipment_name, None),
        "element_name": _to_text(equipment_name, None),
        "panel_name": _to_text(panel_name, None),
        "part_type": part_type,
        "equipment_type": equipment_type,
        "parameter_values": parameter_values,
        "parameter_sources": parameter_sources,
        "voltage": voltage,
        "poles": poles,
        "distribution_system": primary_profile,
        "distribution_system_secondary": secondary_profile,
        "supply_connections": supply_connections,
        "supply_circuits": [int(x.get("circuit_id") or 0) for x in supply_connections],
        "branch_circuits": _system_ids(assigned_systems),
        "branch_circuit_options": options,
        "mains_rating": mains_rating,
        "mains_type": mains_type,
        "has_ocp": bool(ocp_rating not in (None, "", 0)),
        "ocp_type": _to_text(mains_type, None),
        "ocp_rating": ocp_rating,
        "has_feed_thru_lugs": has_feed_thru_lugs,
        "has_neutral_bus": has_neutral_bus,
        "has_ground_bus": has_ground_bus,
        "has_isolated_ground_bus": has_isolated_ground_bus,
        "max_poles": max_poles,
        "short_circuit_rating": short_circuit_rating,
        "power_connected_total": totals.get("power_connected_total"),
        "current_connected_total": totals.get("current_connected_total"),
        "power_demand_total": totals.get("power_demand_total"),
        "current_demand_total": totals.get("current_demand_total"),
        "branch_current_phase_a": totals.get("branch_current_phase_a"),
        "branch_current_phase_b": totals.get("branch_current_phase_b"),
        "branch_current_phase_c": totals.get("branch_current_phase_c"),
        "branch_load_phase_a": totals.get("branch_load_phase_a"),
        "branch_load_phase_b": totals.get("branch_load_phase_b"),
        "branch_load_phase_c": totals.get("branch_load_phase_c"),
    }

    if part_type == PART_TYPE_TRANSFORMER:
        base_kwargs.update(
            {
                "xfmr_rating": parameter_values.get("transformer_rating"),
                "xfmr_impedance": parameter_values.get("transformer_impedance_percent"),
                "xfmr_kfactor": None,
            }
        )
        return Transformer(**base_kwargs)

    if part_type in (PART_TYPE_PANELBOARD, PART_TYPE_SWITCHBOARD, PART_TYPE_OTHER_PANEL):
        panel_configuration = _panel_configuration_for_equipment(equipment, part_type)
        base_kwargs.update(
            {
                "has_panel_schedule": bool(schedule_view is not None),
                "panel_configuration": panel_configuration,
            }
        )
        return PowerBus(**base_kwargs)

    return DistributionEquipment(**base_kwargs)
