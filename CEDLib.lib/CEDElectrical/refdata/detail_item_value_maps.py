# -*- coding: utf-8 -*-
"""Canonical circuit/panel -> one-line detail item parameter maps.

The one-line diagram detail items carry CED shared parameters that mirror
circuit and panel data from the model. Two identity parameters say which
circuit/panel a detail item represents:

    CKT_Panel_CEDT + CKT_Circuit Number_CEDT  -> which circuit
    Panel Name_CEDT                           -> which panel (equipment symbols)

Everything else is copied data. These maps are the single source of truth
for that copy; they mirror the maps in "Sync One Line.pushbutton/script.py"
(which should eventually import from here instead of keeping its own copy).
"""

from pyrevit import DB

DETAIL_PARAM_CKT_PANEL = "CKT_Panel_CEDT"
DETAIL_PARAM_CKT_NUMBER = "CKT_Circuit Number_CEDT"
DETAIL_PARAM_CKT_LOAD_NAME = "CKT_Load Name_CEDT"
DETAIL_PARAM_PANEL_NAME = "Panel Name_CEDT"

# Detail item parameter name -> source on the circuit (ElectricalSystem).
# Values are either a BuiltInParameter or a shared-parameter name string.
CIRCUIT_VALUE_MAP = {
    "x VD Schedule": "x VD Schedule",
    "Circuit Tree Sort_CED": "Circuit Tree Sort_CED",
    "CKT_Circuit Type_CEDT": "CKT_Circuit Type_CEDT",
    "CKT_Panel_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM,
    "CKT_Circuit Number_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER,
    "CKT_Load Name_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME,
    "CKT_Rating_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
    "CKT_Frame_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
    "CKT_Schedule Notes_CEDT": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM,
    "CKT_Length_CED": DB.BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM,
    "Number of Poles_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES,
    "Voltage_CED": DB.BuiltInParameter.RBS_ELEC_VOLTAGE,
    "Wire Material_CEDT": "Wire Material_CEDT",
    "Wire Insulation_CEDT": "Wire Insulation_CEDT",
    "Wire Temparature Rating_CEDT": "Wire Temparature Rating_CEDT",
    "Wire Size_CEDT": "Wire Size_CEDT",
    "Conduit and Wire Size_CEDT": "Conduit and Wire Size_CEDT",
    "Conduit Type_CEDT": "Conduit Type_CEDT",
    "Conduit Size_CEDT": "Conduit Size_CEDT",
    "Conduit Fill Percentage_CED": "Conduit Fill Percentage_CED",
    "Voltage Drop Percentage_CED": "Voltage Drop Percentage_CED",
    "Circuit Load Current_CED": "Circuit Load Current_CED",
}

# Detail item parameter name -> source on the panel (ElectricalEquipment).
PANEL_VALUE_MAP = {
    "Panel Name_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_NAME,
    "Mains Rating_CED": "Mains Rating_CED",
    "Mains Type_CEDT": "Mains Type_CEDT",
    "Phase_CED": "Phase_CED",
    "Main Breaker Rating_CED": DB.BuiltInParameter.RBS_ELEC_PANEL_MCB_RATING_PARAM,
    "Short Circuit Rating_CEDT": DB.BuiltInParameter.RBS_ELEC_SHORT_CIRCUIT_RATING,
    "Mounting_CEDT": DB.BuiltInParameter.RBS_ELEC_MOUNTING,
    "Panel Modifications_CEDT": DB.BuiltInParameter.RBS_ELEC_MODIFICATIONS,
    "Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM,
    "Secondary Distribution System_CEDR": DB.BuiltInParameter.RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS,
    "Total Connected Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALLOAD_PARAM,
    # Demand LOAD comes from Total Estimated Demand (apparent power), not the
    # demand-current param - amps written into this VA-typed detail param
    # displayed as garbage (e.g. 3195 A -> "297 VA").
    "Total Demand Load_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTALESTLOAD_PARAM,
    "Total Connected Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_CONNECTED_CURRENT_PARAM,
    "Total Demand Current_CEDR": DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM,
    "Max Number of Single Pole Breakers_CED": DB.BuiltInParameter.RBS_ELEC_MAX_POLE_BREAKERS,
    "Max Number of Circuits_CED": DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_CIRCUITS,
    "Transformer Rating_CEDT": "Transformer Rating_CEDT",
    "Transformer Rating_CED": "Transformer Rating_CEDT",
    "Transformer Primary Description_CEDT": "Transformer Primary Description_CEDT",
    "Transformer Secondary Description_CEDT": "Transformer Secondary Description_CEDT",
    "Transformer %Z_CED": "Transformer %Z_CED",
    "Panel Feed_CEDT": DB.BuiltInParameter.RBS_ELEC_PANEL_FEED_PARAM,
}


def read_source_param_value(elem, param_key, allow_type_fallback=True):
    """Read a BuiltInParameter or shared parameter (by name) off an element.

    Checks the instance first, then the type when allowed. Returns
    string/int/float (or the value string for ElementId params) or None.
    """
    param = None
    if isinstance(param_key, DB.BuiltInParameter):
        param = elem.get_Parameter(param_key)
    elif isinstance(param_key, str):
        param = elem.LookupParameter(param_key)
    else:
        return None

    if not param and allow_type_fallback:
        try:
            type_elem = elem.Document.GetElement(elem.GetTypeId())
            if type_elem:
                if isinstance(param_key, DB.BuiltInParameter):
                    param = type_elem.get_Parameter(param_key)
                else:
                    param = type_elem.LookupParameter(param_key)
        except Exception:
            param = None

    if not param:
        return None

    st = param.StorageType
    if st == DB.StorageType.String:
        return param.AsString()
    if st == DB.StorageType.Integer:
        return param.AsInteger()
    if st == DB.StorageType.Double:
        return param.AsDouble()
    if st == DB.StorageType.ElementId:
        return param.AsValueString()
    return None


def _set_param(param, value):
    try:
        st = param.StorageType
        if st == DB.StorageType.String:
            param.Set("" if value is None else str(value))
            return True
        if value is None:
            return False
        if st == DB.StorageType.Integer:
            param.Set(int(round(float(value))))
            return True
        if st == DB.StorageType.Double:
            param.Set(float(value))
            return True
        return False
    except Exception:
        return False


def _apply_value_map(source_elem, detail_item, value_map):
    written = []
    missing = []
    skipped = []
    for detail_name, source_key in value_map.items():
        target = None
        try:
            target = detail_item.LookupParameter(detail_name)
        except Exception:
            target = None
        if target is None:
            missing.append(detail_name)
            continue
        if target.IsReadOnly:
            skipped.append(detail_name)
            continue
        value = read_source_param_value(source_elem, source_key)
        if _set_param(target, value):
            written.append(detail_name)
        else:
            skipped.append(detail_name)
    return {"written": written, "missing": missing, "skipped": skipped}


def inject_circuit_values(circuit, detail_item):
    """Copy every CIRCUIT_VALUE_MAP value from a circuit onto a detail item.

    Caller owns the transaction. Returns a dict of parameter-name lists:
    written / missing (param not on the family) / skipped (empty source
    for a numeric target or set failure).
    """
    return _apply_value_map(circuit, detail_item, CIRCUIT_VALUE_MAP)


def inject_panel_values(panel, detail_item):
    """Copy every PANEL_VALUE_MAP value from a panel onto a detail item."""
    return _apply_value_map(panel, detail_item, PANEL_VALUE_MAP)


# Suffix convention for CED shared parameters. The name-match sweep only
# touches these so generic params (Comments, Mark, ...) are never clobbered.
CED_PARAM_SUFFIXES = ("_CED", "_CEDT", "_CEDR", "_CEDI")


def _has_source_param(elem, name):
    try:
        if elem.LookupParameter(name) is not None:
            return True
    except Exception:
        pass
    try:
        type_elem = elem.Document.GetElement(elem.GetTypeId())
        return type_elem is not None and type_elem.LookupParameter(name) is not None
    except Exception:
        return False


def inject_matching_values(source_elem, detail_item, skip_names=None):
    """Name-match sweep: copy every CED-suffixed detail item parameter that
    also exists on the source element (instance first, then type).

    Covers family-specific params the static maps do not know about
    (Enclosure Type_CEDT, Neutral Bus_CED, Feed Thru Lugs_CED, ...).
    Params whose name is in skip_names, or missing from the source element,
    are left untouched. Caller owns the transaction.
    """
    skip = set(skip_names or [])
    written = []
    skipped = []
    for target in detail_item.Parameters:
        try:
            name = target.Definition.Name
        except Exception:
            continue
        if not name or name in skip:
            continue
        if not name.endswith(CED_PARAM_SUFFIXES):
            continue
        if target.IsReadOnly:
            continue
        if not _has_source_param(source_elem, name):
            continue
        value = read_source_param_value(source_elem, name)
        if _set_param(target, value):
            written.append(name)
        else:
            skipped.append(name)
    return {"written": written, "missing": [], "skipped": skipped}
