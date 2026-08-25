# -*- coding: utf-8 -*-
"""Create Circuits by Device Parameter - Revit reads (collect devices + panels, read parameters)."""

import re

from pyrevit import DB

import cg_core
from Snippets import revit_helpers
from Snippets import revit_units

try:
    from CEDElectrical.Infrastructure.Revit.repositories import (
        panel_schedule_repository as _panel_schedule_repo,
    )
except Exception:
    _panel_schedule_repo = None

element_id_value = revit_helpers.get_elementid_value
element_id_from_value = revit_helpers.elementid_from_value


# ---------------------------------------------------------------------------
# Source parameters
# ---------------------------------------------------------------------------
PARAM_CIRCUIT_NUMBER = "CKT_Circuit Number_CEDT"
PARAM_PANEL = "CKT_Panel_CEDT"
PARAM_RATING = "CKT_Rating_CED"
PARAM_LOAD_NAME = "CKT_Load Name_CEDT"
PARAM_SCHEDULE_NOTES = "CKT_Schedule Notes_CEDT"

# The display-only Identity Mark column is assembled from these three values.
# The separator name intentionally supports the project's historic
# "Seperator" spelling as well as the conventional spelling.
PARAM_IDENTITY_TYPE_MARK = "Identity Type Mark"
PARAM_IDENTITY_LABEL_SEPARATOR_NAMES = (
    "Identity Label Seperator",
    "Identity Label Separator",
)
PARAM_IDENTITY_MARK = "Identity Mark"

# >>> CONFIRM THESE TWO <<<  read-only Voltage / Poles columns + mismatch check.
# You picked "specific named params"; defaulting to the generic shared params.
# If the RCD case-controller family stores them under different names, change
# the first entry (the list is tried in order, first hit wins).
VOLTAGE_PARAM_NAMES = ["Voltage", "Voltage_CED", "Voltage Nominal_CED"]
POLES_PARAM_NAMES = ["Number of Poles", "Number of Poles_CED"]

_ELECTRICAL_DATA_VOLTAGE_POLES = re.compile(
    r"(?<![\w.])"
    r"(?P<voltage>[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+))\s*"
    r"(?P<unit>[kKmM]?\s*(?:V|Volt(?:s)?))?\s*/\s*"
    r"(?P<poles>\d+)"
    r"(?!\w)",
    re.IGNORECASE,
)


def _as_text(param):
    if not param or not param.HasValue:
        return ""
    try:
        val = param.AsString()
        if val:
            return val
    except Exception:
        pass
    try:
        val = param.AsValueString()
        return val or ""
    except Exception:
        return ""


def _lookup(elem, name):
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def _first_param(elem, names):
    for name in names:
        p = _lookup(elem, name)
        if p is not None and p.HasValue:
            return p
    return None


def _volts_from_internal(raw):
    """Convert a raw internal electrical-potential value to volts using the
    Forge unit id. Revit's internal unit for ELECTRICAL_POTENTIAL is NOT volts
    (a raw 120 V reads as ~1291), so this conversion is required for the value
    to display / compare as a real nominal voltage."""
    try:
        return revit_units.internal_to_unit(raw, DB.UnitTypeId.Volts)
    except Exception:
        return raw


def _parse_number_token(text):
    """Parse a localized numeric token without performing unit math."""
    value = str(text or "").strip().replace(" ", "")
    if "," in value and "." not in value:
        value = value.replace(",", ".")
    else:
        value = value.replace(",", "")
    try:
        return float(value)
    except Exception:
        return None


def _read_electrical_data(elem, doc):
    """Read voltage/poles from the built-in formatted Electrical Data text.

    This is deliberately a fallback for MEP elements only. The voltage number
    is interpreted in the project's ElectricalPotential display unit and then
    converted through ForgeTypeId APIs to true volts.
    """
    if getattr(elem, "MEPModel", None) is None:
        return None, None
    bip = getattr(DB.BuiltInParameter, "RBS_ELECTRICAL_DATA", None)
    if bip is None:
        return None, None
    try:
        param = elem.get_Parameter(bip)
    except Exception:
        param = None
    source = _as_text(param)
    if not source:
        return None, None
    match = _ELECTRICAL_DATA_VOLTAGE_POLES.search(source)
    if match is None:
        return None, None

    display_voltage = _parse_number_token(match.group("voltage"))
    poles = None
    try:
        poles = int(match.group("poles"))
    except Exception:
        pass
    volts = None
    if display_voltage is not None:
        try:
            volts = revit_units.electrical_potential_display_to_volts(
                doc, display_voltage)
        except Exception:
            volts = None
    return volts, poles


def _parse_leading_number(text):
    """First numeric token in a string ('208 V' -> 208.0), or None."""
    if not text:
        return None
    num = ""
    for ch in str(text).strip():
        if ch.isdigit() or ch in ".-":
            num += ch
        elif num:
            break
    try:
        return float(num) if num not in ("", "-", ".", "-.") else None
    except ValueError:
        return None


def _spec_is_electrical_potential(param):
    """True when the parameter's Forge spec is ElectricalPotential, i.e. its
    AsDouble is an internal value that must be unit-converted. Only these need
    conversion; a plain Number 'Voltage' does not (converting it is exactly what
    produced the bogus '2239 V' readings). Falls back to the pre-2022
    ParameterType enum on older Revit."""
    try:
        defn = param.Definition
    except Exception:
        return False
    try:
        from Autodesk.Revit.DB import SpecTypeId
        return defn.GetDataType() == SpecTypeId.ElectricalPotential
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ParameterType
        return defn.ParameterType == ParameterType.ElectricalPotential
    except Exception:
        return False


def _param_volts(param):
    """Numeric voltage in volts for a parameter, or None.

    Convert from internal units ONLY when the Forge spec is ElectricalPotential;
    a unitless Number/Integer 'Voltage' is taken at face value, and anything
    else is parsed from its formatted value string."""
    try:
        st = param.StorageType
    except Exception:
        st = None
    if st == DB.StorageType.Double:
        try:
            raw = param.AsDouble()
        except Exception:
            raw = None
        if not raw:
            return None
        if _spec_is_electrical_potential(param):
            return _volts_from_internal(raw)
        return raw
    if st == DB.StorageType.Integer:
        try:
            v = param.AsInteger()
        except Exception:
            v = None
        return float(v) if v else None
    try:
        return _parse_leading_number(param.AsValueString() or param.AsString())
    except Exception:
        return None


def _read_voltage(elem, doc=None, electrical_data=None):
    """Return (display_text, integer_key) for the voltage column.

    The key is the voltage snapped to the nearest standard nominal
    (120 / 208 / 240 / 480 ...), converted from internal units only when the
    parameter's spec is ElectricalPotential, so both the column and the
    mismatch flag show a real voltage instead of a raw internal value."""
    p = _first_param(elem, VOLTAGE_PARAM_NAMES)
    key = cg_core.snap_voltage(_param_volts(p)) if p is not None else None
    if key is None and electrical_data is None and doc is not None:
        electrical_data = _read_electrical_data(elem, doc)
    if key is None and electrical_data:
        key = cg_core.snap_voltage(electrical_data[0])
    if key is not None:
        return "{} V".format(key), key
    # Value did not resolve to a recognized nominal - show nothing rather than
    # risk displaying an unconverted internal value.
    return "", None


def _read_poles(elem, doc=None, electrical_data=None):
    """Return (display_text, integer_value) for the poles column."""
    p = _first_param(elem, POLES_PARAM_NAMES)
    if p is None:
        # fall back to the built-in poles parameter (not found by name)
        bip = getattr(DB.BuiltInParameter, "RBS_ELEC_NUMBER_OF_POLES", None)
        if bip is not None:
            try:
                cand = elem.get_Parameter(bip)
                if cand is not None and cand.HasValue:
                    p = cand
            except Exception:
                p = None
    fallback_poles = None
    if electrical_data:
        fallback_poles = electrical_data[1]
    if p is None:
        if fallback_poles is None and doc is not None:
            electrical_data = _read_electrical_data(elem, doc)
            fallback_poles = electrical_data[1] if electrical_data else None
        if fallback_poles:
            return cg_core.poles_label(fallback_poles), fallback_poles
        return "", None
    val = None
    try:
        if p.StorageType == DB.StorageType.Double:
            d = p.AsDouble()
            val = int(round(d)) if d else None
        else:
            val = p.AsInteger()
    except Exception:
        val = None
    if not val:
        # some families type it as text
        digits = "".join(ch for ch in _as_text(p) if ch.isdigit())
        if digits:
            try:
                val = int(digits)
            except ValueError:
                val = None
    if not val and electrical_data is None and doc is not None:
        electrical_data = _read_electrical_data(elem, doc)
        fallback_poles = electrical_data[1] if electrical_data else None
    if not val and fallback_poles:
        val = fallback_poles
    return cg_core.poles_label(val), val


def _read_rating(elem):
    """Return the breaker rating as a number-only string ('20'), '' if unset."""
    p = _lookup(elem, PARAM_RATING)
    if p is None or not p.HasValue:
        return ""
    try:
        amps = p.AsDouble()  # ELECTRICAL_CURRENT internal unit is amps
        if amps:
            return cg_core.format_amps_number(amps)
    except Exception:
        pass
    amps, valid, _ = cg_core.parse_rating(_as_text(p))
    return cg_core.format_amps_number(amps) if valid else ""


def _connector_is_primary(connector):
    """Return whether Revit marks this connector as the family's primary one.

    ``None`` means the runtime/family cannot expose the primary metadata; the
    caller can then use a conservative compatibility fallback.
    """
    try:
        info = connector.GetMEPConnectorInfo()
        if info is None:
            return None
        return bool(info.IsPrimary)
    except Exception:
        return None


def _is_power_connector(connector):
    try:
        return (
            connector.Domain == DB.Domain.DomainElectrical and
            connector.ElectricalSystemType ==
            DB.Electrical.ElectricalSystemType.PowerCircuit
        )
    except Exception:
        return False


def primary_power_connector(elem):
    """Return this element's primary power connector, if it has one.

    Circuit Grouper intentionally never falls back to a secondary connector.
    That keeps multi-connector families from being re-circuited through a
    connector the user did not intend this tool to touch.
    """
    mep = getattr(elem, "MEPModel", None)
    if mep is None:
        return None
    try:
        cm = mep.ConnectorManager
    except Exception:
        cm = None
    if cm is None:
        return None
    fallback = None
    try:
        for conn in cm.Connectors:
            if not _is_power_connector(conn):
                continue
            if fallback is None:
                fallback = conn
            if _connector_is_primary(conn) is True:
                return conn
    except Exception:
        pass
    # Older Revit runtimes and some family connector implementations do not
    # expose GetMEPConnectorInfo/IsPrimary. Keep those circuitable elements in
    # the list rather than silently dropping them; the assigned-system check
    # still prevents re-circuiting an occupied connector.
    return fallback


def _type_element(doc, elem):
    """Return an element's type/symbol, or None when it cannot be resolved."""
    try:
        type_id = elem.GetTypeId()
        if type_id is not None and type_id != DB.ElementId.InvalidElementId:
            return doc.GetElement(type_id)
    except Exception:
        pass
    try:
        return elem.Symbol
    except Exception:
        return None


def _named_value(doc, elem, names, prefer_type=False):
    """Read the first non-blank named value from an instance/type pair.

    The type mark and label separator are normally type parameters but can be
    instance parameters in individual families. Identity Mark is read from
    the instance only.
    """
    type_elem = _type_element(doc, elem)
    sources = [type_elem, elem] if prefer_type else [elem]
    for source in sources:
        if source is None:
            continue
        for name in names:
            value = _as_text(_lookup(source, name)).strip()
            if value:
                return value
    return ""


def _identity_mark_label(doc, elem):
    """Concatenate identity components without inserting any extra spaces."""
    return "".join((
        _named_value(doc, elem, (PARAM_IDENTITY_TYPE_MARK,), prefer_type=True),
        _named_value(doc, elem, PARAM_IDENTITY_LABEL_SEPARATOR_NAMES, prefer_type=True),
        _named_value(doc, elem, (PARAM_IDENTITY_MARK,), prefer_type=False),
    ))


def primary_power_connector_is_unused(connector):
    """Return True only when the specified primary connector has no system.

    The status check is deliberately connector-specific: a system on a
    secondary or non-power connector does not make the primary power connector
    unavailable.  If Revit cannot report the connector's system, skip it
    rather than risk adding it to a second circuit.
    """
    if connector is None:
        return False
    try:
        return connector.MEPSystem is None
    except Exception:
        return False


def _is_already_circuited(elem):
    """True when Revit reports the element as belonging to a power circuit.

    Loads commonly expose their circuit through ``GetAssignedElectricalSystems``;
    panel/distribution equipment can expose it through ``GetElectricalSystems``.
    Check both APIs first, then use the primary connector as a compatibility
    fallback for families that do not expose either MEPModel method.
    """
    mep = getattr(elem, "MEPModel", None)
    if mep is not None:
        for method in ("GetAssignedElectricalSystems", "GetElectricalSystems"):
            fn = getattr(mep, method, None)
            if not callable(fn):
                continue
            try:
                systems = fn() or []
            except Exception:
                continue
            for system in systems:
                try:
                    if system.SystemType == DB.Electrical.ElectricalSystemType.PowerCircuit:
                        return True
                except Exception:
                    # A returned electrical system without a readable type is
                    # safer to treat as assigned than to duplicate.
                    return True

    connector = primary_power_connector(elem)
    return connector is not None and not primary_power_connector_is_unused(
        connector)


def has_power_connector(elem):
    """True when the element has a primary electrical power connector.

    When Revit cannot expose primary metadata, ``primary_power_connector``
    returns the first power connector as a compatibility fallback. This is the
    single definition of "circuitable" used during collection and apply.
    """
    return primary_power_connector(elem) is not None


def _level_name_and_elevation(doc, elem):
    """Return (level_name, level_elevation_ft) for the element's associated
    level, or ('', None) if it has none."""
    lid = getattr(elem, "LevelId", None)
    lvl = None
    try:
        if lid is not None and lid != DB.ElementId.InvalidElementId:
            lvl = doc.GetElement(lid)
    except Exception:
        lvl = None
    if lvl is None:
        for bip_name in ("FAMILY_LEVEL_PARAM",
                         "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM",
                         "INSTANCE_REFERENCE_LEVEL_PARAM",
                         "RBS_START_LEVEL_PARAM"):
            bip = getattr(DB.BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            try:
                p = elem.get_Parameter(bip)
                if p and p.HasValue:
                    cand = doc.GetElement(p.AsElementId())
                    if cand is not None and hasattr(cand, "Elevation"):
                        lvl = cand
                        break
            except Exception:
                continue
    if lvl is None:
        return "", None
    name = ""
    try:
        name = lvl.Name or ""
    except Exception:
        name = ""
    elev = None
    try:
        elev = lvl.Elevation
    except Exception:
        elev = None
    return name, elev


def _element_xyz(elem):
    """(x, y, z) tuple for the element (location point, else bounding-box
    center), or None. Feeds the closest-panel pick, so plan distance matters
    more than exact insertion origin."""
    loc = getattr(elem, "Location", None)
    pt = getattr(loc, "Point", None) if loc is not None else None
    if pt is not None:
        try:
            return (pt.X, pt.Y, pt.Z)
        except Exception:
            pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            return ((bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5)
    except Exception:
        pass
    return None


def _element_z(elem):
    xyz = _element_xyz(elem)
    return xyz[2] if xyz is not None else None


def _format_length(doc, internal_feet):
    """Format an internal (feet) length using the project's display units."""
    try:
        from Autodesk.Revit.DB import UnitFormatUtils, SpecTypeId
        return UnitFormatUtils.Format(
            doc.GetUnits(), SpecTypeId.Length, internal_feet, False, False)
    except Exception:
        pass
    try:  # Revit 2021 and earlier
        from Autodesk.Revit.DB import UnitFormatUtils, UnitType
        return UnitFormatUtils.Format(
            doc.GetUnits(), UnitType.UT_Length, internal_feet, False, False)
    except Exception:
        pass
    return "{:.2f}'".format(internal_feet)


def _read_elevation(doc, elem, level_elev):
    """Elevation of the element above its associated level, in project units."""
    z = _element_z(elem)
    if z is None:
        return ""
    rel = z - (level_elev if level_elev is not None else 0.0)
    return _format_length(doc, rel)


def _param_value_string(p):
    """Best stable string for a parameter's value, for use as a group key."""
    if p is None:
        return ""
    try:
        st = p.StorageType
    except Exception:
        st = None
    try:
        if st == DB.StorageType.String:
            return p.AsString() or ""
        if st == DB.StorageType.Integer:
            vs = p.AsValueString()
            return vs if vs else str(p.AsInteger())
        if st == DB.StorageType.Double:
            vs = p.AsValueString()
            return vs if vs else "{:g}".format(p.AsDouble())
        if st == DB.StorageType.ElementId:
            return p.AsValueString() or ""
    except Exception:
        pass
    try:
        return p.AsValueString() or ""
    except Exception:
        return ""


def _collect_group_values(elem):
    """Map of {instance parameter name: value string} for the element."""
    vals = {}
    try:
        for p in elem.Parameters:
            try:
                name = p.Definition.Name
            except Exception:
                name = None
            if not name:
                continue
            vals[name] = _param_value_string(p)
    except Exception:
        pass
    return vals


def _last_phase(doc):
    """The project's last phase (used to resolve the Space overload), or None."""
    try:
        phases = doc.Phases
        if phases is not None and phases.Size > 0:
            last = phases.Size - 1
            try:
                return phases.get_Item(last)
            except Exception:
                return phases[last]
    except Exception:
        pass
    return None


def _read_space_label(doc, elem):
    """Human label for the Space the element occupies ('101 - Office'), or
    '(No Space)' when it is not inside one.

    Space is a *property* of the element (its physical location inside a Space),
    not a lookup parameter, so it is read from the FamilyInstance.Space
    property, falling back to the phase overload when the direct property is
    null."""
    sp = None
    try:
        sp = getattr(elem, "Space", None)
    except Exception:
        sp = None
    if sp is None:
        get_space = getattr(elem, "get_Space", None)
        if callable(get_space):
            phase = _last_phase(doc)
            try:
                sp = get_space(phase) if phase is not None else None
            except Exception:
                sp = None
    if sp is None:
        return "-"
    number = ""
    name = ""
    try:
        number = (sp.Number or "").strip()
    except Exception:
        number = ""
    try:
        p = sp.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
        if p is not None and p.HasValue:
            name = (p.AsString() or "").strip()
    except Exception:
        name = ""
    if number and name:
        return "{} - {}".format(number, name)
    return number or name or "-"


def _family_type_label(doc, elem):
    try:
        sym = doc.GetElement(elem.GetTypeId())
    except Exception:
        sym = None
    if sym is None:
        try:
            sym = elem.Symbol
        except Exception:
            sym = None
    fam = ""
    typ = ""
    if sym is not None:
        try:
            fam = sym.FamilyName or ""
        except Exception:
            fam = ""
        if not fam:
            try:
                fam = sym.Family.Name or ""
            except Exception:
                fam = ""
        try:
            typ = sym.Name or ""
        except Exception:
            typ = ""
        if not typ:
            bip = getattr(DB.BuiltInParameter, "SYMBOL_NAME_PARAM", None)
            if bip is not None:
                try:
                    typ = _as_text(sym.get_Parameter(bip))
                except Exception:
                    typ = ""
    if fam or typ:
        return "{}: {}".format(fam, typ).strip(": ").strip()
    try:
        return elem.Name or ""
    except Exception:
        return ""


def _resolve_categories(spec):
    cats = []
    for name in spec.get("categories", []):
        bic = getattr(DB.BuiltInCategory, name, None)
        if bic is not None:
            cats.append(bic)
    return cats


def collect_devices(doc, element_ids=None):
    """Return a list of plain dicts, one per circuitable element (any family
    instance with an electrical power connector). No populated-parameter
    requirement - every circuitable element is gathered.

    If ``element_ids`` is given (an iterable of Revit ElementId objects or
    numeric values, e.g. the current selection), only those elements are
    considered; otherwise the whole model is scanned. The window wraps the
    returned dicts into row VMs.
    """
    if element_ids is not None:
        candidates = []
        for identifier in element_ids:
            # Selection.GetElementIds() already supplies real ElementId
            # instances. Only rebuild ids when a numeric data-model value is
            # supplied by a caller.
            if isinstance(identifier, DB.ElementId):
                el = doc.GetElement(identifier)
            else:
                el = doc.GetElement(element_id_from_value(identifier))
            if el is not None:
                candidates.append(el)
    else:
        candidates = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilyInstance)
            .WhereElementIsNotElementType()
        )

    seen = set()
    rows = []
    for elem in candidates:
        if not has_power_connector(elem):
            continue
        eid = element_id_value(elem.Id)
        if eid in seen:
            continue
        seen.add(eid)

        family_type = _family_type_label(doc, elem)
        level_name, level_elev = _level_name_and_elevation(doc, elem)
        space_label = _read_space_label(doc, elem)
        identity_label = _identity_mark_label(doc, elem)
        # Shared parameters are the primary source. Only invoke the formatted
        # Electrical Data parser when one of those primary values is absent or
        # unusable; this keeps the fallback from adding a regex/unit scan to
        # every element in a large model.
        voltage_text, voltage_key = _read_voltage(elem)
        poles_text, poles_value = _read_poles(elem)
        if voltage_key is None or poles_value is None:
            electrical_data = _read_electrical_data(elem, doc)
            if voltage_key is None:
                voltage_text, voltage_key = _read_voltage(
                    elem, electrical_data=electrical_data)
            if poles_value is None:
                poles_text, poles_value = _read_poles(
                    elem, electrical_data=electrical_data)

        # All instance parameters are eligible grouping keys. Family / Type is
        # display-only and is deliberately not injected as a synthetic key.
        group_values = _collect_group_values(elem)
        group_values.pop("Family:Type", None)
        group_values["Level"] = level_name
        # Space is a location property exposed as a synthetic Group By option;
        # store it here so the regular grouping machinery can key on it.
        group_values[cg_core.SPACE_GROUP_KEY] = space_label
        # Identity is a display-only concatenation, stored under its synthetic
        # grouping key so the Group By picker can use the same value.
        group_values[cg_core.IDENTITY_GROUP_KEY] = identity_label

        rows.append({
            "element_id": eid,
            "family_type": family_type,
            "identity_mark": identity_label,
            "panel": _as_text(_lookup(elem, PARAM_PANEL)).strip(),
            "circuit_number": _as_text(_lookup(elem, PARAM_CIRCUIT_NUMBER)).strip(),
            "rating": _read_rating(elem),
            "load_name": _as_text(_lookup(elem, PARAM_LOAD_NAME)).strip(),
            "schedule_notes": _as_text(_lookup(elem, PARAM_SCHEDULE_NOTES)).strip(),
            "voltage_text": voltage_text,
            "voltage_key": voltage_key,
            "poles_text": poles_text,
            "poles_value": poles_value,
            "space": space_label,
            "level": level_name,
            "elevation_text": _read_elevation(doc, elem, level_elev),
            "location": _element_xyz(elem),
            "group_values": group_values,
            "already_circuited": _is_already_circuited(elem),
        })
    rows.sort(key=lambda r: (r["family_type"], r["identity_mark"], r["element_id"]))
    return rows


def _panel_open_slots(elem):
    """Open pole-slots on a panel: 'Max Number of Single Pole Breakers' minus
    the poles consumed by every power circuit already assigned to it (spares
    are circuits, so they count). Returns None when the panel declares no max
    - capacity unknown, treated by the picker as unlimited."""
    max_slots = None
    bip = getattr(DB.BuiltInParameter, "RBS_ELEC_MAX_POLE_BREAKERS", None)
    if bip is not None:
        try:
            p = elem.get_Parameter(bip)
            if p is not None and p.HasValue:
                max_slots = p.AsInteger()
        except Exception:
            max_slots = None
    if not max_slots or max_slots <= 0:
        return None
    used = 0
    mep = getattr(elem, "MEPModel", None)
    fn = getattr(mep, "GetAssignedElectricalSystems", None) if mep is not None else None
    if callable(fn):
        try:
            systems = fn() or []
        except Exception:
            systems = []
        for s in systems:
            try:
                if s.SystemType != DB.Electrical.ElectricalSystemType.PowerCircuit:
                    continue
            except Exception:
                pass
            poles = None
            try:
                poles = s.PolesNumber
            except Exception:
                poles = None
            used += poles if poles else 1
    return max(0, max_slots - used)


def collect_panels(doc):
    """Return (display_names, name_to_id, panel_info) for electrical panels.

    display_names: sorted list of strings for the combo box.
    name_to_id: dict display -> the original Revit ElementId object.
    panel_info: dict display -> {"location": (x,y,z)|None,
        "open_slots": int|None, "profile": distribution-system metadata},
        consumed by panel compatibility filtering and multi-panel auto-pick.
    """
    name_to_id = {}
    panel_info = {}
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)
        .WhereElementIsNotElementType()
    )
    for elem in collector:
        name = ""
        try:
            p = elem.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
            if p and p.HasValue:
                name = p.AsString() or ""
        except Exception:
            name = ""
        if not name:
            try:
                name = elem.Name or ""
            except Exception:
                name = ""
        name = name.strip()
        if not name:
            continue
        # first instance wins on duplicate display names
        if name not in name_to_id:
            # Keep the API object intact. Reconstructing an ElementId later
            # from a Python integer is unnecessary and is not reliable across
            # Revit API versions, especially where ElementId.Value replaced
            # IntegerValue.
            name_to_id[name] = elem.Id
            profile = {}
            if _panel_schedule_repo is not None:
                try:
                    profile = _panel_schedule_repo.get_panel_distribution_profile(
                        doc, elem) or {}
                except Exception:
                    profile = {}
            panel_info[name] = {
                "location": _element_xyz(elem),
                "open_slots": _panel_open_slots(elem),
                "profile": profile,
            }
    return sorted(name_to_id.keys()), name_to_id, panel_info


def collect_target_panel_ids(doc, requested_names):
    """Resolve only the panel ids needed by Run, without UI metadata scans."""
    targets = set(
        str(name or "").strip() for name in list(requested_names or [])
        if str(name or "").strip()
    )
    if not targets:
        return {}
    resolved = {}
    collector = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_ElectricalEquipment)
        .WhereElementIsNotElementType()
    )
    for element in collector:
        name = ""
        try:
            parameter = element.get_Parameter(
                DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
            if parameter and parameter.HasValue:
                name = parameter.AsString() or ""
        except Exception:
            name = ""
        if not name:
            try:
                name = element.Name or ""
            except Exception:
                name = ""
        name = str(name or "").strip()
        if name in targets and name not in resolved:
            resolved[name] = element.Id
            if len(resolved) == len(targets):
                break
    return resolved
