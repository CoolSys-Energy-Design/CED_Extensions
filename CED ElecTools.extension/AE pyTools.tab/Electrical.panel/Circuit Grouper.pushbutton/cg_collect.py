# -*- coding: utf-8 -*-
"""Circuit Grouper - Revit reads (collect devices + panels, read parameters)."""

from pyrevit import DB

import cg_core


def element_id_value(eid):
    """Numeric value of an ``ElementId``, compatible across Revit versions.

    Revit 2024+ exposes ``ElementId.Value`` (Int64) and deprecates
    ``IntegerValue`` (which is removed in newer builds, so it can't be
    relied on for 2025 / 2026). Prefer ``Value`` and fall back to
    ``IntegerValue`` on pre-2024 builds. Returns ``None`` for a missing
    id.
    """
    if eid is None:
        return None
    v = getattr(eid, "Value", None)
    if v is None:
        v = getattr(eid, "IntegerValue", None)
    return v


# ---------------------------------------------------------------------------
# Source parameters
# ---------------------------------------------------------------------------
PARAM_CIRCUIT_NUMBER = "CKT_Circuit Number_CEDT"
PARAM_PANEL = "CKT_Panel_CEDT"
PARAM_RATING = "CKT_Rating_CED"
PARAM_LOAD_NAME = "CKT_Load Name_CEDT"

# >>> CONFIRM THESE TWO <<<  read-only Voltage / Poles columns + mismatch check.
# You picked "specific named params"; defaulting to the generic shared params.
# If the RCD case-controller family stores them under different names, change
# the first entry (the list is tried in order, first hit wins).
VOLTAGE_PARAM_NAMES = ["Voltage", "Voltage_CED", "Voltage Nominal_CED"]
POLES_PARAM_NAMES = ["Number of Poles", "Number of Poles_CED"]


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
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertFromInternalUnits(raw, UnitTypeId.Volts)
    except Exception:
        pass
    try:  # Revit 2021 and earlier (pre-ForgeTypeId)
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertFromInternalUnits(raw, DisplayUnitType.DUT_VOLTS)
    except Exception:
        return raw


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


def _read_voltage(elem):
    """Return (display_text, integer_key) for the voltage column.

    The key is the voltage snapped to the nearest standard nominal
    (120 / 208 / 240 / 480 ...), converted from internal units only when the
    parameter's spec is ElectricalPotential, so both the column and the
    mismatch flag show a real voltage instead of a raw internal value."""
    p = _first_param(elem, VOLTAGE_PARAM_NAMES)
    if p is None:
        return "", None
    key = cg_core.snap_voltage(_param_volts(p))
    if key is not None:
        return "{} V".format(key), key
    # Value did not resolve to a recognized nominal - show nothing rather than
    # risk displaying an unconverted internal value.
    return "", None


def _read_poles(elem):
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
    if p is None:
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


def _is_already_circuited(elem):
    """True if the element already belongs to a power circuit.

    Checks BOTH GetAssignedElectricalSystems and GetElectricalSystems: a load
    reports its circuit through the former, but panelboards / distribution
    equipment report theirs only through the latter. Either way the element is
    not a valid new load and must not be re-circuited."""
    mep = getattr(elem, "MEPModel", None)
    if mep is None:
        return False
    for method in ("GetAssignedElectricalSystems", "GetElectricalSystems"):
        fn = getattr(mep, method, None)
        if not callable(fn):
            continue
        try:
            systems = fn()
        except Exception:
            continue
        if not systems:
            continue
        for s in systems:
            try:
                if s.SystemType == DB.Electrical.ElectricalSystemType.PowerCircuit:
                    return True
            except Exception:
                return True
    return False


def has_power_connector(elem):
    """True if the element has an electrical power connector (i.e. it can be
    placed on a power circuit). This is the single definition of "circuitable"
    used both to gather candidates here and to filter members in cg_apply."""
    mep = getattr(elem, "MEPModel", None)
    if mep is None:
        return False
    try:
        cm = mep.ConnectorManager
    except Exception:
        cm = None
    if cm is None:
        return False
    try:
        for conn in cm.Connectors:
            try:
                if conn.Domain == DB.Domain.DomainElectrical:
                    if conn.ElectricalSystemType == DB.Electrical.ElectricalSystemType.PowerCircuit:
                        return True
            except Exception:
                continue
    except Exception:
        return False
    return False


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


def _element_z(elem):
    loc = getattr(elem, "Location", None)
    pt = getattr(loc, "Point", None) if loc is not None else None
    if pt is not None:
        try:
            return pt.Z
        except Exception:
            pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min.Z + bb.Max.Z) * 0.5
    except Exception:
        pass
    return None


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
        return "(No Space)"
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
    return number or name or "(No Space)"


def _family_type_label(doc, elem):
    try:
        sym = doc.GetElement(elem.GetTypeId())
    except Exception:
        sym = None
    fam = ""
    typ = ""
    if sym is not None:
        try:
            fam = sym.FamilyName or ""
        except Exception:
            fam = ""
        try:
            typ = sym.Name or ""
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

    If ``element_ids`` is given (an iterable of numeric ElementIds, e.g. the
    current selection), only those elements are considered; otherwise the
    whole model is scanned. The window wraps the returned dicts into row VMs.
    """
    if element_ids is not None:
        wanted = set(int(e) for e in element_ids)
        candidates = []
        for eid in wanted:
            el = doc.GetElement(DB.ElementId(int(eid)))
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
        voltage_text, voltage_key = _read_voltage(elem)
        poles_text, poles_value = _read_poles(elem)

        # all instance params, plus synthetic keys that are always present so
        # the user can always group by Family:Type / Level
        group_values = _collect_group_values(elem)
        group_values["Family:Type"] = family_type
        group_values["Level"] = level_name
        # Space is a location property, offered via its own control (see
        # cg_core.SPACE_GROUP_KEY); stored here so grouping can key on it.
        group_values[cg_core.SPACE_GROUP_KEY] = _read_space_label(doc, elem)

        rows.append({
            "element_id": eid,
            "family_type": family_type,
            "identity_mark": _as_text(_lookup(elem, "Identity Mark")).strip(),
            "panel": _as_text(_lookup(elem, PARAM_PANEL)).strip(),
            "circuit_number": _as_text(_lookup(elem, PARAM_CIRCUIT_NUMBER)).strip(),
            "rating": _read_rating(elem),
            "load_name": _as_text(_lookup(elem, PARAM_LOAD_NAME)).strip(),
            "voltage_text": voltage_text,
            "voltage_key": voltage_key,
            "poles_text": poles_text,
            "poles_value": poles_value,
            "elevation_text": _read_elevation(doc, elem, level_elev),
            "group_values": group_values,
            "already_circuited": _is_already_circuited(elem),
        })
    rows.sort(key=lambda r: (r["family_type"], r["identity_mark"], r["element_id"]))
    return rows


def collect_panels(doc):
    """Return (display_names, name_to_id) for electrical panels.

    display_names: sorted list of strings for the combo box.
    name_to_id: dict display -> ElementId.IntegerValue.
    """
    name_to_id = {}
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
        name_to_id.setdefault(name, element_id_value(elem.Id))
    return sorted(name_to_id.keys()), name_to_id
