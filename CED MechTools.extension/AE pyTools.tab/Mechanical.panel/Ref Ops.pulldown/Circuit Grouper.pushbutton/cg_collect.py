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
PARAM_CIRCUIT_NUMBER = "CKT_Circuit Number_CEDT"   # grouping key (must be populated)
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


def _read_voltage(elem):
    """Return (display_text, integer_key) for the voltage column."""
    p = _first_param(elem, VOLTAGE_PARAM_NAMES)
    if p is None:
        return "", None
    text = _as_text(p)
    key = None
    try:
        raw = p.AsDouble()
        # Revit stores ELECTRICAL_POTENTIAL internally in volts.
        if raw:
            key = int(round(raw))
    except Exception:
        key = None
    if not text and key is not None:
        text = "{} V".format(key)
    return text, key


def _read_poles(elem):
    """Return (display_text, integer_value) for the poles column."""
    p = _first_param(elem, POLES_PARAM_NAMES)
    if p is None:
        return "", None
    val = None
    try:
        val = p.AsInteger()
    except Exception:
        val = None
    if not val:
        # some families type it as text
        txt = _as_text(p)
        digits = "".join(ch for ch in txt if ch.isdigit())
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


def collect_devices(doc, spec_key=None):
    """Return a list of plain dicts, one per circuitable device that has
    CKT_Circuit Number_CEDT populated. The window wraps these into row VMs."""
    spec = cg_core.DEVICE_SPECS[spec_key or cg_core.DEFAULT_SPEC_KEY]
    fam_filter = (spec.get("family_contains") or "").strip().lower() or None

    seen = set()
    rows = []
    for bic in _resolve_categories(spec):
        collector = (
            DB.FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
        )
        for elem in collector:
            eid = element_id_value(elem.Id)
            if eid in seen:
                continue

            ckt = _lookup(elem, PARAM_CIRCUIT_NUMBER)
            circuit_number = _as_text(ckt).strip()
            if not circuit_number:
                continue  # require a populated circuit number

            identity_mark = _as_text(_lookup(elem, "Identity Mark")).strip()
            if not identity_mark:
                continue  # require the Identity Mark parameter (no Mark fallback)

            family_type = _family_type_label(doc, elem)
            if fam_filter and fam_filter not in family_type.lower():
                continue

            seen.add(eid)
            voltage_text, voltage_key = _read_voltage(elem)
            poles_text, poles_value = _read_poles(elem)
            rows.append({
                "element_id": eid,
                "family_type": family_type,
                "identity_mark": identity_mark,
                "panel": _as_text(_lookup(elem, PARAM_PANEL)).strip(),
                "circuit_number": circuit_number,
                "rating": _read_rating(elem),
                "load_name": _as_text(_lookup(elem, PARAM_LOAD_NAME)).strip(),
                "voltage_text": voltage_text,
                "voltage_key": voltage_key,
                "poles_text": poles_text,
                "poles_value": poles_value,
                "already_circuited": _is_already_circuited(elem),
            })
    rows.sort(key=lambda r: (r["circuit_number"], r["family_type"], r["identity_mark"]))
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
