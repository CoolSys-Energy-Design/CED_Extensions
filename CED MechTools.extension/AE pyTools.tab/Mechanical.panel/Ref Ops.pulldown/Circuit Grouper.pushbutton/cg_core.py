# -*- coding: utf-8 -*-
"""Circuit Grouper - pure logic (no Revit API, no .NET).

Kept import-free of Revit/System so the grouping + validation rules can be
unit tested in plain CPython (see _test_cg_core.py).

Expandability: DEVICE_SPECS is the registry of "what can be grouped". Today it
only holds the case-controller spec; add another entry (categories + optional
family substring) to extend the tool to other device families later.
"""

# ---------------------------------------------------------------------------
# Device collection specs (the pluggable scope)
# ---------------------------------------------------------------------------
DEVICE_SPECS = {
    "case_controllers": {
        "label": "Case Controllers",
        "categories": [
            "OST_MechanicalControlDevices",
            "OST_ElectricalEquipment",
            "OST_ElectricalFixtures",
        ],
        "family_contains": None,
    },
}

DEFAULT_SPEC_KEY = "case_controllers"

# ---------------------------------------------------------------------------
# Group-by parameter selection
# ---------------------------------------------------------------------------
# The window lets the user pick which parameter the circuits are grouped on.
# The offered list is the set of parameters COMMON to every gathered fixture
# (so grouping is always well-defined). These names, when present in that
# common set, are surfaced first because they are the usual grouping keys;
# everything else common follows alphabetically.
PREFERRED_GROUP_PARAMS = [
    "CKT_Circuit Number_CEDT",
    "Identity Mark",
    "CKT_Load Name_CEDT",
    "CKT_Panel_CEDT",
    "Family:Type",
    "Level",
]


def common_group_params(rows):
    """Parameter names present on EVERY gathered row, ordered with the usual
    grouping keys first. ``rows`` are the plain dicts from cg_collect, each
    carrying a ``group_values`` dict of {param_name: value_string}."""
    name_sets = [set((r.get("group_values") or {}).keys()) for r in rows]
    if not name_sets:
        return []
    common = name_sets[0]
    for s in name_sets[1:]:
        common = common & s
    preferred = [p for p in PREFERRED_GROUP_PARAMS if p in common]
    rest = sorted(n for n in common if n not in PREFERRED_GROUP_PARAMS)
    return preferred + rest


def default_group_param(options):
    """Pick the initial group-by parameter: prefer circuit number, then
    identity mark, otherwise the first available option."""
    for p in ("CKT_Circuit Number_CEDT", "Identity Mark"):
        if p in options:
            return p
    return options[0] if options else ""

# ---------------------------------------------------------------------------
# Breaker rating options + parsing
# ---------------------------------------------------------------------------
# Combo lists numbers only ("A" is implied).
RATING_OPTIONS = [
    "15", "20", "25", "30", "35", "40", "45", "50", "60", "70", "80", "90",
    "100", "110", "125", "150", "175", "200", "225", "250", "300", "350", "400",
]
DEFAULT_RATING = "20"

STANDARD_AMPS = set([
    15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 100, 110, 125, 150,
    175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800,
])


def parse_rating(text):
    """Parse a breaker rating field.

    Accepts '20', '20 A', '20A', '20amp', '20 amps', '20.0' (case/space
    insensitive). Returns a tuple (amps_float_or_None, valid, standard):
      - amps:     float amperage, or None if not parseable
      - valid:    True only if it parsed to a positive number (garbage -> False)
      - standard: True if it is a standard breaker frame size
    """
    if text is None:
        return None, False, False
    if isinstance(text, (int, float)):
        amps = float(text)
        if amps <= 0:
            return None, False, False
        return amps, True, int(round(amps)) in STANDARD_AMPS

    s = str(text).strip().lower()
    if not s:
        return None, False, False
    for suffix in ("amps", "amp", "a"):
        if s.endswith(suffix):
            s = s[:-len(suffix)].strip()
            break
    try:
        amps = float(s)
    except ValueError:
        return None, False, False
    if amps <= 0:
        return None, False, False
    is_int = abs(amps - round(amps)) < 0.01
    standard = is_int and int(round(amps)) in STANDARD_AMPS
    return amps, True, standard


def format_amps_number(amps):
    """Float amps -> number-only string, e.g. 20.0 -> '20'."""
    if amps is None:
        return ""
    try:
        a = float(amps)
    except (TypeError, ValueError):
        return ""
    if a <= 0:
        return ""
    if abs(a - round(a)) < 0.01:
        return str(int(round(a)))
    return "{:g}".format(a)


def format_amps(amps):
    """Float amps -> display string with unit, e.g. 20.0 -> '20 A'."""
    n = format_amps_number(amps)
    return (n + " A") if n else ""


def poles_label(value):
    try:
        n = int(value)
    except (TypeError, ValueError):
        return ""
    if n <= 0:
        return ""
    return "{}P".format(n)


# ---------------------------------------------------------------------------
# Load-name prepopulation
# ---------------------------------------------------------------------------
def longest_common_prefix(strings):
    items = [s for s in strings if s]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    prefix = items[0]
    for s in items[1:]:
        while prefix and not s.startswith(prefix):
            prefix = prefix[:-1]
        if not prefix:
            return ""
    return prefix


def default_load_name(identity_marks, fallback=""):
    """Common leading characters across the members' Identity Marks, trimmed of
    trailing separators (e.g. ['RA-12A','RA-12B'] -> 'RA-12'). Falls back to the
    given fallback (circuit number) when there is no common stem."""
    marks = [(m or "").strip() for m in identity_marks]
    prefix = longest_common_prefix(marks).rstrip(" -_/.").strip()
    return prefix or (fallback or "")


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def effective_rows(rows):
    """Rows that will actually be circuited: included and not already on a
    circuit."""
    return [
        r for r in rows
        if getattr(r, "include", False) and not getattr(r, "already_circuited", False)
    ]


def validate_members(members):
    """Voltage/pole compatibility problems for a set of (effective) members.
    Empty list == compatible."""
    problems = []
    poles = sorted(set(
        m.poles_value for m in members if getattr(m, "poles_value", None)
    ))
    if len(poles) > 1:
        problems.append("Pole mismatch: {}".format("/".join(poles_label(p) for p in poles)))
    volts = sorted(set(
        m.voltage_key for m in members if getattr(m, "voltage_key", None) is not None
    ))
    if len(volts) > 1:
        problems.append("Voltage mismatch: {}".format("/".join("{}V".format(v) for v in volts)))
    return problems


def build_group_plan(group_key, load_name, panel, rating, members):
    """Build a circuit plan for one group. `members` must already be the
    effective members (included, not already circuited).

    Returns a dict including readiness + the parsed rating. `ready` is True only
    when there are members, no voltage/pole mismatch, and a valid rating.
    """
    amps, rating_valid, rating_standard = parse_rating(rating)
    problems = list(validate_members(members))
    if not rating_valid:
        problems.append("Invalid breaker rating")
    return {
        "group_key": group_key,
        "load_name": (load_name or "").strip() or group_key,
        "panel": (panel or "").strip(),
        "rating_amps": amps,
        "rating_valid": rating_valid,
        "rating_standard": rating_standard,
        "element_ids": [int(getattr(m, "element_id")) for m in members],
        "problems": problems,
        "ready": len(problems) == 0 and len(members) > 0,
    }


def next_new_group_key(existing_keys):
    existing = set(str(k) for k in existing_keys)
    n = 1
    while True:
        candidate = "NEW-{}".format(n)
        if candidate not in existing:
            return candidate
        n += 1
