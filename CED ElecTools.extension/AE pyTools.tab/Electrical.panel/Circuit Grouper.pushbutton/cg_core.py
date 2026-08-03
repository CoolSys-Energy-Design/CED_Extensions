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
# Group-by Space (a location property, not a parameter)
# ---------------------------------------------------------------------------
# Space is not a LookupParameter - it is the Space the element physically sits
# in. cg_collect stores a per-element Space label under this synthetic key so
# the same grouping machinery can group on it, but the window offers it through
# a SEPARATE control (checkbox) rather than the parameter combo, so it is
# excluded from the parameter option list.
SPACE_GROUP_KEY = "Space"

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
    # Space is offered through its own control, never the parameter combo.
    common = common - set([SPACE_GROUP_KEY])
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
# Name-by parameter selection
# ---------------------------------------------------------------------------
# Independently of the group-by choice, the window lets the user pick which
# parameter SEEDS each circuit's Load Name (the user can still hand-edit any
# name after). The option list is the same common-parameter set as group-by.
def default_name_param(options):
    """Pick the initial name-by parameter: prefer the existing load name, then
    identity mark, otherwise the first available option."""
    for p in ("CKT_Load Name_CEDT", "Identity Mark"):
        if p in options:
            return p
    return options[0] if options else ""


def name_from_values(values, fallback=""):
    """Circuit name derived from the members' values of the name-by parameter:
    the shared value when all non-blank values agree, else their common leading
    characters trimmed of trailing separators, else the fallback."""
    vals = [(v or "").strip() for v in values]
    nonblank = [v for v in vals if v]
    if not nonblank:
        return fallback or ""
    if len(set(nonblank)) == 1:
        return nonblank[0]
    prefix = longest_common_prefix(nonblank).rstrip(" -_/.").strip()
    return prefix or (fallback or "")

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


# ---------------------------------------------------------------------------
# Voltage normalization
# ---------------------------------------------------------------------------
# Revit's internal unit for electrical potential is NOT volts, so a raw
# AsDouble() reads as a large number (~10.764x the real voltage, e.g. ~1291 for
# 120 V, ~2239 for 208 V). The caller converts to volts first (via the Forge
# unit id when the spec is ElectricalPotential); snap_voltage then accepts the
# value ONLY if it lands on a recognized nominal. This is the guard that keeps
# an unconverted value from ever being displayed: 1291 / 2239 match no nominal,
# so they are rejected (None) rather than shown.
STANDARD_VOLTAGES = [24, 48, 120, 208, 240, 277, 347, 480, 600]


def snap_voltage(volts, tolerance=8.0):
    """Snap a measured (already volt-converted) voltage to the nearest standard
    nominal when within ``tolerance``. Returns None for missing / non-positive
    input OR for any value that does not resolve to a nominal - so unconverted
    internal readings (1291, 2239, ...) are never emitted."""
    if volts is None:
        return None
    try:
        v = float(volts)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    best = min(STANDARD_VOLTAGES, key=lambda s: abs(s - v))
    if abs(best - v) <= tolerance:
        return best
    return None


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
# Multi-panel candidate resolution ("RA, RB, RC, RD" -> pick one)
# ---------------------------------------------------------------------------
# Some fixtures arrive with SEVERAL allowed panels listed in CKT_Panel_CEDT.
# The tool then pre-selects the closest listed panel that still has room for
# the circuit; when none has room the combo is left blank and the group gets a
# validation note (the user picks manually). A single-name value is untouched.
PANEL_NOTE_NO_ROOM = "no listed panel has room - pick a panel manually"
PANEL_NOTE_UNKNOWN = "none of the listed panels exist in the model - pick a panel manually"


def parse_panel_candidates(text, known_panels=None):
    """Split a panel value like 'RA, RB, RC' into candidate names.

    Splits on comma / semicolon / slash, trims, and dedupes preserving order.
    When ``known_panels`` is given, candidates are matched case-insensitively
    against it and returned in the model's canonical spelling; names that
    match no known panel are dropped.
    """
    if not text:
        return []
    s = str(text).replace(";", ",").replace("/", ",")
    canon = None
    if known_panels is not None:
        canon = {}
        for k in known_panels:
            canon.setdefault(str(k).strip().lower(), str(k).strip())
    seen = set()
    out = []
    for token in s.split(","):
        name = token.strip()
        if not name:
            continue
        if canon is not None:
            name = canon.get(name.lower())
            if name is None:
                continue
        key = name.lower()
        if key not in seen:
            seen.add(key)
            out.append(name)
    return out


def centroid(points):
    """Average of the given (x, y, z) tuples, ignoring Nones. None if empty."""
    pts = [p for p in points if p is not None]
    if not pts:
        return None
    n = float(len(pts))
    return (sum(p[0] for p in pts) / n,
            sum(p[1] for p in pts) / n,
            sum(p[2] for p in pts) / n)


def distance(a, b):
    """Euclidean distance between two (x, y, z) tuples, or None if either is
    missing."""
    if a is None or b is None:
        return None
    return ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2) ** 0.5


def resolve_panel_assignments(requests, panel_info):
    """Choose a panel for each multi-candidate group: the CLOSEST listed panel
    that still has room for the circuit's poles.

    requests: list of dicts
        {"group_key": str, "candidates": [panel names], "centroid": xyz|None,
         "poles": int|None}
    panel_info: {panel_name: {"location": xyz|None, "open_slots": int|None}}
        open_slots None means capacity unknown -> treated as unlimited and
        never reserved against.

    Slots are RESERVED as groups claim them, so several groups sharing the
    same candidate list overflow onto the next-closest panel instead of all
    piling onto one. Groups are processed nearest-first (whoever sits closest
    to its best panel claims it first); groups with no measurable distance go
    last and take their candidates in listed order.

    Returns {group_key: {"panel": name or None, "note": ""}}; ``note`` is a
    PANEL_NOTE_* string when no panel could be chosen.
    """
    remaining = {}
    for name, info in (panel_info or {}).items():
        remaining[name] = (info or {}).get("open_slots", None)

    def _dist(req, name):
        info = (panel_info or {}).get(name) or {}
        return distance(req.get("centroid"), info.get("location"))

    def _best_distance(req):
        ds = [_dist(req, c) for c in req.get("candidates", [])]
        ds = [d for d in ds if d is not None]
        return min(ds) if ds else None

    # nearest-first claim order; unmeasurable distances last, stable by key
    order = sorted(
        range(len(requests)),
        key=lambda i: (
            _best_distance(requests[i]) is None,
            _best_distance(requests[i]) or 0.0,
            str(requests[i].get("group_key", "")),
        ),
    )

    results = {}
    for i in order:
        req = requests[i]
        key = req.get("group_key", "")
        candidates = [c for c in req.get("candidates", []) if c in remaining]
        if not candidates:
            results[key] = {"panel": None, "note": PANEL_NOTE_UNKNOWN}
            continue
        poles = req.get("poles") or 1
        # candidates by distance (unknown distances keep listed order, last)
        ranked = sorted(
            enumerate(candidates),
            key=lambda pair: (
                _dist(req, pair[1]) is None,
                _dist(req, pair[1]) or 0.0,
                pair[0],
            ),
        )
        chosen = None
        for _, name in ranked:
            slots = remaining.get(name)
            if slots is None or slots >= poles:
                chosen = name
                break
        if chosen is None:
            results[key] = {"panel": None, "note": PANEL_NOTE_NO_ROOM}
            continue
        if remaining.get(chosen) is not None:
            remaining[chosen] -= poles
        results[key] = {"panel": chosen, "note": ""}
    return results


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
