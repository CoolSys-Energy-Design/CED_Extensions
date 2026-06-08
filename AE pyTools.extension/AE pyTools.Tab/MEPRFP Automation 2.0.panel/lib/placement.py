# -*- coding: utf-8 -*-
"""
Placement engine.

Replaces the six legacy placement tools (PlaceLinkedElements + 2 filtered
variants, PlaceCADElements + 2 filtered variants). Single entry point:

    targets = collect_targets(...)
    matches = match_targets(targets, profiles, mode)
    result = execute_placement(doc, matches, options)

Three target sources:
    * Linked Revit model — every FamilyInstance / Group in a chosen
      RevitLinkInstance, with parent_filter-style matching.
    * CSV file — rebased decimal coordinates produced upstream.
    * DWG link — best-effort block-instance walk (block names may be
      partially or fully unavailable depending on the DWG; the
      collector returns whatever it could extract).

Matching modes:
    * ``family_name_strip_suffix`` — for linked Revit. Strips trailing
      ``_NNN`` from both sides and compares family names case-insensitively.
    * ``cad_aliases`` — for CSV / DWG. Each profile may declare a
      comma-separated alias list under ``equipment_properties.cad_aliases``;
      a target whose name matches any alias is placed against that profile.
      The profile's own ``name`` (plus ``parent_filter.family_name_pattern``
      and any ``merged_aliases``) is treated as an implicit alias, so a YAML
      that omits ``cad_aliases`` still matches when the CSV ``Name`` equals
      the profile name. Same trailing-``_NNN`` strip applies.
"""

import io
import math
import os
import re

import clr  # noqa: F401

from System import Enum  # noqa: E402

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    Group,
    GroupType,
    ImportInstance,
    Level,
    LocationPoint,
    Options,
    Phase,
    RevitLinkInstance,
    SubTransaction,
    Transform,
    ViewPlan,
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType  # noqa: E402

import directives as _dir
import element_linker as _el
import element_linker_io as _el_io
import geometry
import hosted_annotations
import links


# ---------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------

CAD_ALIASES_KEY = "cad_aliases"

MATCH_FAMILY_NAME_STRIP_SUFFIX = "family_name_strip_suffix"
MATCH_CAD_ALIASES = "cad_aliases"

SOURCE_LINKED_REVIT = "linked_revit"
SOURCE_HOST_MODEL = "host_model"
SOURCE_CSV = "csv"
SOURCE_DWG_LINK = "dwg_link"
SOURCE_PICKED_POINT = "picked_point"


# ---------------------------------------------------------------------
# Pure-logic helpers (testable offline)
# ---------------------------------------------------------------------

_TRAILING_SUFFIX_RE = re.compile(r"_\d+$")


def strip_trailing_suffix(name):
    """Drop a trailing ``_NNN`` (one or more digits) from ``name``.

    ``"AC_BLOCK"`` -> ``"AC_BLOCK"``
    ``"AC_BLOCK_1"`` -> ``"AC_BLOCK"``
    ``"AC_BLOCK_42"`` -> ``"AC_BLOCK"``
    ``"AC_BLOCK_2A"`` -> ``"AC_BLOCK_2A"`` (only stripped if all digits)
    """
    if not name:
        return ""
    return _TRAILING_SUFFIX_RE.sub("", str(name))


def normalize_name(name):
    """Lowercase + strip trailing ``_NNN`` for case-insensitive matching.

    Whitespace is trimmed *before* the suffix strip so that a trailing
    space doesn't hide the suffix from the regex.
    """
    return strip_trailing_suffix((name or "").strip()).lower()


def collect_profile_aliases(profile):
    """Return the set of CAD aliases declared on ``profile``.

    Reads ``equipment_properties[cad_aliases]``. Value can be a list of
    strings or a single comma-separated string. Empty / missing -> set().
    """
    if not isinstance(profile, dict):
        return set()
    props = profile.get("equipment_properties") or {}
    if not isinstance(props, dict):
        return set()
    raw = props.get(CAD_ALIASES_KEY)
    if raw is None:
        return set()
    items = []
    if isinstance(raw, list):
        items = [str(x) for x in raw if x is not None]
    elif isinstance(raw, str):
        items = [s for s in raw.split(",")]
    else:
        items = [str(raw)]
    out = set()
    for item in items:
        norm = normalize_name(item)
        if norm:
            out.add(norm)
    return out


def profile_family_names(profile):
    """Return the set of family-name keys we'll match against for linked
    Revit. Pulls from:
        * ``parent_filter.family_name_pattern``
        * the profile's own name (``"Family : Type"`` -> ``"Family"``)
        * every entry in ``merged_aliases`` (split + family part)

    All keys are normalised through ``normalize_name`` (lowercase +
    suffix-strip), so the family-name match is case-insensitive and
    handles trailing ``_NNN`` revisions automatically.
    """
    if not isinstance(profile, dict):
        return set()
    out = set()

    def _add_family(value):
        if not value:
            return
        if " : " in value:
            family_part, _ = value.split(" : ", 1)
        else:
            family_part = value
        key = normalize_name(family_part)
        if key:
            out.add(key)

    pf = profile.get("parent_filter") or {}
    if isinstance(pf, dict):
        _add_family(pf.get("family_name_pattern"))
    _add_family(profile.get("name") or "")
    for alias in profile.get("merged_aliases") or []:
        if isinstance(alias, str):
            _add_family(alias)
    return {n for n in out if n}


def profile_family_names_raw(profile):
    """Like ``profile_family_names`` but returns family strings with
    only case-folding applied — *no* trailing ``_NNN`` strip. Used to
    detect exact 1:1 alignments between a target and a profile so the
    matcher can prefer them over the suffix-stripped fallback.
    """
    if not isinstance(profile, dict):
        return set()
    out = set()

    def _add(value):
        if not value:
            return
        if " : " in value:
            family_part, _ = value.split(" : ", 1)
        else:
            family_part = value
        key = (family_part or "").strip().lower()
        if key:
            out.add(key)

    pf = profile.get("parent_filter") or {}
    if isinstance(pf, dict):
        _add(pf.get("family_name_pattern"))
    _add(profile.get("name") or "")
    for alias in profile.get("merged_aliases") or []:
        if isinstance(alias, str):
            _add(alias)
    return out


def collect_profile_aliases_raw(profile):
    """Like ``collect_profile_aliases`` but case-folded only, no suffix
    strip. Used by the strict match tier for CAD blocks."""
    if not isinstance(profile, dict):
        return set()
    props = profile.get("equipment_properties") or {}
    if not isinstance(props, dict):
        return set()
    raw = props.get(CAD_ALIASES_KEY)
    if raw is None:
        return set()
    items = []
    if isinstance(raw, list):
        items = [str(x) for x in raw if x is not None]
    elif isinstance(raw, str):
        items = [s for s in raw.split(",")]
    else:
        items = [str(raw)]
    out = set()
    for item in items:
        norm = (item or "").strip().lower()
        if norm:
            out.add(norm)
    return out


def _profile_id_label(profile):
    return "{}  ({})".format(
        profile.get("name") or "(unnamed)",
        profile.get("id") or "?",
    )


# ---------------------------------------------------------------------
# Target dataclass-equivalent (kept Py2/3 compatible)
# ---------------------------------------------------------------------

class Target(object):
    """One anchor location for placement.

    ``source``      one of SOURCE_*
    ``name``        block name / family name / csv name
    ``world_pt``    (x, y, z) in the host doc's coordinate frame, feet
    ``rotation_deg`` rotation around Z, host frame, degrees
    ``level_id``    ElementId.Value (int) for placement level, or None
    ``link_inst``   RevitLinkInstance if linked-Revit, else None
    ``link_elem_id`` int ElementId of the linked element, else None
    """

    __slots__ = (
        "source", "name", "world_pt", "rotation_deg",
        "level_id", "link_inst", "link_elem_id",
    )

    def __init__(self, source, name, world_pt, rotation_deg=0.0,
                 level_id=None, link_inst=None, link_elem_id=None):
        self.source = source
        self.name = name or ""
        self.world_pt = tuple(world_pt) if world_pt else (0.0, 0.0, 0.0)
        self.rotation_deg = float(rotation_deg or 0.0)
        self.level_id = level_id
        self.link_inst = link_inst
        self.link_elem_id = link_elem_id

    def __repr__(self):
        return "Target(source={}, name={!r}, world_pt={})".format(
            self.source, self.name, self.world_pt
        )


class Match(object):
    """One (target, profile) pair the user can choose to place or skip."""

    __slots__ = ("target", "profile", "skip")

    def __init__(self, target, profile, skip=False):
        self.target = target
        self.profile = profile
        self.skip = skip


# ---------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------

def _match_one_linked_revit(target, profiles):
    """A linked-Revit target's ``name`` is the element's family name.

    Two-tier match. Tier 1 looks for profiles whose raw family name is
    a case-insensitive *exact* match to the target name (suffix included).
    If any tier-1 hits exist they win, and the suffix-stripped fallback
    is skipped — that's how the engine automatically resolves the
    "three Stinger Cart_1/2/3 profiles vs. three Stinger Cart_1/2/3
    targets" case 1:1 instead of producing a 3×3 cross-product.

    Tier 2 is the legacy suffix-stripped match, used only when no
    profile aligns exactly with the target name.
    """
    target_name_lower = (target.name or "").strip().lower()
    target_key = normalize_name(target.name)
    if not target_key:
        return []

    if target_name_lower:
        strict = [
            p for p in profiles
            if target_name_lower in profile_family_names_raw(p)
        ]
        if strict:
            return strict

    return [
        p for p in profiles
        if target_key in profile_family_names(p)
    ]


def _match_one_cad(target, profiles):
    """A CAD target's ``name`` is the block name. Same two-tier rule as
    linked-Revit: prefer exact-name aliases over suffix-stripped ones.

    Alias source is the union of explicit ``equipment_properties.cad_aliases``
    and the implicit aliases derived from the profile itself (its ``name``,
    ``parent_filter.family_name_pattern``, and ``merged_aliases``) — same
    set the linked-Revit matcher uses. That way a YAML with no
    ``cad_aliases`` declared still matches a CSV ``Name`` that equals the
    profile name.
    """
    target_name_lower = (target.name or "").strip().lower()
    target_key = normalize_name(target.name)
    if not target_key:
        return []

    if target_name_lower:
        strict = [
            p for p in profiles
            if target_name_lower in collect_profile_aliases_raw(p)
            or target_name_lower in profile_family_names_raw(p)
        ]
        if strict:
            return strict

    return [
        p for p in profiles
        if target_key in collect_profile_aliases(p)
        or target_key in profile_family_names(p)
    ]


def match_targets(targets, profiles, mode):
    """Cross-product match: every (target, profile) pair where the
    matcher accepts becomes a ``Match``. One target may match multiple
    profiles (via overlapping aliases / multiple family-name keys)."""
    matcher = (
        _match_one_linked_revit if mode == MATCH_FAMILY_NAME_STRIP_SUFFIX
        else _match_one_cad
    )
    out = []
    for t in targets:
        for p in matcher(t, profiles):
            out.append(Match(t, p))
    return out


def dedupe_matches_per_target(matches):
    """Collapse the cross-product to **one match per target anchor**.

    Used to suppress legacy duplicates: when several profiles share a
    family name (because the same family was captured/merged multiple
    times), a single CAD/linked target ends up with N matches and the
    placement engine stacks N fixtures on the same anchor. This filter
    keeps exactly one match per target, choosing by:

        1. Profile whose ``parent_filter.family_name_pattern`` exactly
           equals the target name (case-insensitive). Most specific.
        2. Profile that has the most LEDs (richest data).
        3. Lowest profile id alphabetically — deterministic tie-break.

    Returns a new list; input is not mutated.
    """
    if not matches:
        return []

    def _target_key(target):
        wp = target.world_pt or (0.0, 0.0, 0.0)
        return (
            round(float(wp[0]), 3),
            round(float(wp[1]), 3),
            round(float(wp[2]), 3),
            (target.name or "").strip().lower(),
        )

    def _led_count(profile):
        n = 0
        for s in profile.get("linked_sets") or []:
            if isinstance(s, dict):
                n += len(s.get("linked_element_definitions") or [])
        return n

    bucketed = {}
    order = []
    for m in matches:
        k = _target_key(m.target)
        if k not in bucketed:
            bucketed[k] = []
            order.append(k)
        bucketed[k].append(m)

    out = []
    for k in order:
        group = bucketed[k]
        if len(group) == 1:
            out.append(group[0])
            continue
        target_name = group[0].target.name or ""
        target_name_lower = target_name.strip().lower()

        exact = [
            m for m in group
            if ((m.profile.get("parent_filter") or {}).get("family_name_pattern") or "")
                .strip().lower() == target_name_lower
        ]
        candidates = exact if exact else group

        candidates = sorted(
            candidates,
            key=lambda m: (
                -_led_count(m.profile),       # more LEDs first
                m.profile.get("id") or "",    # then alphabetical id
            ),
        )
        out.append(candidates[0])
    return out


# ---------------------------------------------------------------------
# Skip-already-placed
# ---------------------------------------------------------------------

def _placed_led_anchor_signatures(doc):
    """Return a set of ``(led_id, anchor_x_rounded, anchor_y_rounded)``
    for every placed fixture in the active document.

    Used by ``execute_placement`` to skip re-placing the **same LED**
    onto the **same anchor**. The key is per-LED, not per-set, so a
    later run that selects a new category (e.g. Data Devices) can land
    its LEDs at anchors where other LEDs from the same set were placed
    in a prior run — the bug being fixed here is that the old per-set
    key skipped the entire match as soon as any LED from that set
    existed at the anchor.
    """
    signatures = set()
    for klass in (FamilyInstance, Group):
        collector = FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType()
        for elem in collector:
            linker = _el_io.read_from_element(elem)
            if linker is None:
                continue
            led_id = getattr(linker, "led_id", None)
            if not led_id:
                continue
            anchor_pt = linker.parent_location_ft
            if anchor_pt and len(anchor_pt) >= 2:
                key = (
                    led_id,
                    round(float(anchor_pt[0]), 1),
                    round(float(anchor_pt[1]), 1),
                )
                signatures.add(key)
    return signatures


# ---------------------------------------------------------------------
# Target collection: linked Revit
# ---------------------------------------------------------------------

def collect_linked_revit_link_instances(doc):
    """Return ``[RevitLinkInstance, ...]`` for every loaded Revit link."""
    out = []
    for link_inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        if link_inst is None:
            continue
        try:
            if link_inst.GetLinkDocument() is None:
                continue
        except Exception:
            continue
        out.append(link_inst)
    return out


def _build_host_level_index(host_doc):
    """Return ``[(elevation_ft, level_id_int), ...]`` sorted by elevation
    for every Level in ``host_doc``. Used to map a linked element's
    elevation to the host doc's nearest Level by elevation rather than
    by name — so the placement engine respects elevation alignment even
    when the host and link use different naming conventions
    (e.g. ``L1`` vs ``LEVEL 1`` vs ``Finished Floor``).
    """
    out = []
    if host_doc is None:
        return out
    try:
        collector = FilteredElementCollector(host_doc).OfClass(Level)
    except Exception:
        return out
    for lv in collector:
        try:
            eid = lv.Id
            val = getattr(eid, "Value", None)
            if val is None:
                val = getattr(eid, "IntegerValue", None)
            if val is None:
                continue
            out.append((float(lv.Elevation), val))
        except Exception:
            continue
    out.sort()
    return out


def _resolve_level_by_elevation(elevation_ft, host_index, tolerance_ft=1.0):
    """Return the host-doc Level ElementId integer whose elevation is
    closest to ``elevation_ft``, provided the gap is within
    ``tolerance_ft``. Returns ``None`` when ``host_index`` is empty or
    no level is within tolerance — caller falls back to
    ``options.default_level_id``.

    Default tolerance is 1 ft: typical floor-to-floor is 10-14 ft, so
    a 1-ft window safely identifies the same level across host/link
    pairings without absorbing the adjacent floor.
    """
    if not host_index:
        return None
    best = min(host_index, key=lambda pair: abs(pair[0] - elevation_ft))
    if abs(best[0] - elevation_ft) <= tolerance_ft:
        return best[1]
    return None


# Ordered list of Level-related BuiltInParameter NAMES. Resolution is
# robust: tried via ``getattr`` first (fast, works for most BIPs), then
# via ``Enum.Parse`` as a fallback for the pythonnet quirk where some
# enum values don't surface as Python attributes even though they
# exist in the underlying .NET enum. ``INSTANCE_LEVEL_PARAM`` is the
# primary candidate for the CED MEP family catalog and was the key
# missing piece from the legacy engine.
_INSTANCE_LEVEL_BIP_NAMES = (
    "INSTANCE_LEVEL_PARAM",
    "FAMILY_LEVEL_PARAM",
    "SCHEDULE_LEVEL_PARAM",
    "INSTANCE_REFERENCE_LEVEL_PARAM",
    "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM",
)


def _resolve_bip(name):
    """Resolve a ``BuiltInParameter`` enum value by string name.

    Tries ``getattr`` first (cheap), then falls back to ``Enum.Parse``
    against the CLR type. The fallback exists because pythonnet 3 has
    been observed to skip some enum values when surfacing the type as
    a Python class — a value that ``Enum.Parse`` finds may show up as
    ``None`` via ``getattr``. Returns the BIP enum value or ``None``.
    """
    bip = getattr(BuiltInParameter, name, None)
    if bip is not None:
        return bip
    try:
        return Enum.Parse(clr.GetClrType(BuiltInParameter), name)
    except Exception:
        return None


_LINKER_LEVEL_ID_RE = re.compile(r"LevelId\s*:\s*(-?\d+)")


def _captured_led_level_id(doc, led, return_source=False):
    """Extract the captured ``LevelId`` integer from the LED's YAML
    parameters and verify it resolves to a real Level in ``doc``.

    Looks in two places, matching the legacy MEP Automation panel's
    ``_get_linker_template`` behavior because the YAML capture format
    is inconsistent across LED records:

        1. ``parameters['LevelId']`` — top-level integer key. Newer
           captures use this.
        2. ``parameters['Element_Linker Parameter']`` or
           ``parameters['Element_Linker']`` — a single string of
           ``key: value, ...`` pairs that embeds ``LevelId: NNN``.
           Older captures store the LevelId only inside this string.

    Returns the integer ElementId value if the captured id resolves to
    a Level in ``doc``; otherwise ``None``. When ``return_source`` is
    True, returns ``(level_id, source_str)`` where ``source_str`` is
    ``'top-key'``, ``'linker-payload'``, or one of the failure tags
    (``'no-params'``, ``'no-value'``, ``'parse-fail'``, ``'negative'``,
    ``'get-failed'``, ``'not-a-level'``).
    """
    if doc is None or not isinstance(led, dict):
        return (None, "no-params") if return_source else None
    params = led.get("parameters") or {}
    if not isinstance(params, dict):
        return (None, "no-params") if return_source else None

    source = "top-key"
    raw = params.get("LevelId")
    if raw in (None, "", "None"):
        # Fallback: extract LevelId from the Element_Linker Parameter
        # payload string. ELDs captured under the older serialization
        # only stored the LevelId there.
        source = "linker-payload"
        raw = None
        for key in ("Element_Linker Parameter", "Element_Linker"):
            payload = params.get(key)
            if isinstance(payload, str) and payload:
                m = _LINKER_LEVEL_ID_RE.search(payload)
                if m:
                    raw = m.group(1)
                    break
        if raw in (None, "", "None"):
            return (None, "no-value") if return_source else None

    try:
        level_id = int(raw)
    except (TypeError, ValueError):
        return (None, "parse-fail") if return_source else None
    if level_id < 0:
        return (None, "negative") if return_source else None
    try:
        elem = doc.GetElement(ElementId(level_id))
    except Exception:
        return (None, "get-failed") if return_source else None
    if elem is None or not isinstance(elem, Level):
        return (None, "not-a-level") if return_source else None
    return (level_id, source) if return_source else level_id


def _is_user_level(level):
    """Filter out legend / placeholder Levels from fallback chains.

    Many Revit projects carry a special Level at very negative elevation
    (commonly named ``XX - Legend Level`` at -100 ft) used purely as the
    workplane for legend views. If our engine falls back to "lowest
    elevation Level in the doc" without filtering, every placed fixture
    ends up bound to that legend level — visually invisible on the
    floor plans but wrong everywhere else. This helper screens those
    out: a Level is "user" only if its elevation is plausibly within a
    building's normal range AND its name doesn't look like a legend /
    placeholder marker.
    """
    if level is None:
        return False
    try:
        elev = float(getattr(level, "Elevation", 0.0))
    except Exception:
        return False
    # -10 ft lower bound covers basements; anything below is a placeholder.
    if elev < -10.0:
        return False
    name = (getattr(level, "Name", "") or "").strip().lower()
    if not name:
        return False
    if name.startswith("xx") or name.startswith("zz"):
        return False
    if "legend" in name or "template" in name or "placeholder" in name:
        return False
    return True


def _first_level_id_from_active_view(uidoc, doc):
    """Return the integer ElementId value of a real Level to use as a
    fallback when the LED's captured ``LevelId`` is missing or doesn't
    resolve. Priority:

        1. Active view's ``GenLevel`` (when the active view is a plan).
        2. Lowest-elevation Level that passes ``_is_user_level`` —
           legend / placeholder levels are excluded so a non-plan
           active view doesn't silently park every fixture at the
           legend.

    Returns ``None`` if no usable Level exists.
    """
    if uidoc is not None:
        try:
            active = uidoc.ActiveView
            if active is not None:
                gen = getattr(active, "GenLevel", None)
                if gen is not None:
                    eid = gen.Id
                    val = getattr(eid, "Value", None)
                    if val is None:
                        val = getattr(eid, "IntegerValue", None)
                    if val is not None:
                        return val
        except Exception:
            pass
    if doc is None:
        return None
    try:
        levels = [
            lv for lv in FilteredElementCollector(doc).OfClass(Level)
            if _is_user_level(lv)
        ]
        levels.sort(key=lambda lv: getattr(lv, "Elevation", 0.0))
        if levels:
            eid = levels[0].Id
            val = getattr(eid, "Value", None)
            if val is None:
                val = getattr(eid, "IntegerValue", None)
            return val
    except Exception:
        pass
    return None


def _ensure_instance_level_param(doc, inst, level_id, uidoc=None, diagnostics_by_family=None):
    """Set the first writeable Level-related instance parameter on
    ``inst`` to the live Level element's ``ElementId``. Returns True
    if a write succeeded.

    Mirrors the legacy engine's ``_apply_recorded_level`` exactly:
        1. Look up the Level via ``doc.GetElement(ElementId(int(...)))``.
           If the lookup fails or returns a non-Level, bail out — we
           don't want to point a Level parameter at a non-Level
           element by mistake.
        2. Use the live element's ``.Id`` rather than a freshly
           constructed ``ElementId(int)``. In Revit 2024+ where
           ElementId is 64-bit, the constructor variant can produce a
           subtly different value than the live property; using ``.Id``
           is the canonical way and is what the legacy engine uses.
        3. Try each ``BuiltInParameter`` in priority order. **Stop at
           the first successful write** — fanning out across multiple
           Level BIPs has been observed to confuse Revit on families
           that expose more than one.
        4. Tier 2: ``LookupParameter`` by display name (``"Level"``,
           ``"Reference Level"``, ``"Schedule Level"``) as a safety
           net for families whose Level parameter is a shared /
           family-defined parameter rather than a BIP.
    """
    if doc is None or inst is None or level_id is None:
        return False
    try:
        level_element = doc.GetElement(ElementId(int(level_id)))
    except Exception:
        level_element = None
    if level_element is None:
        return False
    try:
        live_eid = level_element.Id
    except Exception:
        return False

    # Diagnostic record key — one entry per family so 100 placements of
    # the same family produce 1 line, not 100.
    fam_key = None
    if diagnostics_by_family is not None:
        try:
            sym = getattr(inst, "Symbol", None)
            fam = getattr(sym, "Family", None) if sym is not None else None
            fam_key = getattr(fam, "Name", None) if fam is not None else None
        except Exception:
            fam_key = None

    attempts = []

    def _record(label, outcome):
        if diagnostics_by_family is None or fam_key is None:
            return
        attempts.append("{}={}".format(label, outcome))

    any_write = False

    # Tier 1 — BIPs. Try every candidate, set every writeable one. We
    # used to return on the first success (mirroring the legacy
    # ``_apply_recorded_level``), but that meant the parameter the user
    # actually sees in the Properties palette might never be written
    # if a different BIP succeeded first. Setting all writeable Level
    # BIPs to the same level keeps every Level-shaped instance
    # parameter consistent.
    for name in _INSTANCE_LEVEL_BIP_NAMES:
        bip = _resolve_bip(name)
        if bip is None:
            _record(name, "BIP-missing")
            continue
        try:
            param = inst.get_Parameter(bip)
        except Exception as exc:
            _record(name, "get-threw:{}".format(exc))
            continue
        if param is None:
            _record(name, "not-on-family")
            continue
        try:
            ro = param.IsReadOnly
        except Exception:
            ro = True
        if ro:
            _record(name, "read-only")
            continue
        try:
            param.Set(live_eid)
            _record(name, "SET-OK")
            any_write = True
        except Exception as exc:
            _record(name, "set-threw:{}".format(exc))
            continue
    # Tier 2 — LookupParameter by display name. ALWAYS runs (even if a
    # BIP succeeded) to catch family-defined / shared Level parameters
    # that aren't backed by a BIP. Common name on CED families is just
    # "Level" exposed under Constraints.
    for label in ("Level", "Reference Level", "Schedule Level"):
        try:
            param = inst.LookupParameter(label)
        except Exception:
            param = None
        if param is None:
            _record("Lookup({})".format(label), "not-found")
            continue
        try:
            ro = param.IsReadOnly
        except Exception:
            ro = True
        if ro:
            _record("Lookup({})".format(label), "read-only")
            continue
        try:
            param.Set(live_eid)
            _record("Lookup({})".format(label), "SET-OK")
            any_write = True
        except Exception as exc:
            _record("Lookup({})".format(label), "set-threw:{}".format(exc))
            continue
    if diagnostics_by_family is not None and fam_key is not None:
        prefix = "[{}]{} ".format(
            fam_key, "" if any_write else " no-write:"
        )
        diagnostics_by_family.setdefault(fam_key, prefix + ", ".join(attempts))
    return any_write


def _find_plan_view_for_level(doc, level_id_int):
    """Return a non-template ``ViewPlan`` whose ``GenLevel`` ElementId
    equals ``level_id_int`` — or ``None`` when no such view exists.

    Used by ``execute_placement`` to switch the active view to the
    target level's plan before each ``NewFamilyInstance`` call. That's
    the only way to make Revit populate the ``Level`` instance
    parameter on workplane-based families (which is most of the CED
    MEP catalog); the level-bearing ``NewFamilyInstance`` overload
    leaves Level absent for those families.
    """
    if doc is None or level_id_int is None:
        return None
    try:
        collector = FilteredElementCollector(doc).OfClass(ViewPlan)
    except Exception:
        return None
    for v in collector:
        try:
            if v.IsTemplate:
                continue
            gen = v.GenLevel
            if gen is None:
                continue
            gid = gen.Id
            gval = getattr(gid, "Value", None)
            if gval is None:
                gval = getattr(gid, "IntegerValue", None)
            if gval == level_id_int:
                return v
        except Exception:
            continue
    return None


def active_view_is_level_plan(doc, level_id_int):
    """True when ``doc.ActiveView`` is a plan view whose ``GenLevel``
    ElementId equals ``level_id_int`` — i.e. the current view can already
    bind this level's workplane-based instances, so the gateway needn't
    switch the active view (and disturb the user's open view) at all.
    """
    if doc is None or level_id_int is None:
        return False
    try:
        av = doc.ActiveView
    except Exception:
        return False
    if av is None:
        return False
    try:
        gen = getattr(av, "GenLevel", None)
    except Exception:
        gen = None
    if gen is None:
        return False
    try:
        gval = getattr(gen.Id, "Value", None)
        if gval is None:
            gval = getattr(gen.Id, "IntegerValue", None)
        return int(gval) == int(level_id_int)
    except Exception:
        return False


def _element_phase_id_value(elem):
    """Return the integer ElementId of the element's PHASE_CREATED, or
    ``None`` when the parameter is absent / has no value / is the
    ``ElementId(-1)`` "none" sentinel.

    Phase ElementIds are document-local — callers must compare against
    a phase id pulled from the same document the element lives in.
    """
    if elem is None:
        return None
    try:
        param = elem.get_Parameter(BuiltInParameter.PHASE_CREATED)
    except Exception:
        return None
    if param is None:
        return None
    try:
        eid = param.AsElementId()
    except Exception:
        return None
    if eid is None:
        return None
    val = getattr(eid, "Value", None)
    if val is None:
        val = getattr(eid, "IntegerValue", None)
    if val is None or val == -1:
        return None
    return val


def find_targets_in_host_model(doc, phase_id=None):
    """Walk the host doc's FamilyInstances + Groups and produce
    one Target per element. Anchors are matched by family name
    (same suffix-strip rules as linked Revit).

    ``phase_id`` (optional) restricts the walk to elements whose
    ``PHASE_CREATED`` ElementId equals this value. Elements with no
    PHASE_CREATED parameter — or with ``ElementId(-1)`` — are dropped
    when a phase filter is active.

    Each ``Target.level_id`` is the source element's own Level
    ElementId (host-doc namespace) when readable — so the placed child
    inherits its parent's level rather than the active view's level.
    """
    if doc is None:
        return []
    out = []
    identity = Transform.Identity
    for klass in (FamilyInstance, Group):
        collector = (
            FilteredElementCollector(doc)
            .OfClass(klass)
            .WhereElementIsNotElementType()
        )
        for elem in collector:
            if phase_id is not None:
                if _element_phase_id_value(elem) != phase_id:
                    continue
            name = _element_family_name(elem)
            if not name:
                continue
            pt = _element_location_point(elem)
            if pt is None:
                continue
            rot_deg = _element_rotation_deg(elem, identity)
            eid = elem.Id
            eid_val = getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
            out.append(Target(
                source=SOURCE_HOST_MODEL,
                name=name,
                world_pt=(pt.X, pt.Y, pt.Z),
                rotation_deg=rot_deg,
                level_id=_element_level_id_value(elem),
                link_inst=None,
                link_elem_id=eid_val,
            ))
    return out


def find_targets_in_linked_revit(link_inst, phase_id=None):
    """Walk the linked doc's FamilyInstances + Groups and produce
    one Target per element.

    ``phase_id`` (optional) is the linked doc's Phase ElementId value;
    elements whose ``PHASE_CREATED`` doesn't match are dropped. Phase
    ids are document-local — the caller must source ``phase_id`` from
    the same link's GetLinkDocument().

    Each ``Target.level_id`` is the **host** doc's Level whose elevation
    is closest to the linked element's level (after adding the link
    transform's Z offset). Match is by elevation, not by name — so
    host/link naming mismatches don't break the inheritance. When no
    host level is within the elevation tolerance, ``level_id`` is
    ``None`` and ``execute_placement`` falls back to
    ``options.default_level_id``.
    """
    if link_inst is None:
        return []
    link_doc = link_inst.GetLinkDocument()
    if link_doc is None:
        return []
    transform = links.get_link_transform(link_inst) or Transform.Identity
    host_doc = getattr(link_inst, "Document", None)
    host_level_index = _build_host_level_index(host_doc)
    # Link transform's Z translation — added to a linked-doc Level's
    # ``Elevation`` to get the equivalent host-world elevation. Assumes
    # the link isn't rotated around X/Y (always true in practice for
    # building models — they're only rotated around Z and offset).
    try:
        z_offset = float(transform.Origin.Z)
    except Exception:
        z_offset = 0.0
    # Per-linked-doc-level-id cache: avoid re-resolving the same level
    # for every element on that level.
    linked_to_host_level = {}
    out = []
    for klass in (FamilyInstance, Group):
        collector = (
            FilteredElementCollector(link_doc)
            .OfClass(klass)
            .WhereElementIsNotElementType()
        )
        for elem in collector:
            if phase_id is not None:
                if _element_phase_id_value(elem) != phase_id:
                    continue
            name = _element_family_name(elem)
            if not name:
                continue
            pt = _element_location_point(elem)
            if pt is None:
                continue
            world_pt = transform.OfPoint(pt)
            rot_deg = _element_rotation_deg(elem, transform)
            eid = elem.Id
            eid_val = getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
            host_level_id = None
            linked_level_id = _element_level_id_value(elem)
            if linked_level_id is not None:
                if linked_level_id in linked_to_host_level:
                    host_level_id = linked_to_host_level[linked_level_id]
                else:
                    try:
                        linked_level = link_doc.GetElement(
                            ElementId(int(linked_level_id))
                        )
                        if linked_level is not None:
                            target_elev = float(linked_level.Elevation) + z_offset
                            host_level_id = _resolve_level_by_elevation(
                                target_elev, host_level_index
                            )
                    except Exception:
                        host_level_id = None
                    linked_to_host_level[linked_level_id] = host_level_id
            out.append(Target(
                source=SOURCE_LINKED_REVIT,
                name=name,
                world_pt=(world_pt.X, world_pt.Y, world_pt.Z),
                rotation_deg=rot_deg,
                level_id=host_level_id,
                link_inst=link_inst,
                link_elem_id=eid_val,
            ))
    return out


def _element_family_name(elem):
    if isinstance(elem, FamilyInstance):
        sym = getattr(elem, "Symbol", None)
        if sym is not None:
            family = getattr(sym, "Family", None)
            return family.Name if family else ""
    if isinstance(elem, Group):
        gtype = getattr(elem, "GroupType", None)
        return gtype.Name if gtype else ""
    return ""


def _element_location_point(elem):
    loc = getattr(elem, "Location", None)
    if loc is None:
        return None
    pt = getattr(loc, "Point", None)
    if pt is not None:
        return pt
    bbox = elem.get_BoundingBox(None)
    if bbox is not None:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
    return None


def _element_rotation_deg(elem, transform):
    """Note: ``hasattr(loc, 'Rotation')`` would invoke the property
    getter in pythonnet, which throws InvalidOperationException for
    line-based families etc. Use ``isinstance(loc, LocationPoint)``
    instead — that's a safe type check that doesn't touch the property.
    """
    rad = 0.0
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        try:
            rad = loc.Rotation
        except Exception:
            rad = 0.0
    if transform is not None:
        try:
            local = XYZ(math.cos(rad), math.sin(rad), 0.0)
            v = transform.OfVector(local)
            return geometry.normalize_angle(math.degrees(math.atan2(v.Y, v.X)))
        except Exception:
            pass
    return geometry.normalize_angle(math.degrees(rad))


# ---------------------------------------------------------------------
# Target collection: CSV
# ---------------------------------------------------------------------

def collect_csv_target_files():
    """No-op stub: caller picks the file via a UI dialog."""
    return []


_FT_IN_RE = re.compile(
    r"""^\s*
        (?P<sign>-)?\s*
        (?:(?P<feet>\d+)\s*'\s*)?
        (?:-?\s*(?P<whole>\d+))?
        (?:\s*(?P<num>\d+)\s*/\s*(?P<den>\d+))?
        \s*\"?\s*$
    """,
    re.VERBOSE,
)


def parse_feet_value(value):
    """Parse a length value to decimal feet.

    Accepts:
        ``"12.5"``                  -> 12.5
        ``"12'-6\""``               -> 12.5
        ``"12'-6 1/2\""``           -> 12.5417
        ``"-3'-2\""``               -> -3.1667
        ``"-1.25"``                 -> -1.25

    Returns ``None`` on parse failure.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    # Try plain decimal first.
    try:
        return float(text)
    except (TypeError, ValueError):
        pass
    m = _FT_IN_RE.match(text)
    if not m:
        return None
    sign = -1.0 if m.group("sign") else 1.0
    feet = float(m.group("feet") or 0)
    whole_in = float(m.group("whole") or 0)
    num = m.group("num")
    den = m.group("den")
    frac_in = (float(num) / float(den)) if (num and den and float(den) != 0) else 0.0
    inches = whole_in + frac_in
    if feet == 0 and whole_in == 0 and frac_in == 0:
        return None
    return sign * (feet + inches / 12.0)


_HEADER_NAME_KEYS = ("name", "block_name", "block", "blockname", "family")
_HEADER_X_KEYS = ("x", "position x", "position_x", "pos_x", "px")
_HEADER_Y_KEYS = ("y", "position y", "position_y", "pos_y", "py")
_HEADER_Z_KEYS = ("z", "position z", "position_z", "pos_z", "pz", "elevation")
_HEADER_ROT_KEYS = ("rotation", "rot", "rotation_deg", "angle", "rotation (deg)")


def _resolve_header(headers, keys):
    lower = [(h or "").strip().lower() for h in headers]
    for key in keys:
        if key in lower:
            return lower.index(key)
    return None


def find_targets_in_csv(path):
    """Read ``path`` as CSV. Headers are auto-detected; the file must
    have at least a name column and X / Y columns. Z and rotation are
    optional (default 0)."""
    if not path or not os.path.isfile(path):
        return []
    import csv
    rows = []
    with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for raw in reader:
            rows.append(raw)
    if not rows:
        return []
    headers = rows[0]
    name_idx = _resolve_header(headers, _HEADER_NAME_KEYS)
    x_idx = _resolve_header(headers, _HEADER_X_KEYS)
    y_idx = _resolve_header(headers, _HEADER_Y_KEYS)
    z_idx = _resolve_header(headers, _HEADER_Z_KEYS)
    rot_idx = _resolve_header(headers, _HEADER_ROT_KEYS)
    if name_idx is None or x_idx is None or y_idx is None:
        raise CsvParseError(
            "CSV missing required columns. Need at least a name column "
            "(name/block_name) and X/Y columns. Headers seen: {}".format(headers)
        )
    targets = []
    for ridx, raw in enumerate(rows[1:], start=2):
        if not raw:
            continue
        try:
            name = (raw[name_idx] or "").strip()
        except IndexError:
            continue
        if not name:
            continue
        x = parse_feet_value(raw[x_idx] if x_idx < len(raw) else None)
        y = parse_feet_value(raw[y_idx] if y_idx < len(raw) else None)
        z = parse_feet_value(raw[z_idx] if z_idx is not None and z_idx < len(raw) else None) or 0.0
        rot = parse_feet_value(raw[rot_idx] if rot_idx is not None and rot_idx < len(raw) else None) or 0.0
        if x is None or y is None:
            continue
        targets.append(Target(
            source=SOURCE_CSV,
            name=name,
            world_pt=(x, y, z),
            rotation_deg=rot,
            level_id=None,
        ))
    return targets


class CsvParseError(Exception):
    pass


# ---------------------------------------------------------------------
# Target collection: DWG link (best-effort)
# ---------------------------------------------------------------------

def collect_dwg_link_instances(doc):
    """Return loaded DWG-link ImportInstances. Imports (not links) are
    excluded — they can't be reliably differentiated from blocks."""
    out = []
    for inst in FilteredElementCollector(doc).OfClass(ImportInstance):
        if inst is None:
            continue
        try:
            if not inst.IsLinked:
                continue
        except Exception:
            pass
        out.append(inst)
    return out


def find_targets_in_dwg_link(import_inst):
    """Best-effort walk of a DWG link's geometry. Revit's API does not
    reliably expose individual block names, so this may return zero
    usable targets — when it does, names are derived from the geometry's
    own ``Symbol.Name`` if available.
    """
    if import_inst is None:
        return []
    out = []
    try:
        opts = Options()
        opts.IncludeNonVisibleObjects = False
        opts.ComputeReferences = False
        geom = import_inst.get_Geometry(opts)
    except Exception:
        return []
    base_transform = import_inst.GetTotalTransform()
    for obj in geom or []:
        if obj is None:
            continue
        # Each insert / nested xref shows up as a GeometryInstance.
        sym = getattr(obj, "Symbol", None)
        try:
            tx = obj.Transform
        except Exception:
            tx = None
        name = ""
        if sym is not None:
            name = getattr(sym, "Name", "") or ""
        if not name:
            # Some DWGs surface block names via a category mapping; try
            # the GraphicsStyle (a per-layer record). This is unreliable
            # but better than nothing.
            try:
                gs = obj.GraphicsStyleId
                gs_elem = import_inst.Document.GetElement(gs) if gs else None
                if gs_elem is not None:
                    name = getattr(gs_elem, "Name", "") or ""
            except Exception:
                pass
        if not name:
            continue
        if tx is None:
            world_pt = (base_transform.Origin.X, base_transform.Origin.Y, base_transform.Origin.Z)
            rot_deg = 0.0
        else:
            origin = base_transform.OfPoint(tx.Origin)
            world_pt = (origin.X, origin.Y, origin.Z)
            try:
                bx = base_transform.OfVector(tx.BasisX)
                rot_deg = geometry.normalize_angle(math.degrees(math.atan2(bx.Y, bx.X)))
            except Exception:
                rot_deg = 0.0
        out.append(Target(
            source=SOURCE_DWG_LINK,
            name=name,
            world_pt=world_pt,
            rotation_deg=rot_deg,
            level_id=None,
        ))
    return out


# ---------------------------------------------------------------------
# Placement execution
# ---------------------------------------------------------------------

class PlacementResult(object):
    def __init__(self):
        self.placed_fixture_count = 0
        self.placed_annotation_count = 0
        self.element_linker_writes = 0
        self.static_param_writes = 0
        self.parent_directive_writes = 0
        self.warnings = []
        self.errors = []
        self.skipped_already_placed = 0
        self.normalized_match_count = 0
        self.substituted_type_count = 0


# Sentinel for ``PlacementOptions.only_level_id`` meaning "no per-level
# filter — place every LED regardless of its resolved level." Distinct
# from ``None``, which the modeless gateway uses to mean "place only the
# LEDs that have no resolvable level." A plain object() can't collide
# with any real ElementId integer.
_ALL_LEVELS = object()


class PlacementOptions(object):
    """Knobs the dialog feeds into ``execute_placement``."""

    def __init__(self,
                 skip_already_placed=True,
                 default_level_id=None,
                 transaction_action="Place from CAD or Linked Model",
                 allow_type_substitution=False,
                 category_filter=None,
                 uidoc=None,
                 only_level_id=_ALL_LEVELS):
        self.skip_already_placed = skip_already_placed
        self.default_level_id = default_level_id
        self.transaction_action = transaction_action
        self.allow_type_substitution = allow_type_substitution
        # Per-level placement filter. ``_ALL_LEVELS`` (default) places
        # every matched LED. When set to an int (a Level ElementId) or
        # ``None``, ``execute_placement`` places ONLY the LEDs whose
        # resolved level equals this value — the modeless gateway drives
        # one pass per level so it can switch the active view to that
        # level's plan (outside any transaction) before each pass, which
        # is the only way Revit binds the Level on workplane-based
        # families. See ``collect_target_levels`` and ``placement_apply``.
        self.only_level_id = only_level_id
        # Set of LED category names the user chose in the dialog. When
        # non-empty, ``execute_placement`` skips any LED whose
        # ``category`` field isn't in this set — so a profile that mixes
        # Electrical Fixtures + Mechanical Equipment LEDs only places
        # the LEDs in the categories the user actually selected.
        # ``None`` or empty set means "no LED-level filter."
        self.category_filter = (
            set(category_filter) if category_filter else None
        )
        # ``UIDocument`` — needed for the active-view-switching trick
        # that makes the placed instance's ``Level`` parameter exist
        # and be populated correctly for workplane-based families. When
        # ``None``, placement still works but the Level parameter may
        # not appear on the placed instance for workplane-based
        # families (the level-bearing ``NewFamilyInstance`` overload
        # is unreliable for those).
        self.uidoc = uidoc


def _activate_symbol(symbol):
    if symbol is None:
        return
    try:
        if not symbol.IsActive:
            symbol.Activate()
    except Exception:
        pass


def _resolve_family_symbol(doc, family_name, type_name, allow_type_substitution=False):
    """Tiered FamilySymbol lookup.

    Returns ``(symbol_or_None, status, available_types)``::

        status='exact'           exact match on both family and type
        status='normalized'      case-insensitive + ``_NNN`` suffix strip
        status='substituted'     family matched, type didn't; used the
                                 first available type from the family
                                 (only when ``allow_type_substitution``)
        status='family_missing'  no family matched even with normalization
        status='type_missing'    family matched, type didn't, and
                                 substitution was not allowed

    ``available_types`` is the list of types loaded under the matched
    family (sorted) when the family is found; empty otherwise. The
    caller uses it to build actionable warning messages.
    """
    target_family_norm = normalize_name(family_name)
    target_type_norm = normalize_name(type_name)

    # Bucket every loaded FamilySymbol by its normalized family name so
    # we can branch on "family missing" vs "type missing" cheaply.
    by_family_norm = {}  # norm_family -> [(family_actual, type_actual, sym)]
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        family = getattr(sym, "Family", None)
        if family is None:
            continue
        fam_actual = family.Name
        type_actual = sym.Name
        by_family_norm.setdefault(normalize_name(fam_actual), []).append(
            (fam_actual, type_actual, sym)
        )

    matching = by_family_norm.get(target_family_norm) or []
    available_types = sorted({t for _, t, _ in matching})

    # Tier 1 — exact match.
    for fam_actual, type_actual, sym in matching:
        if fam_actual == family_name and type_actual == type_name:
            return sym, "exact", available_types

    # Tier 2 — normalized match on type.
    for fam_actual, type_actual, sym in matching:
        if normalize_name(type_actual) == target_type_norm:
            return sym, "normalized", available_types

    if not matching:
        return None, "family_missing", []

    # Tier 3 — family matched, type didn't. Optionally substitute.
    if allow_type_substitution:
        _, _, sym = matching[0]
        return sym, "substituted", available_types
    return None, "type_missing", available_types


def _resolve_group_type(doc, group_name, group_index=None):
    """Returns ``(group_type_or_None, status)`` where status is
    ``'exact'``, ``'normalized'``, or ``'missing'``.

    ``group_index`` is the optional pre-built doc cache produced by
    ``build_group_type_index``. Pass it for any call inside a hot
    placement loop to avoid an O(N*M) FilteredElementCollector pass
    per LED. When omitted, this function falls back to the live
    collector — so single one-off lookups still work.
    """
    if not group_name:
        return None, "missing"

    if group_index is not None:
        gt = group_index["by_name"].get(group_name)
        if gt is not None:
            return gt, "exact"
        gt = group_index["by_norm"].get(normalize_name(group_name))
        if gt is not None:
            return gt, "normalized"
        return None, "missing"

    target_norm = normalize_name(group_name)
    normalized_hit = None
    for gt in FilteredElementCollector(doc).OfClass(GroupType):
        if gt.Name == group_name:
            return gt, "exact"
        if normalized_hit is None and normalize_name(gt.Name) == target_norm:
            normalized_hit = gt
    if normalized_hit is not None:
        return normalized_hit, "normalized"
    return None, "missing"


def build_group_type_index(doc):
    """Walk every ``GroupType`` once and return a cache for fast lookups.

    Used by the placement loop so each LED's group-vs-family decision
    is O(1) instead of O(group-count) per LED. Both model and detail
    groups arrive on the same collector; ``doc.Create.PlaceGroup``
    fails fast if you hand it a detail group, so we don't pre-filter —
    the family-fallback path catches it.
    """
    by_name = {}
    by_norm = {}
    if doc is None:
        return {"by_name": by_name, "by_norm": by_norm}
    try:
        collector = FilteredElementCollector(doc).OfClass(GroupType)
    except Exception:
        return {"by_name": by_name, "by_norm": by_norm}
    for gt in collector:
        try:
            name = gt.Name
        except Exception:
            continue
        if not name:
            continue
        by_name.setdefault(name, gt)
        by_norm.setdefault(normalize_name(name), gt)
    return {"by_name": by_name, "by_norm": by_norm}


def _split_label(label):
    """``"Family : Type"`` -> ``("Family", "Type")``. Single-word labels
    fall back to ``(label, "")``."""
    if not label:
        return "", ""
    if " : " in label:
        family, type_name = label.split(" : ", 1)
        return family.strip(), type_name.strip()
    return label.strip(), ""


def _place_fixture(doc, led, anchor_world_pt, anchor_rotation_deg, level_id,
                   allow_type_substitution=False, group_index=None,
                   uidoc=None):
    """Place one LED at the resolved world point.

    Returns ``(placed_elem_or_None, status, info)`` where ``status`` is
    one of ``'exact'``, ``'normalized'``, ``'substituted'``,
    ``'family_missing'``, ``'type_missing'``, ``'group_missing'``,
    ``'no_label'``, or ``'create_failed'``. ``info`` is a dict with
    extra context (available_types, requested_family, requested_type,
    requested_group) used by the caller to build warnings.

    Level handling, three-layer:
        1. The caller sets the active view to a plan whose ``GenLevel``
           is ``level_id`` before invoking us. That populates the
           workplane / level reference for workplane-based families.
        2. ``NewFamilyInstance`` is called via the level-bearing
           overload — ``(point, symbol, level, NonStructural)`` — so
           level-based families get bound directly. Falls back to the
           no-level overload only when Revit rejects the level arg
           (rare, but happens for some hosted families).
        3. After placement and ``doc.Regenerate()``, the instance's
           Level-related instance parameters
           (``FAMILY_LEVEL_PARAM`` / ``INSTANCE_REFERENCE_LEVEL_PARAM``
           / ``INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM`` /
           ``SCHEDULE_LEVEL_PARAM``) are explicitly set to the target
           level. Belt-and-suspenders for families where the
           constructor didn't auto-bind.

    Group-vs-family resolution: the YAML ``is_group`` flag is
    unreliable (V5 capture flagged real model groups as
    ``is_group: false`` for some entries), so we ALSO try a model-
    group lookup whenever the label looks group-shaped — i.e. the
    family-name half equals the type-name half (`"X : X"`). That's
    the canonical pattern when a Group is serialized as a label.
    Group lookup uses the ``family_name`` half (groups are named
    ``"X"`` not ``"X : X"`` in Revit's GroupType collection).
    Successful group placement returns immediately; otherwise we
    fall through to the family-symbol path so a misclassified LED
    still lands.
    """
    label = led.get("label") or ""
    is_group = bool(led.get("is_group"))
    offsets_list = led.get("offsets") or []
    offset = offsets_list[0] if offsets_list else {
        "x_inches": 0.0, "y_inches": 0.0, "z_inches": 0.0, "rotation_deg": 0.0,
    }
    target_pt_t = geometry.target_point_from_offsets(
        anchor_world_pt, anchor_rotation_deg, offset
    )
    target_rot_deg = geometry.child_rotation_from_offsets(
        anchor_rotation_deg, offset
    )
    target_pt = XYZ(target_pt_t[0], target_pt_t[1], target_pt_t[2])

    family_name, type_name = _split_label(label)

    # "Looks like a group" heuristic: a label captured from a Revit
    # Group is serialized as "<group name> : <group name>" because
    # groups have no real "type" axis. Catches cases where the YAML
    # ``is_group`` flag is wrong (e.g. P_-prefixed model groups in V5
    # that captured as ``is_group: false``). The P_ prefix is the
    # project-wide naming convention for model groups, so we treat any
    # P_-prefixed label as group-shaped regardless of the family/type
    # split.
    looks_like_group = bool(
        family_name and (
            (type_name and family_name == type_name)
            or family_name.startswith("P_")
            or label.startswith("P_")
        )
    )

    if is_group or looks_like_group:
        # Always look up by the family-name half — groups are named
        # ``"X"`` in the GroupType collection, not ``"X : X"``.
        group_type, gstatus = _resolve_group_type(
            doc, family_name or label, group_index=group_index,
        )
        if group_type is not None:
            try:
                group = doc.Create.PlaceGroup(target_pt, group_type)
            except Exception:
                # Detail group, or some other PlaceGroup-rejecting
                # type — fall through to family path so a real family
                # with the same name (rare) still gets a shot.
                group = None
            if group is not None:
                if abs(target_rot_deg) > geometry.Tolerances.ROTATION_DEG:
                    try:
                        from Autodesk.Revit.DB import ElementTransformUtils, Line
                        axis = Line.CreateBound(
                            target_pt,
                            XYZ(target_pt.X, target_pt.Y, target_pt.Z + 1.0),
                        )
                        ElementTransformUtils.RotateElement(
                            doc, group.Id, axis, math.radians(target_rot_deg),
                        )
                    except Exception:
                        pass
                try:
                    doc.Regenerate()
                except Exception:
                    pass
                return group, gstatus, {"requested_group": family_name or label}
        # Group lookup failed (or PlaceGroup rejected). If the YAML
        # explicitly said is_group AND we have nothing else to try,
        # surface a clean group_missing — otherwise the label might
        # legitimately be a family with family==type (rare but
        # possible) so we fall through.
        if is_group and not looks_like_group and " : " not in label:
            return None, "group_missing", {"requested_group": family_name or label}

    if not family_name:
        return None, "no_label", {}
    symbol, status, available_types = _resolve_family_symbol(
        doc, family_name, type_name,
        allow_type_substitution=allow_type_substitution,
    )
    info = {
        "requested_family": family_name,
        "requested_type": type_name,
        "available_types": available_types,
    }
    if symbol is None:
        return None, status, info

    _activate_symbol(symbol)
    # PER USER CONTRACT: a real Level is ALWAYS passed to
    # NewFamilyInstance. The no-level overload is never used because
    # it leaves the placed instance unbound (LevelId=-1, no
    # FAMILY_LEVEL_PARAM). If the originally-requested level throws,
    # we retry against the active view's GenLevel, then against the
    # first Level we can find in the doc. Each path records what it
    # did in ``info["overload"]`` so the diagnostic surfaced by
    # ``execute_placement`` shows conclusively which fallback ran.
    info["level_throw"] = None
    info["overload"] = "?"

    def _resolve_level_obj(lid):
        """Return a Level element for ``lid`` (int) or None."""
        if not lid:
            return None
        try:
            el = doc.GetElement(ElementId(int(lid)))
        except Exception:
            return None
        return el if isinstance(el, Level) else None

    def _active_view_level():
        if uidoc is None:
            return None
        try:
            av = uidoc.ActiveView
        except Exception:
            return None
        if av is None:
            return None
        try:
            return getattr(av, "GenLevel", None)
        except Exception:
            return None

    def _any_real_level():
        try:
            for lv in FilteredElementCollector(doc).OfClass(Level):
                if isinstance(lv, Level):
                    return lv
        except Exception:
            return None
        return None

    # Build the ordered list of Level candidates to try. The requested
    # level wins if it resolves; otherwise we open with the active
    # view's level so the very first call still has a valid arg.
    candidates = []
    seen_ids = set()

    def _push(lv, source_tag):
        if lv is None:
            return
        try:
            lvid = lv.Id.Value if hasattr(lv.Id, "Value") else lv.Id.IntegerValue
        except Exception:
            return
        if lvid in seen_ids:
            return
        seen_ids.add(lvid)
        candidates.append((lv, source_tag))

    _push(_resolve_level_obj(level_id), "requested")
    _push(_active_view_level(), "active-view")
    _push(_any_real_level(), "any-doc-level")

    if not candidates:
        # No Level exists in the doc at all — placement genuinely can't
        # proceed with the "always level-bearing" contract.
        info["level_throw"] = "no-level-available"
        return None, "create_failed", info

    # CRITICAL — pythonnet overload resolution. ``NewFamilyInstance`` has
    # two 4-arg overloads: ``(XYZ, FamilySymbol, Level, StructuralType)``
    # and ``(XYZ, FamilySymbol, Element host, StructuralType)``. Because a
    # ``Level`` IS an ``Element``, pythonnet 3 (the CPython engine pyRevit
    # runs under) binds the bare call to the Element-HOST overload — it
    # hosts the instance ON the level instead of associating it WITH the
    # level, yielding ``Host = "Level : X"`` but ``LevelId = -1`` and no
    # editable Level parameter. Workplane-based families dodge this because
    # they bind the Level from the active view, but ``OneLevelBased``
    # families are left unbound. Force the Level overload explicitly. (The
    # no-level fallback elsewhere is unambiguous and needs no selector.)
    _level_ctor = None
    try:
        _level_ctor = doc.Create.NewFamilyInstance.Overloads[
            XYZ, FamilySymbol, Level, StructuralType
        ]
    except Exception:
        _level_ctor = None

    inst = None
    used_source = None
    for lv, source_tag in candidates:
        try:
            if _level_ctor is not None:
                inst = _level_ctor(
                    target_pt, symbol, lv, StructuralType.NonStructural
                )
            else:
                inst = doc.Create.NewFamilyInstance(
                    target_pt, symbol, lv, StructuralType.NonStructural
                )
            used_source = source_tag
            break
        except Exception as exc:
            # Remember the FIRST throw (the one against the requested
            # level) for diagnostic purposes — subsequent throws are
            # just fallback failures.
            if info["level_throw"] is None:
                info["level_throw"] = "{}: {}".format(type(exc).__name__, exc)
            inst = None
            continue

    if inst is None:
        return None, "create_failed", info
    info["overload"] = "level-bearing:" + used_source
    if abs(target_rot_deg) > geometry.Tolerances.ROTATION_DEG:
        try:
            from Autodesk.Revit.DB import ElementTransformUtils, Line
            axis = Line.CreateBound(target_pt, XYZ(target_pt.X, target_pt.Y, target_pt.Z + 1.0))
            ElementTransformUtils.RotateElement(
                doc, inst.Id, axis, math.radians(target_rot_deg)
            )
        except Exception:
            pass
    try:
        doc.Regenerate()
    except Exception:
        pass
    # Note: the editable "Level" instance-parameter write is performed
    # by the caller (``execute_placement``) AFTER the captured-static
    # parameters have been applied — matches the legacy
    # ``placement_engine`` order so a captured parameter can't shadow
    # the level write. ``level_id`` is kept on the signature so
    # callers don't have to change shape if the order ever moves back.
    return inst, status, info


def _write_linker(elem, led, profile, target):
    """Stamp the placed element with an Element_Linker payload so audit
    and re-placement tools can find it.

    The CKT_Circuit Number_CEDT and CKT_Panel_CEDT values are pulled
    from the LED's captured ``parameters`` dict and copied into the
    payload — matches the legacy engine, which embeds those two
    circuit-identity strings in Element_Linker so SuperCircuit and the
    audit tools can read them without a YAML round-trip.
    """
    if elem is None:
        return False
    set_id = None
    for s in profile.get("linked_sets") or []:
        if isinstance(s, dict):
            set_id = s.get("id")
            break
    parent_elem_id = (
        target.link_elem_id
        if target.source in (SOURCE_LINKED_REVIT, SOURCE_HOST_MODEL)
        else None
    )
    led_params = led.get("parameters") if isinstance(led, dict) else None
    if not isinstance(led_params, dict):
        led_params = {}
    payload = _el.ElementLinker(
        led_id=led.get("id"),
        set_id=set_id,
        location_ft=list(_element_world_point(elem)) if elem is not None else None,
        rotation_deg=_element_world_rotation_deg(elem),
        parent_rotation_deg=target.rotation_deg,
        parent_element_id=parent_elem_id,
        level_id=_element_level_id_value(elem),
        element_id=_element_id_value(elem),
        facing=_element_facing(elem),
        host_name=target.name,
        parent_location_ft=list(target.world_pt),
        ckt_circuit_number=_param_str(led_params, "CKT_Circuit Number_CEDT"),
        ckt_panel=_param_str(led_params, "CKT_Panel_CEDT"),
    )
    try:
        _el_io.write_to_element(elem, payload)
        return True
    except _el_io.ElementLinkerIOError:
        return False


def _param_str(params, name):
    if not isinstance(params, dict):
        return None
    value = params.get(name)
    if value is None:
        return None
    if isinstance(value, dict):
        # parent / sibling directive — not a static value to embed
        return None
    text = str(value).strip()
    return text or None


# Stamping-only keys: values mirroring Element_Linker bookkeeping or the
# placed element's identity. Skipped during _apply_static_parameters so
# we don't try to overwrite Revit's own ElementId / Level / Position
# from the YAML capture.
#
# ``Mark`` is here because it's a per-instance unique identifier — copying
# it from the captured source onto every placement guarantees Revit's
# "Elements have duplicate Mark values" warning (the source instance, any
# Legend Component referencing the same family, and every new placement
# would all collide on the same Mark).
_STAMP_ONLY_PARAM_KEYS = frozenset({
    "Element_Linker",
    "Element_Linker Parameter",
    "Linked Element Definition ID",
    "Set Definition ID",
    "Parent ElementId",
    "Parent Element ID",
    "Parent ID",
    "Parent Rotation (deg)",
    "Parent_location",
    "Host Name",
    "Location XYZ (ft)",
    "Rotation (deg)",
    "FacingOrientation",
    "Level",
    "LevelId",
    "Level Id",
    "ElementId",
    "Element ID",
    "Element Id",
    "Mark",
})


def _find_parameter(elem, name):
    """Find a parameter by display name with a ``Parameters``-walk fallback.

    ``LookupParameter`` is the cheap path — fast and covers shared /
    family parameters reliably — but it **misses many built-in
    parameters** (e.g. ``INSTANCE_FREE_HOST_OFFSET_PARAM`` shown to
    the user as "Elevation from Level"). Without the fallback, captured
    LED values for built-in parameters silently never apply at
    placement time.

    The fallback walks ``elem.Parameters`` and matches on
    ``Definition.Name`` — slower but bulletproof. Mirrors the equivalent
    helper in ``annotation_placement._find_parameter``.
    """
    try:
        p = elem.LookupParameter(name)
    except Exception:
        p = None
    if p is not None:
        return p
    try:
        for param in elem.Parameters:
            if param is None:
                continue
            try:
                d = param.Definition
            except Exception:
                d = None
            if d is None:
                continue
            try:
                if d.Name == name:
                    return param
            except Exception:
                continue
    except Exception:
        pass
    return None


def _apply_static_parameters(elem, params_dict, warnings=None):
    """Write LED-captured static parameters onto the placed instance.

    Mirrors the legacy ``PlaceElementsEngine._apply_parameters``: walks
    the YAML LED's ``parameters`` dict and sets each entry on the new
    element via ``_find_parameter`` (LookupParameter + Parameters-walk
    fallback) + a storage-type aware ``Set``. Skips:

      * Element_Linker bookkeeping keys (those are owned by ``_write_linker``).
      * Parameters whose value is a directive dict (``BYPARENT(...)`` /
        ``BYSIBLING(...)``) — those resolve at audit time, not now.
      * Read-only parameters and missing parameters.

    ``warnings`` (optional) collects per-parameter diagnostic strings
    when something doesn't apply, so callers can surface to the user
    why a captured "Elevation from Level" / "Mark" / etc. didn't land
    on the placed instance. Read-only and built-in skips are recorded
    only when ``warnings`` is non-None — silent otherwise (matches the
    historical contract for callers that don't want noise).

    Returns ``(written, skipped)`` counts for callers that want to log.
    """
    if elem is None or not params_dict:
        return 0, 0
    written = 0
    skipped = 0

    def _elem_label():
        try:
            eid = getattr(elem, "Id", None)
            if eid is not None:
                return "{} (id {})".format(
                    type(elem).__name__,
                    getattr(eid, "Value", None)
                    or getattr(eid, "IntegerValue", None),
                )
        except Exception:
            pass
        return type(elem).__name__

    def _fmt_val(v):
        # Quote with double-quotes around the raw string so feet-inches
        # values render as ``"3' - 8""`` instead of Python's
        # repr-escaped ``'3\' - 8"'`` form. Users were reading the
        # backslash as data corruption — purely a display artifact.
        if v is None:
            return '""'
        if isinstance(v, (int, float)):
            return str(v)
        return '"{}"'.format(v)

    for name, value in params_dict.items():
        if not name or name in _STAMP_ONLY_PARAM_KEYS:
            skipped += 1
            continue
        if isinstance(value, dict):
            # parent / sibling directive — defer to audit-time wiring
            skipped += 1
            continue
        param = _find_parameter(elem, name)
        if param is None:
            skipped += 1
            if warnings is not None:
                warnings.append(
                    "Parameter '{}' not found on placed {} — value {} skipped. "
                    "(Likely a TYPE parameter on the family symbol — instance "
                    "writes don't see those.)".format(
                        name, _elem_label(), _fmt_val(value),
                    )
                )
            continue
        try:
            if param.IsReadOnly:
                skipped += 1
                if warnings is not None:
                    warnings.append(
                        "Parameter '{}' is read-only on {} — value {} skipped.".format(
                            name, _elem_label(), _fmt_val(value),
                        )
                    )
                continue
        except Exception:
            skipped += 1
            continue
        if value is None or value == "":
            # Setting empty string clears the parameter; do nothing instead.
            skipped += 1
            continue
        if _set_param_value(param, value):
            written += 1
        else:
            skipped += 1
            if warnings is not None:
                warnings.append(
                    "Failed to write parameter '{}' = {} on {}. (Family-side "
                    "association: the instance parameter is driven by a type/"
                    "global parameter, so per-instance writes are silently "
                    "rejected even when IsReadOnly returns False. Set the "
                    "matching TYPE parameter on the FamilySymbol instead — or "
                    "break the association in the family editor.)".format(
                        name, _fmt_val(value), _elem_label(),
                    )
                )
    return written, skipped


def _resolve_target_parent_element(doc, target):
    """Return the live Revit parent element a ``Target`` was matched
    against, or ``None`` when the target has no resolvable parent.

    * ``SOURCE_HOST_MODEL`` — the parent lives in the active doc;
      ``target.link_elem_id`` is its host-doc ElementId value.
    * ``SOURCE_LINKED_REVIT`` — the parent lives inside a linked
      doc; resolve it through the RevitLinkInstance's link document.
    * ``SOURCE_CSV`` / ``SOURCE_DWG_LINK`` / ``SOURCE_PICKED_POINT``
      — no Revit parent element exists (a spreadsheet row, a CAD
      block, or a bare clicked point), so parent directives have
      nothing to read from. Returns ``None``.
    """
    if doc is None or target is None:
        return None
    try:
        if target.source == SOURCE_HOST_MODEL and target.link_elem_id is not None:
            return doc.GetElement(ElementId(int(target.link_elem_id)))
        if (
            target.source == SOURCE_LINKED_REVIT
            and target.link_inst is not None
            and target.link_elem_id is not None
        ):
            link_doc = target.link_inst.GetLinkDocument()
            if link_doc is None:
                return None
            return link_doc.GetElement(ElementId(int(target.link_elem_id)))
    except Exception:
        return None
    return None


def _read_parent_param_for_directive(parent_elem, param_name):
    """Read ``param_name`` off the parent element for a BYPARENT
    directive, preferring the **unit-bearing display string**.

    The user-facing contract is: a child that inherits the parent's
    "Amps" should land ``"50 A"`` (value + units), not the bare
    internal double. ``AsValueString()`` renders exactly what the
    Properties palette shows — "50 A", "120 V", "1800 VA", a
    feet-inches string, etc. — and ``_set_param_value`` on the child
    side feeds that straight back through ``SetValueString``, so the
    units round-trip correctly regardless of the child parameter's
    internal storage unit.

    Fallback order when ``AsValueString`` is empty/unsupported:
    String → raw ``AsString``; Integer → int; Double → float;
    ElementId → id value. Returns ``None`` when the parameter is
    absent or has no value.
    """
    if parent_elem is None or not param_name:
        return None
    param = _find_parameter(parent_elem, param_name)
    if param is None:
        return None
    try:
        if not param.HasValue:
            return None
    except Exception:
        pass
    # Unit-bearing display string first — this is the whole point of
    # the feature ("50" on the parent → "50 A" on the child).
    try:
        vs = param.AsValueString()
        if vs is not None and str(vs).strip() != "":
            return vs
    except Exception:
        pass
    try:
        storage = param.StorageType.ToString()
    except Exception:
        storage = ""
    try:
        if storage == "String":
            return param.AsString()
        if storage == "Integer":
            return param.AsInteger()
        if storage == "Double":
            return param.AsDouble()
        if storage == "ElementId":
            eid = param.AsElementId()
            return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
        return param.AsString()
    except Exception:
        return None


def _apply_parent_directives(child, parent_elem, params_dict, warnings=None):
    """Resolve every BYPARENT directive in ``params_dict`` against the
    live ``parent_elem`` and write the result onto ``child``.

    Counterpart to ``_apply_static_parameters`` (which deliberately
    skips directive dicts). Sibling directives are left untouched
    here — they reference another LED in the same set, which is a
    separate placement and resolves at audit time, not from the
    parent element.

    ``warnings`` (optional) collects a line per directive that can't
    be resolved or written so the caller can surface why an inherited
    value didn't land.

    Returns ``(written, skipped)``.
    """
    if child is None or not params_dict:
        return 0, 0
    written = 0
    skipped = 0
    for name, value in params_dict.items():
        if not name or name in _STAMP_ONLY_PARAM_KEYS:
            continue
        if not _dir.is_parent_directive(value):
            # static values + sibling directives are handled elsewhere
            continue
        src_name = _dir.parent_param_name(value)
        if not src_name:
            skipped += 1
            continue
        if parent_elem is None:
            skipped += 1
            if warnings is not None:
                warnings.append(
                    "Directive '{}' <- parent '{}' skipped: this target "
                    "has no resolvable Revit parent element (CSV / CAD / "
                    "picked-point source).".format(name, src_name)
                )
            continue
        parent_value = _read_parent_param_for_directive(parent_elem, src_name)
        if parent_value is None or parent_value == "":
            skipped += 1
            if warnings is not None:
                warnings.append(
                    "Directive '{}' <- parent '{}' skipped: parent has no "
                    "value for '{}'.".format(name, src_name, src_name)
                )
            continue
        param = _find_parameter(child, name)
        if param is None:
            skipped += 1
            if warnings is not None:
                warnings.append(
                    "Directive target '{}' not found on placed child — "
                    "inherited value {!r} skipped.".format(name, parent_value)
                )
            continue
        try:
            if param.IsReadOnly:
                skipped += 1
                if warnings is not None:
                    warnings.append(
                        "Directive target '{}' is read-only on child — "
                        "inherited value {!r} skipped.".format(
                            name, parent_value,
                        )
                    )
                continue
        except Exception:
            skipped += 1
            continue
        if _set_param_value(param, parent_value):
            written += 1
        else:
            skipped += 1
            if warnings is not None:
                warnings.append(
                    "Failed to write inherited value {!r} into child "
                    "parameter '{}' (from parent '{}').".format(
                        parent_value, name, src_name,
                    )
                )
    return written, skipped


def _parse_feet_inches(text):
    """Parse a feet-inches display string (``3' - 8"``, ``1'-6"``,
    ``0' - 4 1/2"``, ``8"``, ``3'``, ``1.5``) into a float in feet,
    Revit's internal length unit. Returns ``None`` on parse failure.

    Used as a fallback when ``SetValueString`` declines a length-style
    YAML value — pythonnet 3 has been observed to reject the standard
    Revit-formatted feet-inches string in some Revit 2026 builds even
    though the same string round-trips through the Properties palette.
    """
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    # Plain decimal feet (or unitless number)
    try:
        return float(s)
    except ValueError:
        pass

    feet_part = ""
    inch_part = s
    if "'" in s:
        feet_part, _, inch_part = s.partition("'")

    try:
        feet = float(feet_part.strip()) if feet_part.strip() else 0.0
    except ValueError:
        return None

    inch_part = inch_part.strip().lstrip("-").strip()
    if inch_part.endswith('"'):
        inch_part = inch_part[:-1].strip()

    inches = 0.0
    if inch_part:
        parts = inch_part.split()
        try:
            if len(parts) == 1:
                token = parts[0]
                if "/" in token:
                    num, _, denom = token.partition("/")
                    inches = float(num) / float(denom)
                else:
                    inches = float(token)
            elif len(parts) == 2:
                whole = float(parts[0])
                num, _, denom = parts[1].partition("/")
                inches = whole + (float(num) / float(denom))
            else:
                return None
        except (ValueError, ZeroDivisionError):
            return None

    return feet + inches / 12.0


def _parse_yes_no(text):
    """Parse a Yes/No / True/False / 1/0 display string into 1 or 0.

    Used as an Integer-storage fallback for boolean Revit parameters
    (``Rotate Symbol Label_CED``, ``Show Symbol Labels_CED``, etc.)
    captured as ``"Yes"`` / ``"No"`` strings — ``SetValueString`` on
    those parameters routinely returns ``False`` because the parser
    expects ``1`` / ``0`` instead.
    """
    if text is None:
        return None
    s = str(text).strip().lower()
    if s in ("yes", "y", "true", "1"):
        return 1
    if s in ("no", "n", "false", "0"):
        return 0
    return None


def _set_param_value(param, value):
    """Best-effort write that honours the parameter's StorageType **and**
    its display units.

    For Double / Integer parameters we prefer ``SetValueString`` so a
    YAML value like ``1800`` (VA) lands as 1800 VA on screen instead of
    being interpreted as 1800 in Revit's *internal* unit (which for
    apparent load is watts). ``SetValueString`` parses the input in the
    parameter's display units and converts to internal storage on the
    Revit side.

    Three-stage fallback when SetValueString declines:

      1. Raw ``Set(float|int)`` — works for unitless numerics.
      2. Feet-inches parser (Double only) — covers length values like
         ``"3' - 8\""`` that pythonnet 3 has been observed to reject
         through SetValueString in Revit 2026.
      3. Yes/No parser (Integer only) — covers boolean parameters
         captured as ``"Yes"`` / ``"No"`` strings.

    Returns ``False`` only when every stage fails.
    """
    try:
        storage = param.StorageType.ToString()
    except Exception:
        storage = ""

    # Strings: never go through SetValueString (it'll try to parse).
    if storage == "String":
        try:
            return bool(param.Set(str(value)))
        except Exception:
            return False

    raw = value if isinstance(value, (int, float)) else str(value).strip()
    raw_str = "" if raw is None else str(raw)
    if not raw_str:
        return False

    if storage in ("Double", "Integer"):
        # Try SetValueString first — handles unit conversion (VA, V,
        # A, ft, deg, etc.) so the value displays exactly as it was
        # captured in the YAML.
        try:
            if param.SetValueString(raw_str):
                return True
        except Exception:
            pass
        # Stage 1: raw numeric Set with the right CLR type. Right for
        # unitless ints / doubles or when the param doesn't support
        # SetValueString (rare).
        try:
            if storage == "Integer":
                return bool(param.Set(int(float(raw))))
            return bool(param.Set(float(raw)))
        except (TypeError, ValueError):
            pass
        except Exception:
            pass
        # Stage 2: feet-inches → feet (Double only). Revit's internal
        # length unit is feet; once we parse to a float, raw Set
        # bypasses the SetValueString picky parser entirely.
        if storage == "Double":
            feet = _parse_feet_inches(raw_str)
            if feet is not None:
                try:
                    return bool(param.Set(float(feet)))
                except Exception:
                    pass
        # Stage 3: Yes/No → 1/0 (Integer only).
        if storage == "Integer":
            yn = _parse_yes_no(raw_str)
            if yn is not None:
                try:
                    return bool(param.Set(int(yn)))
                except Exception:
                    pass
        return False

    # ElementId or unknown — last-resort string set.
    try:
        return bool(param.Set(str(value)))
    except Exception:
        return False


def _element_id_value(elem):
    if elem is None:
        return None
    eid = elem.Id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _element_level_id_value(elem):
    lid = getattr(elem, "LevelId", None)
    if lid is None:
        try:
            param = elem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
            if param is not None:
                lid = param.AsElementId()
        except Exception:
            return None
    if lid is None:
        return None
    return getattr(lid, "Value", None) or getattr(lid, "IntegerValue", None)


def _element_world_point(elem):
    pt = _element_location_point(elem)
    if pt is None:
        return (0.0, 0.0, 0.0)
    return (pt.X, pt.Y, pt.Z)


def _element_world_rotation_deg(elem):
    rad = 0.0
    loc = getattr(elem, "Location", None)
    if isinstance(loc, LocationPoint):
        try:
            rad = loc.Rotation
        except Exception:
            rad = 0.0
    return geometry.normalize_angle(math.degrees(rad))


def _element_facing(elem):
    facing = getattr(elem, "FacingOrientation", None)
    if facing is None:
        return None
    return [facing.X, facing.Y, facing.Z]


def _validate_level_id(doc, level_id_int):
    """Return ``level_id_int`` (as int) only if it resolves to a real
    Level in ``doc``. Filters out the ``-1`` invalid sentinel and any id
    that doesn't actually point at a Level — both of which cause
    ``NewFamilyInstance`` to silently degrade to the no-level overload
    and leave the placed instance unbound (``LevelId=-1``, no
    ``FAMILY_LEVEL_PARAM``). Returns ``None`` when invalid.
    """
    if level_id_int is None:
        return None
    try:
        v = int(level_id_int)
    except (TypeError, ValueError):
        return None
    if v < 0:
        return None
    try:
        el = doc.GetElement(ElementId(v))
    except Exception:
        return None
    if el is None or not isinstance(el, Level):
        return None
    return v


def _resolve_led_level_id(doc, led, target_level_id):
    """Resolve the Level ElementId int one LED will bind to. Priority:

        1. The LED's captured ``LevelId`` (validated against ``doc``).
        2. ``target_level_id`` — the per-target fallback the caller
           already validated (host element level / elevation match /
           dialog default).
        3. ``None`` when neither resolves.

    Returns ``(level_id_or_None, source_str)``. Shared by the placement
    loop and the ``collect_target_levels`` pre-pass so the level the
    gateway switches the view to always matches the level the loop binds.
    """
    captured_id, captured_src = _captured_led_level_id(doc, led, return_source=True)
    if captured_id is not None:
        return captured_id, captured_src
    if target_level_id is not None:
        return target_level_id, "target-fallback"
    return None, "none:" + (captured_src or "?")


def _match_target_level_id(doc, match, options):
    """The validated per-target fallback level for one match — host
    element level / elevation match, else the dialog default."""
    target_level_id = _validate_level_id(doc, match.target.level_id)
    if target_level_id is None:
        target_level_id = _validate_level_id(doc, options.default_level_id)
    return target_level_id


def _iter_match_leds(match, options):
    """Yield each ``led`` dict under ``match`` that passes the active
    ``category_filter``. Centralises the match -> set -> LED walk +
    category gate so the pre-pass and the placement loop agree exactly
    on which LEDs are in scope."""
    for set_dict in match.profile.get("linked_sets") or []:
        if not isinstance(set_dict, dict):
            continue
        for led in set_dict.get("linked_element_definitions") or []:
            if not isinstance(led, dict):
                continue
            if options.category_filter:
                led_cat = (led.get("category") or "").strip()
                if led_cat not in options.category_filter:
                    continue
            yield led


def collect_target_levels(doc, matches, options=None):
    """Read-only pre-pass: the ordered, de-duplicated list of Level
    ElementId values ``execute_placement`` will bind across all
    non-skipped matches, ascending. A trailing ``None`` entry is appended
    when at least one in-scope LED has no resolvable level.

    The modeless gateway (``placement_apply``) calls this BEFORE opening
    any transaction so it can switch the active view to each level's plan
    — the active view at ``NewFamilyInstance`` time is what actually binds
    the Level on workplane-based families, and the view can't be switched
    once a transaction is open. Touches the document read-only only
    (level + LED-parameter lookups), so it's safe outside a transaction.
    """
    if options is None:
        options = PlacementOptions()
    levels = []
    seen = set()
    has_none = False
    for m in matches:
        if m.skip:
            continue
        target_level_id = _match_target_level_id(doc, m, options)
        for led in _iter_match_leds(m, options):
            lid, _src = _resolve_led_level_id(doc, led, target_level_id)
            if lid is None:
                has_none = True
            elif lid not in seen:
                seen.add(lid)
                levels.append(lid)
    levels.sort()
    if has_none:
        levels.append(None)
    return levels


def execute_placement(doc, matches, options=None):
    """Place every non-skipped match. Caller manages the transaction.

    When ``options.only_level_id`` is set (an int Level id or ``None``),
    only the LEDs whose resolved level equals it are placed — the
    modeless gateway uses this to do one view-switched pass per level.
    The default (``_ALL_LEVELS``) places everything, preserving the
    single-transaction behaviour for callers that manage their own view.

    Failures and substitutions are reported as deduped warnings — one
    line per (LED id, status) regardless of how many anchors hit it.
    """
    if options is None:
        options = PlacementOptions()
    result = PlacementResult()

    kept = [m for m in matches if not m.skip]

    # Per-LED already-placed signature set: ``(led_id, target_x, target_y)``.
    # Replaces the old per-set / per-match dedup which silently dropped
    # entire matches once any LED from the set existed at the anchor.
    if options.skip_already_placed:
        placed_led_sigs = _placed_led_anchor_signatures(doc)
    else:
        placed_led_sigs = set()

    # (led_id, status) -> info dict from the first occurrence. Used to
    # collapse "same LED, same problem, 27 anchors" into one warning.
    failure_keys = {}
    substitution_keys = {}
    # Per-LED level diagnostic trace. Keyed by
    # ``(led_id, captured_src, view_outcome, overload, level_match)``
    # so each distinct outcome shows up exactly once with the count of
    # anchors that hit it. Surfaced via result.warnings at end of run.
    level_debug = {}

    # Build the GroupType cache once per run so the group-vs-family
    # decision inside _place_fixture is O(1) per LED.
    group_index = build_group_type_index(doc)

    # Active-view tracking. NOTE: ``uidoc.ActiveView = view`` throws
    # ``InvalidOperationException`` while a Revit Transaction is open
    # (and the dialog wraps execute_placement in one). To work around
    # that, the dialog pre-switches the active view BEFORE opening
    # the transaction, picking a plan view for the most-common target
    # level. Inside the loop we still attempt to switch when an LED
    # needs a different level — it'll throw and we record the throw
    # so the diagnostic shows the cause. The binding works when the
    # currently-active view's GenLevel already matches the LED's
    # level (the common case after the dialog's pre-switch).
    uidoc = getattr(options, "uidoc", None)
    plan_view_by_level = {}
    original_active_view = None
    last_activated_level = None
    if uidoc is not None:
        try:
            original_active_view = uidoc.ActiveView
            # Seed last_activated_level from the active view's GenLevel
            # so the diagnostic correctly reports "cached" when the LED
            # already matches the pre-switched view (instead of trying
            # to re-switch and recording a spurious "switch-threw").
            gl = getattr(original_active_view, "GenLevel", None)
            if gl is not None:
                gid = gl.Id
                last_activated_level = (
                    getattr(gid, "Value", None)
                    or getattr(gid, "IntegerValue", None)
                )
        except Exception:
            original_active_view = None

    try:
        for m in kept:
            anchor = m.target.world_pt
            anchor_rot = m.target.rotation_deg
            # Per-target level used as a fallback for any LED whose own
            # captured ``LevelId`` is missing. Validated so a broken
            # parent (e.g. one whose own ``LevelId=-1``) doesn't
            # propagate the broken value down to ``NewFamilyInstance``.
            target_level_id = _match_target_level_id(doc, m, options)
            target_xy = (
                round(float(anchor[0]), 1),
                round(float(anchor[1]), 1),
            )
            # Resolve the live Revit parent ONCE per target — every LED
            # in this profile inherits BYPARENT directive values from
            # the same parent element. ``None`` for CSV / CAD / picked-
            # point sources (no Revit parent to read), in which case
            # parent directives are reported as skipped per-LED.
            parent_elem = _resolve_target_parent_element(doc, m.target)
            for set_dict in m.profile.get("linked_sets") or []:
                if not isinstance(set_dict, dict):
                    continue
                for led in set_dict.get("linked_element_definitions") or []:
                    if not isinstance(led, dict):
                        continue
                    if options.category_filter:
                        led_cat = (led.get("category") or "").strip()
                        if led_cat not in options.category_filter:
                            continue
                    led_id_for_sig = led.get("id")
                    if (
                        placed_led_sigs
                        and led_id_for_sig
                        and (led_id_for_sig, target_xy[0], target_xy[1])
                            in placed_led_sigs
                    ):
                        result.skipped_already_placed += 1
                        continue
                    # Per-LED level resolution, computed BEFORE
                    # ``_place_fixture`` so the level-bearing
                    # ``NewFamilyInstance`` overload writes the right
                    # FAMILY_LEVEL_PARAM value into the instance at
                    # creation time. Priority:
                    #   1. Captured ``LevelId`` from the LED's YAML
                    #      parameters (validated to resolve to a real
                    #      Level in the target doc). This is the legacy
                    #      ``_apply_recorded_level`` source.
                    #   2. ``target_level_id`` — the per-target level
                    #      resolved by ``find_targets_*`` (host element
                    #      level / elevation-matched / dialog default),
                    #      already validated to resolve to a real Level.
                    led_level_id, led_level_src = _resolve_led_level_id(
                        doc, led, target_level_id,
                    )
                    # Per-level gate. When the gateway drives one pass per
                    # level (``only_level_id`` set), skip every LED that
                    # doesn't belong to the level whose plan view is
                    # currently active — those are placed by their own
                    # pass. ``_ALL_LEVELS`` (default) disables the gate.
                    if (options.only_level_id is not _ALL_LEVELS
                            and led_level_id != options.only_level_id):
                        continue
                    # Active-view switch — keyed on the actual level
                    # value about to be handed to NewFamilyInstance, NOT
                    # the per-target level. (The Revit API quirk that
                    # forces this: NewFamilyInstance(point, sym, level,
                    # NonStructural) silently degrades to the no-level
                    # overload — leaving LevelId=-1 and FAMILY_LEVEL_PARAM
                    # absent — when the active view isn't a plan whose
                    # GenLevel matches the target Level. Keying on
                    # ``target_level_id`` instead of ``led_level_id``
                    # missed every LED whose captured LevelId differed
                    # from its parent's broken LevelId, which is what
                    # produced 415 unbound placements in the last run.)
                    view_outcome = "skipped:no-uidoc"
                    if uidoc is None:
                        view_outcome = "skipped:no-uidoc"
                    elif led_level_id is None:
                        view_outcome = "skipped:no-level-id"
                    elif led_level_id == last_activated_level:
                        view_outcome = "cached"
                    else:
                        plan_view = plan_view_by_level.get(led_level_id)
                        if plan_view is None:
                            plan_view = _find_plan_view_for_level(
                                doc, led_level_id
                            )
                            plan_view_by_level[led_level_id] = plan_view
                        if plan_view is None:
                            view_outcome = "no-plan-view-for-level"
                        else:
                            try:
                                uidoc.ActiveView = plan_view
                                last_activated_level = led_level_id
                                view_outcome = "switched"
                            except Exception as exc:
                                view_outcome = "switch-threw:{}".format(
                                    type(exc).__name__
                                )
                    placed, status, info = _place_fixture(
                        doc, led, anchor, anchor_rot, led_level_id,
                        allow_type_substitution=options.allow_type_substitution,
                        group_index=group_index,
                        uidoc=uidoc,
                    )
                    led_id = led.get("id") or "?"
                    led_label = led.get("label") or "?"
                    # Verify the placed instance actually got bound to
                    # the level we intended. ``actual_level_id`` is what
                    # Revit ended up with on FamilyInstance.LevelId.
                    actual_level_id = None
                    if placed is not None:
                        try:
                            alid = placed.LevelId
                            if alid is not None:
                                actual_level_id = (
                                    getattr(alid, "Value", None)
                                    or getattr(alid, "IntegerValue", None)
                                )
                        except Exception:
                            actual_level_id = None
                    # Diagnostic key — collapses identical outcomes
                    # across many anchors. We sample the FIRST anchor's
                    # full context per key.
                    overload_used = (info or {}).get("overload", "?")
                    dbg_key = (
                        led_id,
                        led_level_src,
                        view_outcome,
                        overload_used,
                        "match" if (actual_level_id == led_level_id
                                    and led_level_id is not None)
                        else "mismatch:{}->{}".format(led_level_id, actual_level_id),
                    )
                    entry = level_debug.get(dbg_key)
                    if entry is None:
                        entry = {
                            "led_id": led_id,
                            "led_label": led_label,
                            "intended_level": led_level_id,
                            "actual_level": actual_level_id,
                            "src": led_level_src,
                            "view": view_outcome,
                            "overload": overload_used,
                            "level_throw": (info or {}).get("level_throw"),
                            "count": 0,
                        }
                        level_debug[dbg_key] = entry
                    entry["count"] += 1
                    if placed is None:
                        failure_keys.setdefault(
                            (led_id, status),
                            {"label": led_label, "info": info},
                        )
                        continue
                    result.placed_fixture_count += 1
                    if status == "normalized":
                        result.normalized_match_count += 1
                    elif status == "substituted":
                        result.substituted_type_count += 1
                        substitution_keys.setdefault(
                            (led_id, status),
                            {"label": led_label, "info": info},
                        )
                    # Stamp the just-placed LED into the dedup set so
                    # within the same run we don't re-attempt the same
                    # LED at the same anchor (e.g. multiple matches
                    # collapsed to one target).
                    if led_id_for_sig:
                        placed_led_sigs.add(
                            (led_id_for_sig, target_xy[0], target_xy[1])
                        )
                    # Apply the LED's captured static parameters (CKT_*,
                    # Voltage_CED, Number of Poles_CED, Apparent Load
                    # Input_CED, etc.) to the new instance before stamping
                    # the Element_Linker so the linker write picks up the
                    # CKT_* fields fresh from the same source. Order matters:
                    # _apply_static_parameters must run before _write_linker.
                    written, _skipped = _apply_static_parameters(
                        placed, led.get("parameters")
                    )
                    result.static_param_writes += written
                    # Resolve BYPARENT directives against the live parent
                    # and write the inherited (unit-bearing) values onto
                    # the child. Runs AFTER static params so an inherited
                    # value always wins over any stale captured static for
                    # the same parameter, and BEFORE _write_linker so the
                    # linker's CKT_* snapshot reflects inherited circuiting.
                    dir_written, _dir_skipped = _apply_parent_directives(
                        placed, parent_elem, led.get("parameters"),
                        warnings=result.warnings,
                    )
                    result.parent_directive_writes += dir_written
                    # Post-creation Level writeback to SCHEDULE_LEVEL_PARAM
                    # (and any other writeable Level-shaped BIP /
                    # LookupParameter the family exposes). Uses the
                    # SAME ``led_level_id`` we passed to
                    # ``NewFamilyInstance`` so FAMILY_LEVEL_PARAM and
                    # SCHEDULE_LEVEL_PARAM agree on the same Level.
                    if led_level_id is not None:
                        _ensure_instance_level_param(
                            doc, placed, led_level_id,
                            uidoc=uidoc,
                        )
                    if _write_linker(placed, led, m.profile, m.target):
                        result.element_linker_writes += 1
    finally:
        # Restore the user's original active view so the placement run
        # doesn't leave them stranded on a different level's plan.
        if uidoc is not None and original_active_view is not None:
            try:
                uidoc.ActiveView = original_active_view
            except Exception:
                pass

    # Per-LED level diagnostic — one line per unique outcome with the
    # anchor count, so a clean run shows few lines and any failure
    # mode pinpoints itself (which LED, where the LevelId came from,
    # whether the view switched, which overload ran, what the placed
    # instance actually got).
    for dbg_key, entry in level_debug.items():
        line = (
            "LVL_DBG led={led_id} intended={intended_level} got={actual_level} "
            "src={src} view={view} overload={overload} n={count}"
        ).format(**entry)
        if entry.get("level_throw"):
            line += " throw=({})".format(entry["level_throw"])
        result.warnings.append(line)
    for (led_id, status), entry in failure_keys.items():
        result.warnings.append(_format_failure_warning(led_id, entry, status))
    for (led_id, status), entry in substitution_keys.items():
        info = entry["info"] or {}
        avail = info.get("available_types") or []
        result.warnings.append(
            "LED {} ({}): type '{}' missing in family '{}'. Substituted "
            "first available type: '{}'. Other types: {}.".format(
                led_id,
                entry["label"],
                info.get("requested_type") or "",
                info.get("requested_family") or "",
                avail[0] if avail else "?",
                ", ".join(t for t in avail[1:6]) or "(none)",
            )
        )

    return result


def _format_failure_warning(led_id, entry, status):
    label = entry["label"]
    info = entry["info"] or {}
    if status == "family_missing":
        return (
            "LED {} ({}): family '{}' is not loaded in the project — load it, "
            "then re-run.".format(led_id, label, info.get("requested_family") or "")
        )
    if status == "type_missing":
        avail = info.get("available_types") or []
        avail_text = ", ".join(avail[:6]) if avail else "(none)"
        if len(avail) > 6:
            avail_text += ", ..."
        return (
            "LED {} ({}): family '{}' is loaded but type '{}' is not. "
            "Available types: {}. Tip: enable 'Allow type substitution' to "
            "place against the first available type.".format(
                led_id, label,
                info.get("requested_family") or "",
                info.get("requested_type") or "",
                avail_text,
            )
        )
    if status == "group_missing":
        return (
            "LED {} ({}): group '{}' is not loaded in the project.".format(
                led_id, label, info.get("requested_group") or ""
            )
        )
    if status == "no_label":
        return "LED {} ({}): no usable label.".format(led_id, label)
    if status == "create_failed":
        return (
            "LED {} ({}): family/type resolved but Revit refused to create "
            "the instance (likely a hosting / level issue).".format(led_id, label)
        )
    return "LED {} ({}): placement failed ({}).".format(led_id, label, status)
