# -*- coding: utf-8 -*-
"""Offline tests for the pure-logic parts of placement.py.

We can't import placement.py directly here because it imports the
Revit API at the top. The matching helpers are duplicated here so the
test stays runnable in plain CPython 3 with no Revit. If placement.py's
matching rules change, mirror the change in this file.
"""

from __future__ import print_function

import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


# Mirrors of placement.py functions
_TRAILING_SUFFIX_RE = re.compile(r"_\d+$")


def strip_trailing_suffix(name):
    if not name:
        return ""
    return _TRAILING_SUFFIX_RE.sub("", str(name))


def normalize_name(name):
    return strip_trailing_suffix((name or "").strip()).lower()


def collect_profile_aliases(profile):
    if not isinstance(profile, dict):
        return set()
    props = profile.get("equipment_properties") or {}
    if not isinstance(props, dict):
        return set()
    raw = props.get("cad_aliases")
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
    if not isinstance(profile, dict):
        return set()
    out = set()
    pf = profile.get("parent_filter") or {}
    if isinstance(pf, dict):
        fam = pf.get("family_name_pattern")
        if fam:
            out.add(normalize_name(fam))
    name = profile.get("name") or ""
    if " : " in name:
        fam, _ = name.split(" : ", 1)
        if fam:
            out.add(normalize_name(fam))
    return {n for n in out if n}


def profile_family_names_raw(profile):
    """Mirror of placement.profile_family_names_raw — case-folded only,
    NO trailing ``_NNN`` strip. This is what the exact matchers use."""
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


def profile_full_labels_raw(profile):
    """Mirror of placement.profile_full_labels_raw — case-folded,
    unsplit ``"Family : Type"`` labels for the size-exact tier."""
    if not isinstance(profile, dict):
        return set()
    out = set()

    def _add(value):
        if not value:
            return
        key = value.strip().lower()
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
    """Mirror of placement.collect_profile_aliases_raw — case-folded
    only, no suffix strip."""
    if not isinstance(profile, dict):
        return set()
    props = profile.get("equipment_properties") or {}
    if not isinstance(props, dict):
        return set()
    raw = props.get("cad_aliases")
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


def match_one_linked_revit(target_name, profiles, type_name=""):
    """Mirror of placement._match_one_linked_revit (exact, no suffix
    fallback; size-exact tier when ``type_name`` is provided)."""
    target_name_lower = (target_name or "").strip().lower()
    if not target_name_lower:
        return []

    family_matches = [
        p for p in profiles
        if target_name_lower in profile_family_names_raw(p)
    ]

    type_name_lower = (type_name or "").strip().lower()
    if type_name_lower:
        full_label = "{} : {}".format(target_name_lower, type_name_lower)
        size_exact = [
            p for p in family_matches
            if full_label in profile_full_labels_raw(p)
        ]
        if size_exact:
            return size_exact

    return family_matches


class _Target(object):
    """Minimal stand-in for placement.Target in dedupe tests."""

    def __init__(self, name, type_name="", world_pt=(0.0, 0.0, 0.0)):
        self.name = name
        self.type_name = type_name
        self.world_pt = world_pt


class _Match(object):
    """Minimal stand-in for placement.Match in dedupe tests."""

    def __init__(self, target, profile):
        self.target = target
        self.profile = profile


def dedupe_matches_per_target(matches):
    """Mirror of placement.dedupe_matches_per_target."""
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

        type_name_lower = (group[0].target.type_name or "").strip().lower()
        if type_name_lower:
            full_label = "{} : {}".format(target_name_lower, type_name_lower)
            size_exact = [
                m for m in group
                if full_label in profile_full_labels_raw(m.profile)
            ]
            if size_exact:
                group = size_exact

        exact = [
            m for m in group
            if ((m.profile.get("parent_filter") or {}).get("family_name_pattern") or "")
                .strip().lower() == target_name_lower
        ]
        candidates = exact if exact else group

        candidates = sorted(
            candidates,
            key=lambda m: (
                -_led_count(m.profile),
                m.profile.get("id") or "",
            ),
        )
        out.append(candidates[0])
    return out


def match_one_cad(target_name, profiles):
    """Mirror of placement._match_one_cad (exact, no suffix fallback)."""
    target_name_lower = (target_name or "").strip().lower()
    if not target_name_lower:
        return []
    return [
        p for p in profiles
        if target_name_lower in collect_profile_aliases_raw(p)
        or target_name_lower in profile_family_names_raw(p)
    ]


def profile_flag(profile, key, default=False):
    """Mirror of placement.profile_flag."""
    if not isinstance(profile, dict):
        return default
    val = profile.get(key)
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        low = val.strip().lower()
        if low == "true":
            return True
        if low == "false":
            return False
    return default


# ---- tests ---------------------------------------------------------

def test_strip_trailing_suffix():
    print("\n[placement] strip_trailing_suffix")
    _check("no suffix", strip_trailing_suffix("AC_BLOCK") == "AC_BLOCK")
    _check("_1 suffix", strip_trailing_suffix("AC_BLOCK_1") == "AC_BLOCK")
    _check("_42 suffix", strip_trailing_suffix("AC_BLOCK_42") == "AC_BLOCK")
    _check("multi-digit", strip_trailing_suffix("AC_BLOCK_2321321") == "AC_BLOCK")
    _check("_2A leaves it",
           strip_trailing_suffix("AC_BLOCK_2A") == "AC_BLOCK_2A")
    _check("trailing underscore only",
           strip_trailing_suffix("AC_BLOCK_") == "AC_BLOCK_")
    _check("leading underscore", strip_trailing_suffix("_LEAD_5") == "_LEAD")
    _check("digits-only stays (no underscore prefix)",
           strip_trailing_suffix("123") == "123")
    _check("empty stays empty", strip_trailing_suffix("") == "")
    _check("None -> empty", strip_trailing_suffix(None) == "")


def test_normalize_name():
    print("\n[placement] normalize_name")
    _check("lowercase", normalize_name("AC_BLOCK") == "ac_block")
    _check("strip + lowercase",
           normalize_name("AC_BLOCK_5") == "ac_block")
    _check("whitespace trimmed",
           normalize_name("  AC_BLOCK_3 ") == "ac_block")


def test_collect_profile_aliases():
    print("\n[placement] collect_profile_aliases")
    p1 = {"equipment_properties": {"cad_aliases": "BLOCK_A, BLOCK_B"}}
    aliases = collect_profile_aliases(p1)
    _check("comma-separated",
           aliases == {"block_a", "block_b"},
           "got {}".format(aliases))

    p2 = {"equipment_properties": {"cad_aliases": ["X1", "X2", " "]}}
    _check("list form", collect_profile_aliases(p2) == {"x1", "x2"})

    p3 = {"equipment_properties": {"cad_aliases": "AC_BLOCK_42"}}
    _check("alias normalized via strip",
           collect_profile_aliases(p3) == {"ac_block"})

    _check("missing aliases -> empty", collect_profile_aliases({}) == set())
    _check("non-dict profile -> empty",
           collect_profile_aliases(None) == set())


def test_profile_family_names():
    print("\n[placement] profile_family_names")
    p = {
        "name": "ME_Air Curtain_CED : Mars Air PH1284-2E",
        "parent_filter": {"family_name_pattern": "ME_Air Curtain_CED"},
    }
    names = profile_family_names(p)
    _check("collects family from filter + name split",
           "me_air curtain_ced" in names, "got {}".format(names))

    p2 = {"name": "Foo_Bar_5 : Default"}
    _check("name-only family stripped",
           profile_family_names(p2) == {"foo_bar"})


def test_match_linked_revit_logic():
    """Exact matching: 'Foo_Bar_3' must NOT match a 'Foo_Bar' profile
    (the legacy suffix-strip fallback is gone); exact and case-different
    names still match."""
    print("\n[placement] linked-revit match (exact, case-insensitive)")

    profiles = [
        {
            "id": "EQ-001", "name": "Foo_Bar : Default",
            "parent_filter": {"family_name_pattern": "Foo_Bar"},
        },
        {
            "id": "EQ-002", "name": "Foo_Bar_3 : Default",
            "parent_filter": {"family_name_pattern": "Foo_Bar_3"},
        },
        {
            "id": "EQ-003", "name": "Other : Default",
            "parent_filter": {"family_name_pattern": "Other"},
        },
    ]

    matched = match_one_linked_revit("Foo_Bar", profiles)
    _check("Foo_Bar -> Foo_Bar only",
           len(matched) == 1 and matched[0]["id"] == "EQ-001",
           "got {}".format([m["id"] for m in matched]))

    matched = match_one_linked_revit("Foo_Bar_3", profiles)
    _check("Foo_Bar_3 -> Foo_Bar_3 only (no suffix cross-match)",
           len(matched) == 1 and matched[0]["id"] == "EQ-002",
           "got {}".format([m["id"] for m in matched]))

    matched = match_one_linked_revit("FOO_bar", profiles)
    _check("case-insensitive still matches",
           len(matched) == 1 and matched[0]["id"] == "EQ-001")

    profiles_no_exact = [p for p in profiles if p["id"] != "EQ-002"]
    matched = match_one_linked_revit("Foo_Bar_3", profiles_no_exact)
    _check("Foo_Bar_5-style target with no exact profile -> no match",
           matched == [],
           "got {}".format([m["id"] for m in matched]))

    _check("unknown name -> no match",
           match_one_linked_revit("Different", profiles) == [])
    _check("empty name -> no match",
           match_one_linked_revit("", profiles) == [])


def _size_variant_profiles():
    """Four profiles captured from the same family, one per size —
    the WFM MAT refrigeration-case shape. Ids are in capture order,
    which followed the alphabetical type sort (10' first)."""
    fam = "1160, 1161, 1162, 1163 (HILLPHOENIX, OWZGG) COFFIN"
    out = []
    for i, size in enumerate(("10'", "12'", "6'", "8'")):
        out.append({
            "id": "EQ-{}".format(148 + i),
            "name": "{} : {}".format(fam, size),
            "parent_filter": {"family_name_pattern": fam},
            "linked_sets": [
                {"linked_element_definitions": [
                    {"label": "{} : {}".format(fam, size)},
                ]},
            ],
        })
    return fam, out


def test_match_linked_revit_size_exact():
    """Target carrying the link element's type must resolve to the
    same-size profile, not every size variant of the family."""
    print("\n[placement] linked-revit match (size-exact tier)")
    fam, profiles = _size_variant_profiles()

    matched = match_one_linked_revit(fam, profiles, type_name="8'")
    _check("8' target -> 8' profile only",
           len(matched) == 1 and matched[0]["id"] == "EQ-151",
           "got {}".format([m["id"] for m in matched]))

    matched = match_one_linked_revit(fam, profiles, type_name="8'  ")
    _check("type name whitespace-insensitive",
           len(matched) == 1 and matched[0]["id"] == "EQ-151",
           "got {}".format([m["id"] for m in matched]))

    matched = match_one_linked_revit(fam.upper(), profiles, type_name="12'")
    _check("size tier case-insensitive on family too",
           len(matched) == 1 and matched[0]["id"] == "EQ-149",
           "got {}".format([m["id"] for m in matched]))

    matched = match_one_linked_revit(fam, profiles, type_name="4'")
    _check("uncaptured size falls back to all family matches",
           len(matched) == 4,
           "got {}".format([m["id"] for m in matched]))

    matched = match_one_linked_revit(fam, profiles)
    _check("no type name -> family behavior unchanged",
           len(matched) == 4,
           "got {}".format([m["id"] for m in matched]))


def test_full_label_alias_size_matching():
    """The placement window stores per-pair aliases as full
    ``"Family : Type"`` labels. Those must resolve size-exactly to the
    owning profile, and their family half must still feed the family
    fallback for sizes seen later."""
    print("\n[placement] full-label alias (size-exact via merged_aliases)")
    fam, profiles = _size_variant_profiles()

    # Uncaptured 4' size aliased onto the 6' profile (EQ-150).
    profiles[2]["merged_aliases"] = ["{} : 4'".format(fam)]
    matched = match_one_linked_revit(fam, profiles, type_name="4'")
    _check("aliased size -> alias owner only",
           len(matched) == 1 and matched[0]["id"] == "EQ-150",
           "got {}".format([m["id"] for m in matched]))

    # Renamed source family, full-label alias on the 6' profile.
    profiles[2]["merged_aliases"] = ["FAM_RENAMED : 4'"]
    matched = match_one_linked_revit("FAM_RENAMED", profiles, type_name="4'")
    _check("renamed family + aliased size -> alias owner only",
           len(matched) == 1 and matched[0]["id"] == "EQ-150",
           "got {}".format([m["id"] for m in matched]))

    matched = match_one_linked_revit("FAM_RENAMED", profiles, type_name="12'")
    _check("renamed family, other size -> family fallback via alias half",
           len(matched) == 1 and matched[0]["id"] == "EQ-150",
           "got {}".format([m["id"] for m in matched]))


def test_dedupe_prefers_size_exact():
    """When wrong-size profiles share a target's bucket, the size-exact
    one must win over the LED-count / alphabetical-id tie-breaks."""
    print("\n[placement] dedupe (size-exact tier)")
    fam, profiles = _size_variant_profiles()

    # Give the WRONG-size 10' profile more LEDs so the legacy
    # tie-breaks would pick it; size-exact must still win.
    profiles[0]["linked_sets"][0]["linked_element_definitions"].append(
        {"label": "EF_Receptacle : Duplex"}
    )

    target = _Target(fam, type_name="8'")
    deduped = dedupe_matches_per_target(
        [_Match(target, p) for p in profiles]
    )
    _check("8' anchor -> 8' profile despite richer 10' profile",
           len(deduped) == 1 and deduped[0].profile["id"] == "EQ-151",
           "got {}".format([m.profile["id"] for m in deduped]))

    # Uncaptured size: falls through to the legacy tie-breaks (most
    # LEDs -> the 10' profile).
    target = _Target(fam, type_name="4'")
    deduped = dedupe_matches_per_target(
        [_Match(target, p) for p in profiles]
    )
    _check("uncaptured size falls back to legacy tie-break",
           len(deduped) == 1 and deduped[0].profile["id"] == "EQ-148",
           "got {}".format([m.profile["id"] for m in deduped]))

    # No type name at all: legacy behavior untouched.
    target = _Target(fam)
    deduped = dedupe_matches_per_target(
        [_Match(target, p) for p in profiles]
    )
    _check("no type name -> legacy tie-break unchanged",
           len(deduped) == 1 and deduped[0].profile["id"] == "EQ-148",
           "got {}".format([m.profile["id"] for m in deduped]))


def test_match_cad_alias_logic():
    """CAD/CSV matching is exact against aliases + implicit names —
    no suffix-stripped fallback."""
    print("\n[placement] CAD alias match (exact, case-insensitive)")

    profiles = [
        {
            "id": "EQ-001", "name": "ME_Air Curtain_CED : Mars Air",
            "equipment_properties": {"cad_aliases": "AC_BLOCK, AIR_CURTAIN_BLOCK"},
        },
        {
            "id": "EQ-002", "name": "Other : Default",
            "equipment_properties": {"cad_aliases": "OTHER_BLOCK"},
        },
        {"id": "EQ-003", "name": "No aliases", "equipment_properties": {}},
    ]

    cases = [
        ("AC_BLOCK", "EQ-001"),          # exact alias
        ("ac_block", "EQ-001"),          # case-insensitive
        ("AC_BLOCK_1", None),            # suffix no longer stripped
        ("AC_BLOCK_42", None),
        ("AIR_CURTAIN_BLOCK_99", None),
        ("OTHER_BLOCK", "EQ-002"),
        ("ME_Air Curtain_CED", "EQ-001"),  # implicit alias from name
        ("No aliases", "EQ-003"),          # profile name itself
        ("UNRELATED", None),
    ]
    for block_name, expect_eq in cases:
        matched = match_one_cad(block_name, profiles)
        if expect_eq is None:
            _check("'{}' does not match".format(block_name), matched == [],
                   "got {}".format([m["id"] for m in matched]))
        else:
            _check(
                "'{}' -> {}".format(block_name, expect_eq),
                len(matched) == 1 and matched[0]["id"] == expect_eq,
                "got {}".format([m["id"] for m in matched]),
            )


def test_profile_flag():
    print("\n[placement] profile_flag")
    _check("bool True", profile_flag({"k": True}, "k") is True)
    _check("bool False", profile_flag({"k": False}, "k", default=True) is False)
    _check("string 'true'", profile_flag({"k": "true"}, "k") is True)
    _check("string 'False'", profile_flag({"k": "False"}, "k", default=True) is False)
    _check("missing -> default", profile_flag({}, "k") is False)
    _check("missing -> default True", profile_flag({}, "k", default=True) is True)
    _check("junk string -> default", profile_flag({"k": "yes"}, "k") is False)
    _check("non-dict -> default", profile_flag(None, "k") is False)


def run():
    test_strip_trailing_suffix()
    test_normalize_name()
    test_collect_profile_aliases()
    test_profile_family_names()
    test_match_linked_revit_logic()
    test_match_linked_revit_size_exact()
    test_full_label_alias_size_matching()
    test_dedupe_prefers_size_exact()
    test_match_cad_alias_logic()
    test_profile_flag()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[placement] {}".format("PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
