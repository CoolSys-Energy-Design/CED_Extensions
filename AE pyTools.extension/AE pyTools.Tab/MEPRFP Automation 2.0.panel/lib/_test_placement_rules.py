# -*- coding: utf-8 -*-
"""Tests for placement_rules.py (pure label/kind resolution)."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import placement_rules as pr


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def test_split_label():
    print("\n[placement_rules] split_label")
    _check("family : type", pr.split_label("Fam : Type") == ("Fam", "Type"))
    _check("single word", pr.split_label("GroupName") == ("GroupName", ""))
    _check("empty", pr.split_label("") == ("", ""))
    _check("None", pr.split_label(None) == ("", ""))
    _check("only first separator splits",
           pr.split_label("A : B : C") == ("A", "B : C"))
    _check("whitespace trimmed",
           pr.split_label("  Fam  :  Type  ") == ("Fam", "Type"))


def test_explicit_group():
    print("\n[placement_rules] explicit is_group")
    kind, cands, fam, typ = pr.resolve_placement_kind("P_Office Group", True)
    _check("kind group", kind == "group")
    _check("raw label only candidate", cands == ["P_Office Group"])
    _check("family half", fam == "P_Office Group")

    # Group whose own name contains " : " — raw label must be tried
    # FIRST so the un-split name wins over the mangled family half.
    kind, cands, fam, typ = pr.resolve_placement_kind("Odd : Name", True)
    _check("kind group (colon name)", kind == "group")
    _check("raw label first", cands[0] == "Odd : Name")
    _check("family half second", cands[1] == "Odd")

    # Legacy "X : X" serialization with the flag set.
    kind, cands, _, _ = pr.resolve_placement_kind("G1 : G1", True)
    _check("X : X raw first", cands == ["G1 : G1", "G1"])


def test_legacy_heuristic():
    print("\n[placement_rules] legacy X : X heuristic")
    kind, cands, fam, typ = pr.resolve_placement_kind("G1 : G1", False)
    _check("kind maybe_group", kind == "maybe_group")
    _check("family-half candidate", cands == ["G1"])
    _check("family/type halves", (fam, typ) == ("G1", "G1"))


def test_plain_family():
    print("\n[placement_rules] plain family labels")
    kind, cands, fam, typ = pr.resolve_placement_kind(
        "EF-U_Receptacle_CED : Quad Wall", False)
    _check("kind family", kind == "family")
    _check("no group candidates", cands == [])
    _check("split", (fam, typ) == ("EF-U_Receptacle_CED", "Quad Wall"))

    # The old hard-coded P_ prefix heuristic is GONE: an unflagged
    # P_-prefixed family label resolves as a family.
    kind, cands, _, _ = pr.resolve_placement_kind("P_Fam : Type", False)
    _check("P_ prefix no longer implies group", kind == "family")
    _check("P_ prefix no candidates", cands == [])

    # Single-word unflagged label is a (type-less) family, not a group.
    kind, _, fam, typ = pr.resolve_placement_kind("BareName", False)
    _check("single word unflagged -> family", kind == "family")
    _check("single word split", (fam, typ) == ("BareName", ""))


def test_edge_cases():
    print("\n[placement_rules] edge cases")
    kind, cands, _, _ = pr.resolve_placement_kind("", True)
    _check("empty label flagged group", kind == "group")
    _check("empty label no candidates", cands == [])
    kind, _, _, _ = pr.resolve_placement_kind("", False)
    _check("empty label unflagged -> family", kind == "family")
    kind, _, _, _ = pr.resolve_placement_kind(None, False)
    _check("None label -> family", kind == "family")


def run():
    test_split_label()
    test_explicit_group()
    test_legacy_heuristic()
    test_plain_family()
    test_edge_cases()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[placement_rules] {}".format(
        "PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
