# -*- coding: utf-8 -*-
"""Tests for fuzzy_match.py (alias-proposal scoring)."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import fuzzy_match as fm


_FAILS = []


def _check(name, condition, detail=""):
    if condition:
        print("  PASS  {}".format(name))
    else:
        print("  FAIL  {}  {}".format(name, detail))
        _FAILS.append(name)


def test_similarity():
    print("\n[fuzzy_match] similarity")
    _check("identical -> 100",
           abs(fm.similarity("LTG_2x4 Troffer_CED", "LTG_2x4 Troffer_CED") - 100.0) < 1e-9)
    _check("case-insensitive",
           abs(fm.similarity("abc def", "ABC DEF") - 100.0) < 1e-9)
    _check("empty -> 0", fm.similarity("", "anything") == 0.0)
    _check("both empty -> 0", fm.similarity("", "") == 0.0)
    _check("None -> 0", fm.similarity(None, "x") == 0.0)

    # Suffix near-miss (the b1eefcf gap): must clear the 80 threshold.
    s = fm.similarity("LTG_2x4 Troffer_CED_2", "LTG_2x4 Troffer_CED")
    _check("suffix _2 scores >= 80 ({:.0f})".format(s), s >= 80.0)

    # Reordered / re-delimited tokens: token-set view catches these.
    s = fm.similarity("Troffer 2x4 LTG", "LTG_2x4_Troffer")
    _check("reordered tokens >= 80 ({:.0f})".format(s), s >= 80.0)

    # Genuinely different names stay well below threshold.
    s = fm.similarity("EF-U_Receptacle_CED", "MECH_RTU_Curb")
    _check("unrelated < 80 ({:.0f})".format(s), s < 80.0)


def test_token_set_ratio():
    print("\n[fuzzy_match] token_set_ratio")
    _check("identical tokens any order",
           abs(fm.token_set_ratio("a b c", "c b a") - 100.0) < 1e-9)
    _check("empty -> 0", fm.token_set_ratio("", "x") == 0.0)
    _check("underscore == space delimiting",
           abs(fm.token_set_ratio("a_b_c", "a b c") - 100.0) < 1e-9)


def test_propose_aliases():
    print("\n[fuzzy_match] propose_aliases")
    profile_keys = [
        (0, ["ltg_2x4 troffer_ced"]),
        (1, ["ef-u_receptacle_ced"]),
        (2, ["mech_rtu_curb"]),
    ]
    unmatched = [
        "LTG_2x4 Troffer_CED_2",     # near profile 0
        "EF-U_Receptacle_CED 2026",  # near profile 1
        "Totally Unrelated Widget",  # no proposal
    ]
    proposals = fm.propose_aliases(unmatched, profile_keys, threshold=80.0)
    by_name = {p[0]: p for p in proposals}
    _check("two proposals", len(proposals) == 2, str(proposals))
    _check("troffer -> profile 0",
           by_name.get("LTG_2x4 Troffer_CED_2", (None, None))[1] == 0)
    _check("receptacle -> profile 1",
           by_name.get("EF-U_Receptacle_CED 2026", (None, None))[1] == 1)
    _check("unrelated omitted", "Totally Unrelated Widget" not in by_name)
    for p in proposals:
        _check("score >= threshold ({})".format(p[0]), p[3] >= 80.0)

    # One proposal per name even when several profiles clear threshold.
    dup_keys = [(0, ["fixture_a"]), (1, ["fixture_a"])]
    proposals = fm.propose_aliases(["Fixture_A"], dup_keys, threshold=80.0)
    _check("one proposal per name", len(proposals) == 1)
    _check("tie breaks to lower profile index", proposals[0][1] == 0)

    # Deterministic across key order within a profile.
    proposals = fm.propose_aliases(
        ["Fixture_A"], [(5, ["zzz", "fixture_a", "aaa"])], threshold=80.0)
    _check("best key wins", proposals[0][2] == "fixture_a")

    _check("empty input", fm.propose_aliases([], profile_keys) == [])
    _check("no profiles", fm.propose_aliases(["X"], []) == [])


def run():
    test_similarity()
    test_token_set_ratio()
    test_propose_aliases()
    return list(_FAILS)


if __name__ == "__main__":
    fails = run()
    print("\n[fuzzy_match] {}".format(
        "PASS" if not fails else "FAIL: {}".format(fails)))
    sys.exit(0 if not fails else 1)
