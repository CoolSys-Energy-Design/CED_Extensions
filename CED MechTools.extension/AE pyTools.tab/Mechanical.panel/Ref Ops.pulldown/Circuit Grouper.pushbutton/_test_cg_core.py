# -*- coding: utf-8 -*-
"""Pure-logic tests for cg_core. Run in plain CPython:
    python _test_cg_core.py
"""
import cg_core


class Row(object):
    """Minimal stand-in for RowVM (only attributes cg_core reads)."""
    def __init__(self, eid, poles, volts, include=True, already=False, ident=""):
        self.element_id = eid
        self.poles_value = poles
        self.voltage_key = volts
        self.include = include
        self.already_circuited = already
        self.identity_mark = ident


def check(name, cond):
    print(("PASS" if cond else "FAIL") + " - " + name)
    if not cond:
        raise AssertionError(name)


def run():
    # --- rating parsing ---
    for txt in ("20", "20 A", "20A", "20 a", "20amp", "20 amps"):
        amps, valid, std = cg_core.parse_rating(txt)
        check("rating %r -> 20 valid standard" % txt, amps == 20.0 and valid and std)
    amps, valid, std = cg_core.parse_rating("garbage")
    check("garbage invalid", (not valid) and amps is None)
    amps, valid, std = cg_core.parse_rating("")
    check("empty invalid", not valid)
    amps, valid, std = cg_core.parse_rating("37")
    check("37 valid but non-standard", valid and not std and amps == 37.0)
    amps, valid, std = cg_core.parse_rating(225)
    check("numeric 225 standard", valid and std)

    check("format number", cg_core.format_amps_number(20.0) == "20")
    check("format amps", cg_core.format_amps(20.0) == "20 A")

    # --- load name common prefix ---
    check("lcp", cg_core.longest_common_prefix(["RA-12A", "RA-12B"]) == "RA-12")
    check("default load name trims", cg_core.default_load_name(["RA-12A", "RA-12B"]) == "RA-12")
    check("default load name single", cg_core.default_load_name(["RC-07"]) == "RC-07")
    check("default load name fallback", cg_core.default_load_name(["AAA", "BBB"], fallback="R9") == "R9")

    # --- validation ---
    check("pole mismatch", len(cg_core.validate_members([Row(1, 1, 120), Row(2, 3, 120)])) == 1)
    check("voltage mismatch", len(cg_core.validate_members([Row(1, 1, 120), Row(2, 1, 208)])) == 1)
    check("clean ok", cg_core.validate_members([Row(1, 1, 120), Row(2, 1, 120)]) == [])

    # --- effective rows: exclude already-circuited + unchecked ---
    rows = [Row(1, 1, 120), Row(2, 1, 120, already=True), Row(3, 1, 120, include=False)]
    eff = cg_core.effective_rows(rows)
    check("effective excludes circuited+unchecked", [r.element_id for r in eff] == [1])

    # --- build_group_plan ---
    plan = cg_core.build_group_plan("R1", "RA-12", "P1", "20", [Row(1, 1, 120), Row(2, 1, 120)])
    check("plan ready", plan["ready"] and plan["rating_amps"] == 20.0 and plan["rating_standard"])
    check("plan ids", plan["element_ids"] == [1, 2])
    check("plan load", plan["load_name"] == "RA-12")

    plan = cg_core.build_group_plan("R2", "", "P1", "20", [Row(1, 1, 120), Row(2, 3, 120)])
    check("plan blocked on mismatch", not plan["ready"] and plan["load_name"] == "R2")

    plan = cg_core.build_group_plan("R3", "x", "P1", "abc", [Row(1, 1, 120)])
    check("plan blocked on bad rating", not plan["ready"] and "Invalid breaker rating" in plan["problems"])

    plan = cg_core.build_group_plan("R4", "x", "P1", "37", [Row(1, 1, 120)])
    check("plan ready but nonstandard", plan["ready"] and not plan["rating_standard"])

    check("new group key", cg_core.next_new_group_key(["NEW-1", "R1"]) == "NEW-2")

    print("\nAll tests passed.")


if __name__ == "__main__":
    run()
