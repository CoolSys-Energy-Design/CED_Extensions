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

    # --- group-by parameter discovery ---
    rows_data = [
        {"group_values": {"Identity Mark": "RA-1", "Level": "L1", "Family:Type": "A", "Only_A": "x"}},
        {"group_values": {"Identity Mark": "RA-2", "Level": "L1", "Family:Type": "B", "Only_B": "y"}},
    ]
    opts = cg_core.common_group_params(rows_data)
    check("common params intersect", "Identity Mark" in opts and "Level" in opts and "Family:Type" in opts)
    check("common params exclude non-shared", "Only_A" not in opts and "Only_B" not in opts)
    check("preferred ordering", opts.index("Identity Mark") < opts.index("Family:Type"))
    check("empty rows -> no params", cg_core.common_group_params([]) == [])
    check("default prefers circuit number",
          cg_core.default_group_param(["Identity Mark", "CKT_Circuit Number_CEDT"]) == "CKT_Circuit Number_CEDT")
    check("default falls back to identity",
          cg_core.default_group_param(["Level", "Identity Mark"]) == "Identity Mark")
    check("default first when no preferred",
          cg_core.default_group_param(["Level", "Family:Type"]) == "Level")
    check("default empty when none", cg_core.default_group_param([]) == "")

    # --- Space is offered via its own control, never the parameter combo ---
    space_rows = [
        {"group_values": {"Identity Mark": "RA-1", "Space": "101 - Office"}},
        {"group_values": {"Identity Mark": "RA-2", "Space": "102 - Lab"}},
    ]
    space_opts = cg_core.common_group_params(space_rows)
    check("space excluded from param combo", cg_core.SPACE_GROUP_KEY not in space_opts)
    check("space rows keep real params", "Identity Mark" in space_opts)

    # --- voltage snapping (input is already volt-converted) ---
    check("snap 119.98 -> 120", cg_core.snap_voltage(119.98) == 120)
    check("snap 207.6 -> 208", cg_core.snap_voltage(207.6) == 208)
    check("snap 480.2 -> 480", cg_core.snap_voltage(480.2) == 480)
    check("snap 277 -> 277", cg_core.snap_voltage(277) == 277)
    check("snap non-standard stays rounded", cg_core.snap_voltage(190) == 190)
    check("snap None -> None", cg_core.snap_voltage(None) is None)
    check("snap zero -> None", cg_core.snap_voltage(0) is None)

    print("\nAll tests passed.")


if __name__ == "__main__":
    run()
