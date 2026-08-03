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

    # --- name-by parameter: default + name seeding -----------------------
    check("name default prefers load name",
          cg_core.default_name_param(["Identity Mark", "CKT_Load Name_CEDT"]) == "CKT_Load Name_CEDT")
    check("name default falls back to identity",
          cg_core.default_name_param(["Level", "Identity Mark"]) == "Identity Mark")
    check("name default first when no preferred",
          cg_core.default_name_param(["Level", "Family:Type"]) == "Level")
    check("name default empty when none", cg_core.default_name_param([]) == "")

    check("name shared value", cg_core.name_from_values(["BAKERY CASE", "BAKERY CASE"]) == "BAKERY CASE")
    check("name ignores blanks", cg_core.name_from_values(["", "BAKERY CASE", None]) == "BAKERY CASE")
    check("name common stem", cg_core.name_from_values(["RA-12A", "RA-12B"]) == "RA-12")
    check("name fallback on all blank", cg_core.name_from_values(["", None], fallback="R9") == "R9")
    check("name fallback on no stem", cg_core.name_from_values(["AAA", "BBB"], fallback="R9") == "R9")
    check("name empty no fallback", cg_core.name_from_values([]) == "")

    # --- Space is offered via its own control, never the parameter combo ---
    space_rows = [
        {"group_values": {"Identity Mark": "RA-1", "Space": "101 - Office"}},
        {"group_values": {"Identity Mark": "RA-2", "Space": "102 - Lab"}},
    ]
    space_opts = cg_core.common_group_params(space_rows)
    check("space excluded from param combo", cg_core.SPACE_GROUP_KEY not in space_opts)
    check("space rows keep real params", "Identity Mark" in space_opts)

    # --- voltage snapping: standard-or-reject (input already volt-converted) ---
    check("snap 119.98 -> 120", cg_core.snap_voltage(119.98) == 120)
    check("snap 207.6 -> 208", cg_core.snap_voltage(207.6) == 208)
    check("snap 480.2 -> 480", cg_core.snap_voltage(480.2) == 480)
    check("snap 277 -> 277", cg_core.snap_voltage(277) == 277)
    check("snap 24 -> 24", cg_core.snap_voltage(24) == 24)
    # unconverted internal readings must never resolve to a nominal
    check("snap 1291 (unconv 120) -> None", cg_core.snap_voltage(1291) is None)
    check("snap 2239 (unconv 208) -> None", cg_core.snap_voltage(2239) is None)
    check("snap odd 190 -> None", cg_core.snap_voltage(190) is None)
    check("snap None -> None", cg_core.snap_voltage(None) is None)
    check("snap zero -> None", cg_core.snap_voltage(0) is None)

    # --- multi-panel candidate parsing ---
    check("parse multi", cg_core.parse_panel_candidates("RA, RB, RC, RD") ==
          ["RA", "RB", "RC", "RD"])
    check("parse mixed separators", cg_core.parse_panel_candidates("RA; RB / RC") ==
          ["RA", "RB", "RC"])
    check("parse dedupes", cg_core.parse_panel_candidates("RA, ra, RA") == ["RA"])
    check("parse single", cg_core.parse_panel_candidates("RA") == ["RA"])
    check("parse empty", cg_core.parse_panel_candidates("") == [])
    check("parse none", cg_core.parse_panel_candidates(None) == [])
    check("parse canonicalizes to known", cg_core.parse_panel_candidates(
        "ra, rb, RX", known_panels=["RA", "RB", "RC"]) == ["RA", "RB"])

    # --- geometry helpers ---
    check("centroid", cg_core.centroid([(0, 0, 0), (2, 4, 6)]) == (1.0, 2.0, 3.0))
    check("centroid skips None", cg_core.centroid([None, (2, 2, 2)]) == (2.0, 2.0, 2.0))
    check("centroid empty", cg_core.centroid([None]) is None)
    check("distance", cg_core.distance((0, 0, 0), (3, 4, 0)) == 5.0)
    check("distance None", cg_core.distance(None, (0, 0, 0)) is None)

    # --- multi-panel resolution: closest with room wins ---
    panel_info = {
        "RA": {"location": (0, 0, 0), "open_slots": 2},
        "RB": {"location": (100, 0, 0), "open_slots": 2},
        "RC": {"location": (200, 0, 0), "open_slots": None},  # unlimited
    }
    res = cg_core.resolve_panel_assignments(
        [{"group_key": "G1", "candidates": ["RA", "RB", "RC"],
          "centroid": (90, 0, 0), "poles": 1}],
        panel_info)
    check("closest with room", res["G1"]["panel"] == "RB" and res["G1"]["note"] == "")

    # reservation overflow: three 1P groups near RA (2 slots) -> third goes RB
    reqs = [
        {"group_key": "G%d" % i, "candidates": ["RA", "RB"],
         "centroid": (i, 0, 0), "poles": 1}
        for i in (1, 2, 3)
    ]
    res = cg_core.resolve_panel_assignments(reqs, panel_info)
    check("overflow to next closest",
          res["G1"]["panel"] == "RA" and res["G2"]["panel"] == "RA" and
          res["G3"]["panel"] == "RB")

    # pole-aware: a 3P circuit skips a panel with only 2 slots
    res = cg_core.resolve_panel_assignments(
        [{"group_key": "G1", "candidates": ["RA", "RB"],
          "centroid": (0, 0, 0), "poles": 3}],
        {"RA": {"location": (0, 0, 0), "open_slots": 2},
         "RB": {"location": (50, 0, 0), "open_slots": 3}})
    check("pole-aware skip", res["G1"]["panel"] == "RB")

    # no room anywhere -> blank + note
    res = cg_core.resolve_panel_assignments(
        [{"group_key": "G1", "candidates": ["RA"],
          "centroid": (0, 0, 0), "poles": 3}],
        {"RA": {"location": (0, 0, 0), "open_slots": 2}})
    check("no room -> blank + note",
          res["G1"]["panel"] is None and
          res["G1"]["note"] == cg_core.PANEL_NOTE_NO_ROOM)

    # no listed panel exists -> blank + note
    res = cg_core.resolve_panel_assignments(
        [{"group_key": "G1", "candidates": [], "centroid": None, "poles": 1}],
        panel_info)
    check("unknown panels -> blank + note",
          res["G1"]["panel"] is None and
          res["G1"]["note"] == cg_core.PANEL_NOTE_UNKNOWN)

    # unlimited-capacity panel is always eligible and never reserved out
    reqs = [{"group_key": "G%d" % i, "candidates": ["RC"],
             "centroid": (200, 0, 0), "poles": 3} for i in (1, 2, 3, 4)]
    res = cg_core.resolve_panel_assignments(reqs, panel_info)
    check("unlimited never fills",
          all(res["G%d" % i]["panel"] == "RC" for i in (1, 2, 3, 4)))

    # missing centroid still resolves (listed order, after measurable groups)
    res = cg_core.resolve_panel_assignments(
        [{"group_key": "G1", "candidates": ["RB", "RA"],
          "centroid": None, "poles": 1}],
        panel_info)
    check("no centroid -> listed order", res["G1"]["panel"] == "RB")

    print("\nAll tests passed.")


if __name__ == "__main__":
    run()
