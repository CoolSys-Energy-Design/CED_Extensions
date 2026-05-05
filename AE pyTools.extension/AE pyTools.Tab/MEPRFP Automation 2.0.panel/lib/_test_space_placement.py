# -*- coding: utf-8 -*-
"""Tests for space_placement (pure-logic anchor computation)."""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from space_placement import (
    SpaceGeometry,
    anchor_points,
    expand_led_placements,
)
from space_profile_model import (
    PlacementRule,
    SpaceLED,
    KIND_CENTER,
    KIND_DOOR_RELATIVE,
    KIND_WALL_OPPOSITE_DOOR,
    KIND_WALL_RIGHT_OF_DOOR,
    KIND_WALL_LEFT_OF_DOOR,
    KIND_CORNER_FURTHEST_FROM_DOOR,
    KIND_CORNER_CLOSEST_TO_DOOR,
    DOOR_DEPENDENT_KINDS,
    PLACEMENT_KINDS,
)


_FAILS = []


def _check(label, cond, detail=""):
    if cond:
        print("  PASS  {}".format(label))
    else:
        print("  FAIL  {}  {}".format(label, detail))
        _FAILS.append(label)


def _close(a, b, eps=1e-6):
    return abs(a - b) < eps


def _make_geom(xmin=0, ymin=0, xmax=20, ymax=10, z=100, doors=None, name=""):
    return SpaceGeometry(
        bbox=((float(xmin), float(ymin)), (float(xmax), float(ymax))),
        floor_z=float(z),
        door_anchors=doors or [],
        name=name,
    )


# ---------------------------------------------------------------------
# Inventory
# ---------------------------------------------------------------------

def test_kinds_inventory():
    print("\n[kinds] inventory + door-dependence")
    expected = {
        KIND_CENTER,
        KIND_DOOR_RELATIVE,
        KIND_WALL_OPPOSITE_DOOR,
        KIND_WALL_RIGHT_OF_DOOR,
        KIND_WALL_LEFT_OF_DOOR,
        KIND_CORNER_FURTHEST_FROM_DOOR,
        KIND_CORNER_CLOSEST_TO_DOOR,
    }
    _check("PLACEMENT_KINDS == expected set", set(PLACEMENT_KINDS) == expected)
    _check("center NOT door-dependent", KIND_CENTER not in DOOR_DEPENDENT_KINDS)
    _check("door_relative IS door-dependent",
           KIND_DOOR_RELATIVE in DOOR_DEPENDENT_KINDS)
    _check("wall_opposite_door IS door-dependent",
           KIND_WALL_OPPOSITE_DOOR in DOOR_DEPENDENT_KINDS)
    _check("corner_closest_to_door IS door-dependent",
           KIND_CORNER_CLOSEST_TO_DOOR in DOOR_DEPENDENT_KINDS)


# ---------------------------------------------------------------------
# center
# ---------------------------------------------------------------------

def test_center():
    print("\n[anchor] center — no door needed")
    g = _make_geom(0, 0, 20, 10, 100)  # 20x10 box
    pts = anchor_points(PlacementRule({"kind": KIND_CENTER}), g)
    _check("one point", len(pts) == 1)
    x, y, z = pts[0]
    _check("center x", _close(x, 10.0))
    _check("center y", _close(y, 5.0))
    _check("center z=floor", _close(z, 100.0))


# ---------------------------------------------------------------------
# door_relative — door-frame offset semantic preserved
# ---------------------------------------------------------------------

def test_door_relative_zero_offset():
    print("\n[anchor] door_relative at door origin (no offset)")
    doors = [((10.0, 0.0), (0.0, 1.0))]  # south-wall door
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(PlacementRule({"kind": KIND_DOOR_RELATIVE}), g)
    _check("one anchor", len(pts) == 1)
    x, y, _z = pts[0]
    _check("at door x", _close(x, 10.0))
    _check("at door y", _close(y, 0.0))


def test_door_relative_x_pushes_inward():
    print("\n[anchor] door_relative x offset = inward")
    doors = [((10.0, 0.0), (0.0, 1.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(
        PlacementRule({
            "kind": KIND_DOOR_RELATIVE,
            "door_offset_inches": {"x": 12, "y": 0},
        }), g,
    )
    x, y, _z = pts[0]
    _check("y bumped +1ft", _close(y, 1.0))
    _check("x unchanged", _close(x, 10.0))


# ---------------------------------------------------------------------
# wall_opposite_door — opposite wall midpoint, inset toward room
# ---------------------------------------------------------------------

def test_wall_opposite_south_door():
    print("\n[anchor] wall_opposite_door — south door -> north wall midpt")
    doors = [((10.0, 0.0), (0.0, 1.0))]  # south-wall door
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(
        PlacementRule({"kind": KIND_WALL_OPPOSITE_DOOR, "inset_inches": 12}),
        g,
    )
    x, y, _z = pts[0]
    _check("x = bbox center", _close(x, 10.0))
    _check("y = ymax - 1ft inset", _close(y, 9.0))


def test_wall_opposite_west_door():
    print("\n[anchor] wall_opposite_door — west door -> east wall midpt")
    doors = [((0.0, 5.0), (1.0, 0.0))]  # west-wall door
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    pts = anchor_points(
        PlacementRule({"kind": KIND_WALL_OPPOSITE_DOOR, "inset_inches": 12}),
        g,
    )
    x, y, _z = pts[0]
    _check("x = xmax - 1ft inset", _close(x, 19.0))
    _check("y = bbox center", _close(y, 5.0))


# ---------------------------------------------------------------------
# wall_right_of_door / wall_left_of_door
# ---------------------------------------------------------------------

def test_wall_right_left_for_south_door():
    print("\n[anchor] wall_right/left — south door (looking N: right=east, left=west)")
    doors = [((10.0, 0.0), (0.0, 1.0))]  # south-wall door
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    rp = anchor_points(
        PlacementRule({"kind": KIND_WALL_RIGHT_OF_DOOR, "inset_inches": 6}),
        g,
    )[0]
    lp = anchor_points(
        PlacementRule({"kind": KIND_WALL_LEFT_OF_DOOR, "inset_inches": 6}),
        g,
    )[0]
    _check("right.x = xmax - 0.5ft", _close(rp[0], 19.5))
    _check("right.y = bbox center", _close(rp[1], 5.0))
    _check("left.x = xmin + 0.5ft", _close(lp[0], 0.5))
    _check("left.y = bbox center", _close(lp[1], 5.0))


def test_wall_right_left_for_north_door():
    print("\n[anchor] wall_right/left — north door (looking S: right=west, left=east)")
    doors = [((10.0, 10.0), (0.0, -1.0))]  # north-wall door
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    rp = anchor_points(
        PlacementRule({"kind": KIND_WALL_RIGHT_OF_DOOR, "inset_inches": 6}),
        g,
    )[0]
    lp = anchor_points(
        PlacementRule({"kind": KIND_WALL_LEFT_OF_DOOR, "inset_inches": 6}),
        g,
    )[0]
    _check("right.x = xmin + 0.5ft (west wall)", _close(rp[0], 0.5))
    _check("right.y = bbox center", _close(rp[1], 5.0))
    _check("left.x = xmax - 0.5ft (east wall)", _close(lp[0], 19.5))
    _check("left.y = bbox center", _close(lp[1], 5.0))


def test_wall_right_left_for_east_door():
    print("\n[anchor] wall_right/left — east door (looking W: right=south, left=north)")
    doors = [((20.0, 5.0), (-1.0, 0.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    rp = anchor_points(
        PlacementRule({"kind": KIND_WALL_RIGHT_OF_DOOR, "inset_inches": 0}),
        g,
    )[0]
    lp = anchor_points(
        PlacementRule({"kind": KIND_WALL_LEFT_OF_DOOR, "inset_inches": 0}),
        g,
    )[0]
    _check("right.y = ymax (north)", _close(rp[1], 10.0))
    _check("left.y = ymin (south)", _close(lp[1], 0.0))


# ---------------------------------------------------------------------
# corner_closest_to_door / corner_furthest_from_door
# ---------------------------------------------------------------------

def test_corner_closest_furthest():
    print("\n[anchor] corners — closest / furthest relative to door")
    # Door near south-east corner — closest is SE, furthest is NW.
    doors = [((18.0, 0.0), (0.0, 1.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    closest = anchor_points(
        PlacementRule({"kind": KIND_CORNER_CLOSEST_TO_DOOR, "inset_inches": 0}),
        g,
    )[0]
    furthest = anchor_points(
        PlacementRule({"kind": KIND_CORNER_FURTHEST_FROM_DOOR, "inset_inches": 0}),
        g,
    )[0]
    _check("closest = SE corner", _close(closest[0], 20.0) and _close(closest[1], 0.0))
    _check("furthest = NW corner", _close(furthest[0], 0.0) and _close(furthest[1], 10.0))


def test_corner_inset_diagonal():
    print("\n[anchor] corner inset pushes diagonally toward room center")
    # Door at south-east, inset 6 in. Closest corner SE = (20, 0); inset
    # toward center -> (-0.5, +0.5) = (19.5, 0.5).
    doors = [((18.0, 0.0), (0.0, 1.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    closest = anchor_points(
        PlacementRule({"kind": KIND_CORNER_CLOSEST_TO_DOOR, "inset_inches": 6}),
        g,
    )[0]
    _check("inset x", _close(closest[0], 19.5))
    _check("inset y", _close(closest[1], 0.5))


# ---------------------------------------------------------------------
# Door selection: caller-supplied door_anchor wins over geom default
# ---------------------------------------------------------------------

def test_caller_supplied_door_wins():
    print("\n[anchor] caller-supplied door overrides geom default")
    # Two doors: south + north. Default would use first (south); we
    # override with the second (north) and expect different result.
    doors = [
        ((10.0, 0.0), (0.0, 1.0)),    # south
        ((10.0, 10.0), (0.0, -1.0)),  # north
    ]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    rule = PlacementRule({"kind": KIND_WALL_OPPOSITE_DOOR, "inset_inches": 0})

    default_pt = anchor_points(rule, g)[0]
    overridden_pt = anchor_points(rule, g, door_anchor=doors[1])[0]
    _check("default (south door) picks north wall", _close(default_pt[1], 10.0))
    _check("override (north door) picks south wall", _close(overridden_pt[1], 0.0))


def test_door_dependent_with_no_door_returns_empty():
    print("\n[anchor] door-dependent kinds with no door -> []")
    g = _make_geom()  # no doors
    for kind in (
        KIND_DOOR_RELATIVE,
        KIND_WALL_OPPOSITE_DOOR,
        KIND_WALL_RIGHT_OF_DOOR,
        KIND_WALL_LEFT_OF_DOOR,
        KIND_CORNER_CLOSEST_TO_DOOR,
        KIND_CORNER_FURTHEST_FROM_DOOR,
    ):
        pts = anchor_points(PlacementRule({"kind": kind}), g)
        _check("{} -> []".format(kind), pts == [])


def test_unknown_kind_returns_empty():
    print("\n[anchor] unknown kind returns []")
    g = _make_geom(doors=[((10.0, 0.0), (0.0, 1.0))])
    pts = anchor_points(PlacementRule({"kind": "no_such_kind"}), g)
    _check("no anchors", pts == [])


# ---------------------------------------------------------------------
# expand_led_placements
# ---------------------------------------------------------------------

def test_expand_led_no_offsets_yields_one_per_anchor():
    print("\n[expand] LED with no offsets -> one placement per anchor")
    led = SpaceLED({
        "id": "L1",
        "placement_rule": {"kind": KIND_CENTER},
    })
    g = _make_geom(0, 0, 20, 10, 100)
    out = expand_led_placements(led, g)
    _check("one placement", len(out) == 1)
    x, y, z, rot = out[0]
    _check("at center", _close(x, 10.0) and _close(y, 5.0))
    _check("z=floor", _close(z, 100.0))
    _check("rotation 0", _close(rot, 0.0))


def test_expand_door_relative_one_door_two_offsets():
    print("\n[expand] door_relative with 1 door + 2 offsets -> 2 placements")
    doors = [((10.0, 0.0), (0.0, 1.0))]
    g = _make_geom(0, 0, 20, 10, 100, doors=doors)
    led = SpaceLED({
        "id": "L1",
        "placement_rule": {"kind": KIND_DOOR_RELATIVE},
        "offsets": [
            {"z_inches": 18},
            {"z_inches": 96, "rotation_deg": 90},
        ],
    })
    out = expand_led_placements(led, g)
    _check("2 placements", len(out) == 2)
    zs = sorted({round(p[2], 4) for p in out})
    _check("two z values (101.5, 108)", zs == [101.5, 108.0])


def test_expand_unknown_kind_yields_empty():
    print("\n[expand] unknown rule kind -> []")
    led = SpaceLED({
        "id": "L1",
        "placement_rule": {"kind": "no_such_kind"},
        "offsets": [{"z_inches": 18}],
    })
    g = _make_geom(doors=[((10.0, 0.0), (0.0, 1.0))])
    out = expand_led_placements(led, g)
    _check("no placements", out == [])


# ---------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------

def main():
    print("Running space_placement tests")
    test_kinds_inventory()
    test_center()
    test_door_relative_zero_offset()
    test_door_relative_x_pushes_inward()
    test_wall_opposite_south_door()
    test_wall_opposite_west_door()
    test_wall_right_left_for_south_door()
    test_wall_right_left_for_north_door()
    test_wall_right_left_for_east_door()
    test_corner_closest_furthest()
    test_corner_inset_diagonal()
    test_caller_supplied_door_wins()
    test_door_dependent_with_no_door_returns_empty()
    test_unknown_kind_returns_empty()
    test_expand_led_no_offsets_yields_one_per_anchor()
    test_expand_door_relative_one_door_two_offsets()
    test_expand_unknown_kind_yields_empty()

    print("")
    if _FAILS:
        print("FAILED: {} test(s) -- {}".format(len(_FAILS), _FAILS))
        sys.exit(1)
    print("All space_placement tests passed.")


if __name__ == "__main__":
    main()
