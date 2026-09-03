# -*- coding: utf-8 -*-
"""Revit-console geometry checks for Tag by Example.

Run ``run_geometry_round_trip_tests()`` from a pyRevit console or debug
command.  These tests do not create tags or modify the document.
"""

import math
import os
import sys

from pyrevit import DB

from Snippets import revit_helpers
from Snippets.tag_geometry import points_are_equal
from Snippets.tag_geometry import snap_angle_radians
from Snippets.tag_host_transform import HostPlacementFrame
from Snippets.tag_host_transform import host_local_point_to_world
from Snippets.tag_host_transform import world_point_to_host_local


COMMAND_DIRECTORY = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIBRARY_DIRECTORY = os.path.join(COMMAND_DIRECTORY, "lib")
if LIBRARY_DIRECTORY not in sys.path:
    sys.path.append(LIBRARY_DIRECTORY)

from tag_by_example_events import _unique_integer_values


def _frame(origin, angle, mirrored=False):
    cosine = math.cos(angle)
    sine = math.sin(angle)
    axis_x = DB.XYZ(cosine, sine, 0.0)
    axis_y = DB.XYZ(-sine, cosine, 0.0)
    if mirrored:
        axis_y = axis_y.Multiply(-1.0)
    return HostPlacementFrame(
        origin=origin,
        basis_x=axis_x,
        basis_y=axis_y,
        basis_z=DB.XYZ.BasisZ,
        mirrored=mirrored,
        rotation=angle,
        source_kind="test",
    )


def _assert_round_trip(frame, local_point):
    world_point = host_local_point_to_world(local_point, frame)
    recovered_point = world_point_to_host_local(world_point, frame)
    assert points_are_equal(local_point, recovered_point), "Local/world round trip failed."


def run_geometry_round_trip_tests():
    local_point = DB.XYZ(2.25, -1.75, 0.5)
    frames = [
        _frame(DB.XYZ(0.0, 0.0, 0.0), 0.0),
        _frame(DB.XYZ(10.0, 4.0, 0.0), math.pi / 2.0),
        _frame(DB.XYZ(-3.0, 7.0, 0.0), math.radians(37.0)),
        _frame(DB.XYZ(6.0, -2.0, 0.0), math.radians(143.0), mirrored=True),
    ]
    for frame in frames:
        _assert_round_trip(frame, local_point)

    source_frame = _frame(DB.XYZ(1.0, 2.0, 0.0), math.radians(17.0))
    target_frame = _frame(DB.XYZ(20.0, -5.0, 0.0), math.radians(91.0), mirrored=True)
    source_world = host_local_point_to_world(local_point, source_frame)
    source_local = world_point_to_host_local(source_world, source_frame)
    target_world = host_local_point_to_world(source_local, target_frame)
    target_local = world_point_to_host_local(target_world, target_frame)
    assert points_are_equal(local_point, target_local), "Cross-host transform failed."

    assert abs(math.degrees(snap_angle_radians(math.radians(28.5))) - 30.0) < 1e-9
    assert abs(math.degrees(snap_angle_radians(math.radians(91.5))) - 90.0) < 1e-9
    assert abs(math.degrees(snap_angle_radians(math.radians(33.5))) - 33.5) < 1e-9
    assert abs(math.degrees(snap_angle_radians(math.radians(359.0))) - 0.0) < 1e-9
    return True


def run_element_id_boundary_tests():
    """Verify DTO normalization for numeric and native ElementId inputs."""
    numeric_value = 24680
    native_id = revit_helpers.elementid_from_value(numeric_value)
    normalized = _unique_integer_values([
        numeric_value,
        native_id,
        numeric_value,
        DB.ElementId.InvalidElementId,
        None,
    ])
    assert normalized == [numeric_value], (
        "Numeric and native ElementId normalization failed: {}".format(normalized)
    )
    return True


if __name__ == "__main__":
    print("Tag by Example geometry tests: {}".format(run_geometry_round_trip_tests()))
    print("Tag by Example ElementId tests: {}".format(run_element_id_boundary_tests()))
