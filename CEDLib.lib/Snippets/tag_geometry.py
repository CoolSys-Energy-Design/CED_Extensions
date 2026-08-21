# -*- coding: utf-8 -*-
"""Pure tag-placement geometry models and transforms."""

import math

from pyrevit import DB

from Snippets.tag_host_transform import host_local_point_to_world
from Snippets.tag_host_transform import world_point_to_host_local


def _primary_angle_degrees():
    values = set()
    for increment in (30, 45, 60, 90):
        for multiplier in range(0, 13):
            value = increment * multiplier
            if value <= 360:
                values.add(value % 360)
    return sorted(values)


PRIMARY_ANGLE_DEGREES = _primary_angle_degrees()


def snap_angle_radians(angle, tolerance_degrees=3.0):
    """Snap an angle near a common model angle, using circular distance."""
    if angle is None:
        return None
    angle_degrees = math.degrees(float(angle)) % 360.0
    closest_angle = None
    closest_distance = None
    for primary_angle in PRIMARY_ANGLE_DEGREES:
        distance = abs(angle_degrees - primary_angle)
        distance = min(distance, 360.0 - distance)
        if closest_distance is None or distance < closest_distance:
            closest_angle = primary_angle
            closest_distance = distance
    if closest_distance is not None and closest_distance <= float(tolerance_degrees):
        return math.radians(closest_angle)
    return angle


class TagGeometry(object):
    def __init__(self):
        self.head_local = None
        self.elbow_local = None
        self.end_local = None
        self.has_leader = False
        self.leader_end_condition = None
        self.orientation = None
        self.rotation_angle = None
        self.source_rotation = 0.0
        self.tag_type_id = None


def geometry_from_world_points(head, elbow, end, source_frame):
    geometry = TagGeometry()
    geometry.head_local = world_point_to_host_local(head, source_frame)
    geometry.elbow_local = world_point_to_host_local(elbow, source_frame)
    geometry.end_local = world_point_to_host_local(end, source_frame)
    return geometry


def geometry_to_world_points(geometry, target_frame):
    return {
        "head": host_local_point_to_world(geometry.head_local, target_frame),
        "elbow": host_local_point_to_world(geometry.elbow_local, target_frame),
        "end": host_local_point_to_world(geometry.end_local, target_frame),
    }


def copy_geometry_for_target(example_geometry, source_frame, target_frame,
                             preserve_offset=True, adjust_rotation=True):
    """Return target-world points from saved example geometry.

    ``source_frame`` is accepted to make the call explicit and testable.  The
    geometry is already local, so preserving the relative placement is a
    local-to-target conversion rather than a global XYZ subtraction.
    """
    if not preserve_offset:
        zero_point = DB.XYZ(0.0, 0.0, 0.0)
        return {"head": target_frame.origin + zero_point,
                "elbow": None,
                "end": None}
    points = geometry_to_world_points(example_geometry, target_frame)
    if adjust_rotation and example_geometry.rotation_angle is not None:
        angle_delta = target_frame.rotation - source_frame.rotation
        points["rotation_angle"] = snap_angle_radians(
            example_geometry.rotation_angle + angle_delta
        )
    else:
        points["rotation_angle"] = snap_angle_radians(0.0)
    return points


def points_are_equal(first, second, tolerance=1e-6):
    if first is None or second is None:
        return first is None and second is None
    return first.DistanceTo(second) <= tolerance
