# -*- coding: utf-8 -*-
"""Pure math for the Move with Monitored Parent action.

Locations are the JSON-safe dicts produced by ``location_service``
(feet + radians). Element_Linker pose fields use feet + degrees; the
conversion lives here so the Revit-facing service stays unit-free.
"""

from __future__ import print_function

import copy
import math

import comparison_engine
import models

TWO_PI = 2.0 * math.pi


def _valid_location(location):
    return (
        isinstance(location, dict)
        and location.get("state") == models.VALUE_VALID
    )


def _wrap_signed(angle):
    """Wrap an angle in radians to (-pi, pi]."""
    wrapped = math.fmod(float(angle), TWO_PI)
    if wrapped > math.pi:
        wrapped -= TWO_PI
    elif wrapped <= -math.pi:
        wrapped += TWO_PI
    return wrapped


def compute_follow_delta(parent_accepted, parent_current):
    """Translation + signed rotation delta of a monitored parent.

    Returns ``{"translation": (dx, dy, dz), "rotation_delta": radians,
    "pivot": (x, y, z)}`` (pivot = parent's current point) or ``None``
    when either location is missing/unsupported.
    """
    if not _valid_location(parent_accepted) or not _valid_location(parent_current):
        return None
    translation = (
        float(parent_current.get("x", 0.0)) - float(parent_accepted.get("x", 0.0)),
        float(parent_current.get("y", 0.0)) - float(parent_accepted.get("y", 0.0)),
        float(parent_current.get("z", 0.0)) - float(parent_accepted.get("z", 0.0)),
    )
    rotation_delta = _wrap_signed(
        float(parent_current.get("rotation", 0.0) or 0.0)
        - float(parent_accepted.get("rotation", 0.0) or 0.0)
    )
    pivot = (
        float(parent_current.get("x", 0.0)),
        float(parent_current.get("y", 0.0)),
        float(parent_current.get("z", 0.0)),
    )
    return {
        "translation": translation,
        "rotation_delta": rotation_delta,
        "pivot": pivot,
    }


def transform_point(point, delta):
    """Apply a follow delta to a child point: translate, then rotate about
    the pivot's Z axis. Mirrors MoveElement followed by RotateElement."""
    delta = delta or {}
    dx, dy, dz = delta.get("translation") or (0.0, 0.0, 0.0)
    px, py, _pz = delta.get("pivot") or (0.0, 0.0, 0.0)
    angle = float(delta.get("rotation_delta") or 0.0)
    x = float(point[0]) + dx
    y = float(point[1]) + dy
    z = float(point[2]) + dz
    cos_a = math.cos(angle)
    sin_a = math.sin(angle)
    rx = px + (x - px) * cos_a - (y - py) * sin_a
    ry = py + (x - px) * sin_a + (y - py) * cos_a
    return (rx, ry, z)


def significant_delta(delta, translation_tolerance, angular_tolerance):
    """True when the delta exceeds either tolerance (i.e. worth applying)."""
    if delta is None:
        return False
    dx, dy, dz = delta.get("translation") or (0.0, 0.0, 0.0)
    distance = math.sqrt(dx * dx + dy * dy + dz * dz)
    if distance > float(translation_tolerance or 0.0):
        return True
    return abs(float(delta.get("rotation_delta") or 0.0)) > float(
        angular_tolerance or 0.0
    )


def linker_pose_updates(child_location, parent_current):
    """Element_Linker field updates after a follow move.

    ``child_location`` is the child's post-move location dict; the parent
    pose fields are rewritten so MEPRFP Follow Parent does not compute a
    bogus second move. Returns a dict of linker field name -> new value.
    """
    updates = {}
    if _valid_location(child_location):
        updates["location_ft"] = [
            float(child_location.get("x", 0.0)),
            float(child_location.get("y", 0.0)),
            float(child_location.get("z", 0.0)),
        ]
        updates["rotation_deg"] = math.degrees(
            float(child_location.get("rotation", 0.0) or 0.0)
        ) % 360.0
    if _valid_location(parent_current):
        updates["parent_location_ft"] = [
            float(parent_current.get("x", 0.0)),
            float(parent_current.get("y", 0.0)),
            float(parent_current.get("z", 0.0)),
        ]
        updates["parent_rotation_deg"] = math.degrees(
            float(parent_current.get("rotation", 0.0) or 0.0)
        ) % 360.0
    return updates


def accept_parent_location_if_children_clean(tracking_set, parent_persistent_id):
    """Auto-accept a parent's location once every child that follows it has
    no outstanding location change. Returns True when acceptance happened."""
    tracking_set = tracking_set or {}
    elements = tracking_set.get("elements") or {}
    parent_key = str(parent_persistent_id or "")
    parent_record = elements.get(parent_key)
    if parent_record is None:
        return False
    for record in elements.values():
        if str(record.get("parent_persistent_id") or "") != parent_key:
            continue
        if record.get("state") == models.ELEMENT_REMOVED:
            continue
        if models.LOCATION_PROPERTY_KEY in list(
            record.get("changed_property_keys") or []
        ):
            return False
    parent_record["accepted_location"] = copy.deepcopy(
        parent_record.get("current_location")
    )
    comparison_engine.recompute_record(parent_record, tracking_set)
    return True
