# -*- coding: utf-8 -*-
"""Host-local coordinate frames used by Tag by Example.

This module is intentionally independent from ``_rotateutils.py``.  The
Orientation Tools keep their established behavior, while this module exposes
the more general point-frame operations needed by annotation placement.
"""

import math

from pyrevit import DB


def _length(vector):
    try:
        return vector.GetLength()
    except Exception:
        return 0.0


def _unit(vector, fallback):
    if vector is None or _length(vector) < 1e-9:
        return fallback
    return vector.Normalize()


def _project(vector, normal):
    return vector - normal.Multiply(vector.DotProduct(normal))


def _safe_xyz(vector, fallback):
    if vector is None:
        return fallback
    return vector


def _rotation_from_axis(axis):
    if axis is None:
        return 0.0
    return math.atan2(axis.Y, axis.X)


class HostPlacementFrame(object):
    """A right-handed 2D frame embedded in the owner view plane."""

    def __init__(self, origin, basis_x, basis_y, basis_z,
                 mirrored=False, facing_flipped=False, hand_flipped=False,
                 rotation=0.0, source_kind="unknown",
                 facing_orientation=None, hand_orientation=None):
        self.origin = origin
        self.basis_x = basis_x
        self.basis_y = basis_y
        self.basis_z = basis_z
        self.mirrored = bool(mirrored)
        self.facing_flipped = bool(facing_flipped)
        self.hand_flipped = bool(hand_flipped)
        self.rotation = rotation
        self.source_kind = source_kind
        self.facing_orientation = facing_orientation
        self.hand_orientation = hand_orientation


def _view_normal(view, fallback):
    try:
        return _unit(view.ViewDirection, fallback)
    except Exception:
        return fallback


def _location_origin(host, transform):
    try:
        return transform.Origin
    except Exception:
        pass
    try:
        location = host.Location
        if hasattr(location, "Point"):
            return location.Point
    except Exception:
        pass
    raise ValueError("Host has no resolvable placement origin.")


def get_host_placement_frame(host, view):
    """Build a placement frame for a point-based FamilyInstance host.

    Revit's instance transform already includes rotation and mirror state, so
    the returned basis is used directly.  Facing/hand flags are retained as
    metadata for reporting and future adapter-specific behavior; they are not
    manually applied a second time.
    """
    if host is None or not isinstance(host, DB.FamilyInstance):
        raise ValueError("Only FamilyInstance hosts are supported in Phase 1.")

    transform = None
    try:
        transform = host.GetTotalTransform()
    except Exception:
        try:
            transform = host.GetTransform()
        except Exception:
            transform = None
    if transform is None:
        raise ValueError("Host transform could not be resolved.")

    fallback_z = DB.XYZ.BasisZ
    normal = _view_normal(view, fallback_z)
    origin = _location_origin(host, transform)

    raw_x = _safe_xyz(getattr(transform, "BasisX", None), DB.XYZ.BasisX)
    raw_y = _safe_xyz(getattr(transform, "BasisY", None), DB.XYZ.BasisY)
    axis_x = _unit(_project(raw_x, normal), DB.XYZ.BasisX)
    axis_y = _unit(_project(raw_y, normal), DB.XYZ.BasisY)

    # Keep the projected Revit Y direction when it is available.  This is
    # important for mirrored instances; rebuilding Y from a cross product
    # would erase the transform's handedness information.
    if axis_y.DotProduct(_project(raw_y, normal)) < 0.0:
        axis_y = axis_y.Multiply(-1.0)

    mirrored = False
    facing_flipped = False
    hand_flipped = False
    facing_orientation = None
    hand_orientation = None
    rotation = _rotation_from_axis(axis_x)
    for name in ("Mirrored", "FacingFlipped", "HandFlipped"):
        try:
            value = bool(getattr(host, name))
        except Exception:
            value = False
        if name == "Mirrored":
            mirrored = value
        elif name == "FacingFlipped":
            facing_flipped = value
        else:
            hand_flipped = value

    try:
        facing_orientation = host.FacingOrientation
    except Exception:
        pass
    try:
        hand_orientation = host.HandOrientation
    except Exception:
        pass

    try:
        location = host.Location
        if hasattr(location, "Rotation"):
            rotation = float(location.Rotation)
    except Exception:
        pass

    return HostPlacementFrame(
        origin=origin,
        basis_x=axis_x,
        basis_y=axis_y,
        basis_z=normal,
        mirrored=mirrored,
        facing_flipped=facing_flipped,
        hand_flipped=hand_flipped,
        rotation=rotation,
        source_kind="FamilyInstance",
        facing_orientation=facing_orientation,
        hand_orientation=hand_orientation,
    )


def world_point_to_host_local(point, frame):
    """Convert a model/view point to frame-local XYZ coordinates."""
    if point is None:
        return None
    delta = point - frame.origin
    return DB.XYZ(
        delta.DotProduct(frame.basis_x),
        delta.DotProduct(frame.basis_y),
        delta.DotProduct(frame.basis_z),
    )


def host_local_point_to_world(point, frame):
    """Convert frame-local XYZ coordinates to a model/view point."""
    if point is None:
        return None
    return (frame.origin
            + frame.basis_x.Multiply(point.X)
            + frame.basis_y.Multiply(point.Y)
            + frame.basis_z.Multiply(point.Z))


def adapt_target_frame(source_frame, target_frame, adjust_rotation=True,
                       adjust_mirror=True, adjust_facing=True,
                       adjust_hand=True):
    """Apply user-selected transform switches without double-applying Revit state.

    Revit's instance transform remains the source of truth.  When a switch is
    disabled, this function deliberately removes that frame component from the
    basis used for placement; enabled components are never manually rotated a
    second time.
    """
    basis_x = target_frame.basis_x if adjust_rotation else source_frame.basis_x
    basis_y = target_frame.basis_y if adjust_rotation else source_frame.basis_y
    if not adjust_mirror and target_frame.mirrored:
        basis_x = basis_x.Multiply(-1.0)
    if not adjust_facing and target_frame.facing_flipped:
        basis_y = basis_y.Multiply(-1.0)
    if not adjust_hand and target_frame.hand_flipped:
        basis_x = basis_x.Multiply(-1.0)
    return HostPlacementFrame(
        origin=target_frame.origin,
        basis_x=basis_x,
        basis_y=basis_y,
        basis_z=target_frame.basis_z,
        mirrored=target_frame.mirrored,
        facing_flipped=target_frame.facing_flipped,
        hand_flipped=target_frame.hand_flipped,
        rotation=target_frame.rotation if adjust_rotation else source_frame.rotation,
        source_kind=target_frame.source_kind,
        facing_orientation=target_frame.facing_orientation,
        hand_orientation=target_frame.hand_orientation,
    )


def transform_point_between_hosts(point, source_frame, target_frame):
    """Transform a point through source-local coordinates into a target."""
    local_point = world_point_to_host_local(point, source_frame)
    return host_local_point_to_world(local_point, target_frame)
