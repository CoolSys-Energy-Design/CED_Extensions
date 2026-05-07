# -*- coding: utf-8 -*-
"""
Pure-Python anchor computation for space placement rules.

Given a ``SpaceGeometry`` (axis-aligned bounding box of a Revit Space
plus the placement-direction door anchors) and a ``PlacementRule``,
``anchor_points()`` returns the list of world points where one LED
copy should be placed *before* per-instance offsets are layered on
top.

Conventions
-----------
- All distances in *feet*. ``inset_inches`` and door offsets in the
  rule are converted on the fly.
- Coordinate axes: project N/S/E/W = +Y/-Y/+E/-X. The rule's anchor
  kinds (``n``, ``s``, ``e``, ``w``, ``ne``, ``nw``, ``se``, ``sw``)
  reference the bounding-box edges/corners along these axes.
- For ``center`` and edge/corner kinds, the returned z is the space's
  floor elevation. The LED's ``offsets[*]`` list lifts the element to
  its mounting height (e.g. 18 in for wall outlets).
- For ``door_relative``, one anchor is returned per door — fulfilling
  the "place at every door" decision. ``door_offset_inches.x`` runs
  along the door's *inward* normal (into the room), and
  ``door_offset_inches.y`` runs along the door's hinge-to-knob axis
  (90° CCW from inward, so positive y is to the door's "left" when
  looking from inside the room out through the door).

The Revit-API edge that builds a ``SpaceGeometry`` from a Revit
``Space`` element lives near the bottom of this module under a
``try/except`` import guard so the pure-logic core can be unit-tested
without Revit assemblies present.
"""

import math

from space_profile_model import (
    PlacementRule,
    KIND_CENTER,
    KIND_DOOR_RELATIVE,
    KIND_WALL_OPPOSITE_DOOR,
    KIND_WALL_RIGHT_OF_DOOR,
    KIND_WALL_LEFT_OF_DOOR,
    KIND_CORNER_FURTHEST_FROM_DOOR,
    KIND_CORNER_CLOSEST_TO_DOOR,
    DOOR_DEPENDENT_KINDS,
)


# ---------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------

INCHES_PER_FOOT = 12.0


def _in_to_ft(value):
    return float(value or 0.0) / INCHES_PER_FOOT


# ---------------------------------------------------------------------
# Pure-data record
# ---------------------------------------------------------------------

class SpaceGeometry(object):
    """Plain-data view of a Revit Space for placement.

    ``bbox`` is the axis-aligned bounding box of the Space's footprint
    in *feet*: ``((xmin, ymin), (xmax, ymax))``.

    ``floor_z`` is the elevation of the Space's level in *feet*.

    ``door_anchors`` is a list of ``(origin_xy, inward_normal_xy)``
    tuples — one per door bounding the Space. ``origin_xy`` is the
    door's location point (XY only). ``inward_normal_xy`` is a unit
    vector pointing INTO the Space along the door's facing axis.
    """

    __slots__ = ("bbox", "floor_z", "door_anchors", "name", "element_id")

    def __init__(self, bbox=None, floor_z=0.0, door_anchors=None,
                 name="", element_id=None):
        self.bbox = bbox  # ((xmin, ymin), (xmax, ymax)) in feet
        self.floor_z = float(floor_z or 0.0)
        self.door_anchors = list(door_anchors or [])
        self.name = name or ""
        self.element_id = element_id

    @property
    def x_min(self):
        return float(self.bbox[0][0]) if self.bbox else 0.0

    @property
    def y_min(self):
        return float(self.bbox[0][1]) if self.bbox else 0.0

    @property
    def x_max(self):
        return float(self.bbox[1][0]) if self.bbox else 0.0

    @property
    def y_max(self):
        return float(self.bbox[1][1]) if self.bbox else 0.0

    @property
    def x_center(self):
        return (self.x_min + self.x_max) / 2.0

    @property
    def y_center(self):
        return (self.y_min + self.y_max) / 2.0

    def __repr__(self):
        return "<SpaceGeometry id={} name={!r} bbox={} doors={}>".format(
            self.element_id, self.name, self.bbox, len(self.door_anchors)
        )


# ---------------------------------------------------------------------
# Anchor computation (pure logic)
# ---------------------------------------------------------------------

def anchor_points(rule, geom, door_anchor=None):
    """Return ``[(x, y, z), ...]`` anchor points for ``rule`` in ``geom``.

    All non-``center`` kinds depend on a reference door. Pass it in
    via ``door_anchor`` (the ``(origin_xy, inward_xy)`` tuple). When
    omitted, the first entry of ``geom.door_anchors`` is used —
    convenient for unit tests and for spaces with exactly one door.

    Returns an empty list when:

      * ``rule`` / ``geom`` / ``geom.bbox`` is missing.
      * The kind is door-dependent and no door is available
        (caller's cue to emit a comment-only plan).
      * The kind isn't one of the recognised values.
    """
    if rule is None or geom is None or geom.bbox is None:
        return []
    if not isinstance(rule, PlacementRule):
        rule = PlacementRule(dict(rule) if isinstance(rule, dict) else {})

    kind = rule.kind
    z = geom.floor_z

    if kind == KIND_CENTER:
        return [(geom.x_center, geom.y_center, z)]

    # Resolve the reference door. Caller-supplied wins; otherwise
    # fall back to the first door in geom.
    door = door_anchor
    if door is None and geom.door_anchors:
        door = geom.door_anchors[0]
    if door is None:
        return []

    # Normalise the door's inward direction so it always points INTO
    # the space. ``door_to_anchor`` derives ``inward`` by negating
    # ``FacingOrientation``, but that heuristic only holds when the
    # door family was placed with its facing pointing out of the room.
    # Doors placed flipped (or families authored with the opposite
    # convention) come through here with ``inward`` pointing away from
    # the space center, which mirrors every wall- and corner-relative
    # anchor onto the wrong side. The dot-product check is geometry-
    # truthful — flip when the recorded inward and the actual toward-
    # center direction disagree.
    door = _orient_door_inward(door, (geom.x_center, geom.y_center))

    if kind == KIND_DOOR_RELATIVE:
        return [_door_relative_point(rule, door, z)]

    if kind in (
        KIND_WALL_OPPOSITE_DOOR,
        KIND_WALL_RIGHT_OF_DOOR,
        KIND_WALL_LEFT_OF_DOOR,
    ):
        return [_wall_relative_point(kind, rule, geom, door, z)]

    if kind in (
        KIND_CORNER_FURTHEST_FROM_DOOR,
        KIND_CORNER_CLOSEST_TO_DOOR,
    ):
        return [_corner_relative_point(kind, rule, geom, door, z)]

    return []


def _door_relative_point(rule, door_anchor, z):
    """Single point at the door, optionally offset along the door's
    inward (door_offset_x) and sideways (door_offset_y) axes."""
    door_x = _in_to_ft(rule.door_offset_x_inches)
    door_y = _in_to_ft(rule.door_offset_y_inches)
    origin_xy, inward_xy = door_anchor
    ox, oy = float(origin_xy[0]), float(origin_xy[1])
    nx, ny = _normalize_xy(inward_xy)
    # Sideways = inward rotated 90° CCW: (-ny, nx).
    sx, sy = -ny, nx
    x = ox + door_x * nx + door_y * sx
    y = oy + door_x * ny + door_y * sy
    return (x, y, z)


def _wall_relative_point(kind, rule, geom, door_anchor, z):
    """Anchor at the midpoint of the wall identified relative to the
    door (opposite / right / left), inset toward the room interior
    by ``rule.inset_inches``.

    "Right" / "left" are taken from the perspective of someone
    standing in the doorway facing into the room (i.e. along the
    door's inward normal). 90° clockwise from inward is "right".
    """
    inset = _in_to_ft(rule.inset_inches)
    _door_xy, inward_xy = door_anchor
    nx, ny = _normalize_xy(inward_xy)
    xmin = geom.x_min
    xmax = geom.x_max
    ymin = geom.y_min
    ymax = geom.y_max
    cx = geom.x_center
    cy = geom.y_center

    # Pick which bbox edge is nearest each cardinal of the door's
    # frame. ``axis`` is which world axis the inward normal aligns
    # most strongly with.
    if abs(nx) > abs(ny):
        # Door wall is on the X axis (east or west).
        if nx > 0:
            # Door on west wall (inward = +X). Right (90° CW from +X) = -Y.
            opposite = (xmax - inset, cy, z)
            right = (cx, ymin + inset, z)
            left = (cx, ymax - inset, z)
        else:
            # Door on east wall (inward = -X). Right (90° CW from -X) = +Y.
            opposite = (xmin + inset, cy, z)
            right = (cx, ymax - inset, z)
            left = (cx, ymin + inset, z)
    else:
        if ny > 0:
            # Door on south wall (inward = +Y). Right (90° CW from +Y) = +X.
            opposite = (cx, ymax - inset, z)
            right = (xmax - inset, cy, z)
            left = (xmin + inset, cy, z)
        else:
            # Door on north wall (inward = -Y). Right (90° CW from -Y) = -X.
            opposite = (cx, ymin + inset, z)
            right = (xmin + inset, cy, z)
            left = (xmax - inset, cy, z)

    if kind == KIND_WALL_OPPOSITE_DOOR:
        return opposite
    if kind == KIND_WALL_RIGHT_OF_DOOR:
        return right
    if kind == KIND_WALL_LEFT_OF_DOOR:
        return left
    return (cx, cy, z)  # unreachable


def _corner_relative_point(kind, rule, geom, door_anchor, z):
    """Anchor at the bbox corner that's closest to / furthest from
    the door, inset diagonally toward the room interior."""
    inset = _in_to_ft(rule.inset_inches)
    (door_x, door_y), _inward = door_anchor
    xmin = geom.x_min
    xmax = geom.x_max
    ymin = geom.y_min
    ymax = geom.y_max
    cx = geom.x_center
    cy = geom.y_center

    corners = [
        (xmin, ymin),
        (xmin, ymax),
        (xmax, ymin),
        (xmax, ymax),
    ]

    def _dist_sq(c):
        return (c[0] - door_x) ** 2 + (c[1] - door_y) ** 2

    if kind == KIND_CORNER_CLOSEST_TO_DOOR:
        target = min(corners, key=_dist_sq)
    else:  # KIND_CORNER_FURTHEST_FROM_DOOR
        target = max(corners, key=_dist_sq)

    # Inset diagonally toward the room center.
    tx, ty = target
    tx += inset if tx < cx else -inset
    ty += inset if ty < cy else -inset
    return (tx, ty, z)


def _normalize_xy(vec):
    if vec is None:
        return (1.0, 0.0)
    vx, vy = float(vec[0]), float(vec[1])
    length = math.sqrt(vx * vx + vy * vy)
    if length < 1e-9:
        return (1.0, 0.0)
    return (vx / length, vy / length)


def _orient_door_inward(door_anchor, center_xy):
    """Return ``door_anchor`` with ``inward`` flipped if it points away
    from ``center_xy``.

    Used by ``anchor_points`` to make door-relative geometry insensitive
    to which side the door's family was placed on. A zero or near-zero
    dot product (door at the centroid, or inward perpendicular to the
    toward-center direction) leaves the original orientation alone —
    we can't tell which side is "in" in that degenerate case, so we
    don't guess.
    """
    if door_anchor is None or center_xy is None:
        return door_anchor
    try:
        origin_xy, inward_xy = door_anchor
        ox, oy = float(origin_xy[0]), float(origin_xy[1])
        ix, iy = float(inward_xy[0]), float(inward_xy[1])
        cx, cy = float(center_xy[0]), float(center_xy[1])
    except (TypeError, ValueError, IndexError):
        return door_anchor
    vx = cx - ox
    vy = cy - oy
    if (ix * vx + iy * vy) < 0.0:
        return ((ox, oy), (-ix, -iy))
    return door_anchor


# ---------------------------------------------------------------------
# Multi-LED expansion
# ---------------------------------------------------------------------

def expand_led_placements(led, geom, door_anchor=None):
    """Return ``[(x, y, z, rotation_deg), ...]`` for one LED in one space.

    Multiplies the rule's anchor set against the LED's per-instance
    ``offsets`` list. ``door_anchor`` is the user-chosen reference
    door for door-dependent kinds; defaults to the first door of
    ``geom`` when omitted.
    """
    rule = led.placement_rule
    anchors = anchor_points(rule, geom, door_anchor=door_anchor)
    if not anchors:
        return []

    offsets = led.offsets or []
    out = []
    if not offsets:
        # No per-instance offsets — one element per anchor at z=anchor.z.
        for ax, ay, az in anchors:
            out.append((ax, ay, az, 0.0))
        return out

    for ax, ay, az in anchors:
        for off in offsets:
            ox = _in_to_ft(off.x_inches)
            oy = _in_to_ft(off.y_inches)
            oz = _in_to_ft(off.z_inches)
            rot = float(off.rotation_deg or 0.0)
            out.append((ax + ox, ay + oy, az + oz, rot))
    return out


# ---------------------------------------------------------------------
# Revit-edge: build SpaceGeometry from a Revit Space element
# ---------------------------------------------------------------------
#
# Wrapped in a try/except so the pure-logic above runs in plain CPython
# tests without pyrevit / Autodesk.Revit.DB on the path.

try:
    import clr  # noqa: F401
    from Autodesk.Revit.DB import (
        BuiltInCategory,
        FilteredElementCollector,
        RevitLinkInstance,
        SpatialElementBoundaryOptions,
        XYZ,
    )
    _HAS_REVIT = True
except Exception:  # pragma: no cover -- only true outside Revit
    BuiltInCategory = None
    FilteredElementCollector = None
    RevitLinkInstance = None
    SpatialElementBoundaryOptions = None
    XYZ = None
    _HAS_REVIT = False


def build_space_geometry(doc, space):
    """Build a ``SpaceGeometry`` from a Revit Space.

    Walks the Space's boundary segments to compute the XY bounding box
    and collects every door hosted in a wall bounding the Space.
    Returns ``None`` if the Space is unplaced (no boundary).
    """
    if not _HAS_REVIT:
        raise RuntimeError(
            "build_space_geometry requires Revit; only the pure-logic "
            "anchor_points/expand_led_placements run outside Revit."
        )

    if doc is None or space is None:
        return None

    bbox = _space_bbox(space)
    if bbox is None:
        return None

    floor_z = _space_floor_z(doc, space)
    door_anchors = _space_door_anchors(doc, space)

    name = ""
    try:
        name = str(getattr(space, "Name", "") or "").strip()
    except Exception:
        name = ""

    eid = None
    try:
        eid = _element_id_int(getattr(space, "Id", None))
    except Exception:
        eid = None

    return SpaceGeometry(
        bbox=bbox,
        floor_z=floor_z,
        door_anchors=door_anchors,
        name=name,
        element_id=eid,
    )


def _element_id_int(elem_id):
    if elem_id is None:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(elem_id, attr)
        except Exception:
            v = None
        if v is None:
            continue
        try:
            return int(v)
        except Exception:
            continue
    return None


def _space_bbox(space):
    """XY axis-aligned bounding box (in feet) computed from boundary segments."""
    try:
        opts = SpatialElementBoundaryOptions()
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None
    if not loops:
        return None

    xs = []
    ys = []
    for loop in loops:
        for seg in loop:
            try:
                curve = seg.GetCurve()
            except Exception:
                continue
            for sample_t in (0.0, 0.25, 0.5, 0.75, 1.0):
                try:
                    pt = curve.Evaluate(sample_t, True)
                except Exception:
                    pt = None
                if pt is None:
                    continue
                xs.append(pt.X)
                ys.append(pt.Y)
    if not xs or not ys:
        return None
    return ((min(xs), min(ys)), (max(xs), max(ys)))


def _space_floor_z(doc, space):
    try:
        lvl_id = space.LevelId
    except Exception:
        return 0.0
    if lvl_id is None:
        return 0.0
    try:
        lvl = doc.GetElement(lvl_id)
    except Exception:
        return 0.0
    if lvl is None:
        return 0.0
    try:
        return float(lvl.Elevation or 0.0)
    except Exception:
        return 0.0


def _space_door_anchors(doc, space):
    """Return ``[(origin_xy, inward_normal_xy), ...]`` for doors at this space.

    Two-tier search:

      1. Doors *hosted* by walls that ``GetBoundarySegments`` reports
         as bounding this Space. Cheapest path; works for projects
         where architecture lives in the host doc.
      2. Doors in any *linked* Revit instance whose location, after
         transforming through the link's ``GetTotalTransform()``,
         falls within ~1 ft of one of the Space's boundary curves.
         Required when architecture is in a linked model — the
         host wall id check fails because the door's host wall lives
         in the link's id namespace, not this doc's.

    Both tiers are unioned. Duplicates (same physical door appearing
    in host + a link) are deduplicated by location proximity.
    """
    try:
        opts = SpatialElementBoundaryOptions()
        loops = space.GetBoundarySegments(opts)
    except Exception:
        loops = None
    if not loops:
        return []

    wall_ids = set()
    boundary_curves = []
    for loop in loops:
        for seg in loop:
            try:
                wid = seg.ElementId
            except Exception:
                wid = None
            if wid is not None:
                wid_int = _element_id_int(wid)
                # InvalidElementId reports as -1; ignore it.
                if wid_int is not None and wid_int > 0:
                    wall_ids.add(wid_int)
            try:
                curve = seg.GetCurve()
            except Exception:
                curve = None
            if curve is not None:
                boundary_curves.append(curve)

    out = []

    # ----- Tier 1: host doors hosted in our boundary walls ----------
    if wall_ids:
        try:
            host_doors = (
                FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .WhereElementIsNotElementType()
            )
        except Exception:
            host_doors = []
        for door in host_doors:
            host = getattr(door, "Host", None)
            if host is None:
                continue
            host_id = _element_id_int(getattr(host, "Id", None))
            if host_id not in wall_ids:
                continue
            anchor = door_to_anchor(door, transform=None)
            if anchor is not None:
                out.append(anchor)

    # ----- Tier 2: linked doors near the space's boundary -----------
    if boundary_curves and RevitLinkInstance is not None:
        try:
            link_collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
            link_instances = list(link_collector)
        except Exception:
            link_instances = []
        for link in link_instances:
            try:
                link_doc = link.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                continue
            try:
                transform = link.GetTotalTransform()
            except Exception:
                transform = None
            try:
                linked_doors = (
                    FilteredElementCollector(link_doc)
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .WhereElementIsNotElementType()
                )
            except Exception:
                continue
            for door in linked_doors:
                anchor = door_to_anchor(door, transform=transform)
                if anchor is None:
                    continue
                origin_xy = anchor[0]
                if not _point_near_any_curve(origin_xy, boundary_curves, tol=1.0):
                    continue
                if _origin_already_seen(origin_xy, out, tol=0.1):
                    continue
                out.append(anchor)

    return out


def door_to_anchor(door, transform=None):
    """Return ``(origin_xy, inward_xy)`` for a single Door element.

    ``transform`` is the link's total transform when the door lives
    in a linked doc; ``None`` for host doors. Returns ``None`` when
    the door has no Location.Point or no FacingOrientation.

    Public so the pre-placement door-picker (which calls
    ``Selection.PickObject``) can resolve the picked Reference to
    the same anchor tuple shape the workflow expects.
    """
    try:
        loc = door.Location
        pt = getattr(loc, "Point", None)
    except Exception:
        pt = None
    if pt is None:
        return None
    try:
        facing = door.FacingOrientation
    except Exception:
        facing = None
    if facing is None:
        return None
    if transform is not None:
        try:
            pt = transform.OfPoint(pt)
        except Exception:
            pass
        try:
            facing = transform.OfVector(facing)
        except Exception:
            pass
    origin_xy = (float(pt.X), float(pt.Y))
    inward = (-float(facing.X), -float(facing.Y))
    return (origin_xy, inward)


def _point_near_any_curve(point_xy, curves, tol=1.0):
    """True if ``point_xy`` (X, Y in feet) is within ``tol`` ft of
    the closest point on any of ``curves`` (boundary segments)."""
    if not curves or XYZ is None:
        return False
    px, py = point_xy
    for curve in curves:
        try:
            # Use the curve's start Z so Curve.Project doesn't
            # disqualify a coplanar test point on the Z axis.
            start = curve.Evaluate(0.0, True)
            test = XYZ(px, py, start.Z if start is not None else 0.0)
        except Exception:
            continue
        try:
            res = curve.Project(test)
        except Exception:
            res = None
        if res is None:
            continue
        try:
            cp = res.XYZPoint
        except Exception:
            cp = None
        if cp is None:
            continue
        dx = cp.X - px
        dy = cp.Y - py
        if (dx * dx + dy * dy) <= (tol * tol):
            return True
    return False


def _origin_already_seen(point_xy, anchors, tol=0.1):
    """True if any existing anchor's origin is within ``tol`` ft of
    ``point_xy`` — used to dedupe host vs linked sightings of the
    same physical door."""
    px, py = point_xy
    for (origin_xy, _inward) in anchors or ():
        try:
            ox, oy = origin_xy
        except Exception:
            continue
        dx = ox - px
        dy = oy - py
        if (dx * dx + dy * dy) <= (tol * tol):
            return True
    return False
