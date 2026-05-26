# pf_gym_recepts.py — Planet Fitness MCP skills library
# Detect powered gym-equipment blocks in the linked CAD and compute per-block placement points.
# IronPython 2.7. Caller manages Transaction.
#
# Pattern (E101 reference + user feedback on Charlotte project):
#   - A powered block has a small closed polyline (the "power bar center square")
#     near its lateral centerline at a moderate offset from origin.
#   - The receptacle is placed at the NORTH EDGE CENTER of the block's geometry bbox.
#   - The receptacle faces world +Y (north). [user instruction for current project]
#
# Tunable constants for future PF projects.

import math
from Autodesk.Revit import DB

# Detector constants
POWER_BAR_LOCAL_Y_TOL = 0.5     # |local_y| < this  -> on the lateral centerline
POWER_BAR_LOCAL_X_MIN = 0.3     # local_x must be at least this far from origin (offset)
POWER_BAR_LOCAL_X_MAX = 4.0     # ...but not too far (still inside the block)
SMALL_SHAPE_MIN       = 0.05    # ft — minimum bbox extent for "small closed polyline"
SMALL_SHAPE_MAX       = 0.7     # ft — maximum bbox extent

# Placement defaults (override per project)
DEFAULT_FACING_RAD    = 0.0     # rotation_rad applied at placement. 0 = family default.
DEFAULT_MOUNT_HT_IN   = 46      # inches AFF


def _small_closed_polylines(block, world):
    """Return list of small closed-polyline candidates inside this block, with local offsets."""
    bx = world.BasisX; by = world.BasisY; origin = world.Origin
    out = []
    def walk(g, pworld):
        if isinstance(g, DB.GeometryInstance):
            child_world = pworld.Multiply(g.Transform)
            for c in g.GetSymbolGeometry():
                walk(c, child_world)
            return
        if isinstance(g, DB.PolyLine):
            pts = list(g.GetCoordinates())
            if 4 <= len(pts) <= 5:
                wpts = [(pworld.OfPoint(p).X, pworld.OfPoint(p).Y) for p in pts]
                xs = [p[0] for p in wpts]; ys = [p[1] for p in wpts]
                w = max(xs)-min(xs); h = max(ys)-min(ys)
                if SMALL_SHAPE_MIN < max(w, h) < SMALL_SHAPE_MAX:
                    cx = (min(xs)+max(xs))/2.0; cy = (min(ys)+max(ys))/2.0
                    dx = cx - origin.X; dy = cy - origin.Y
                    lx = dx*bx.X + dy*bx.Y
                    ly = dx*by.X + dy*by.Y
                    out.append({'world_xy': (cx, cy), 'size': (w, h), 'local_offset': (lx, ly)})
    for c in block.GetSymbolGeometry():
        walk(c, world)
    return out


def _block_bbox(block, world):
    """Bbox of the block's geometry in WORLD XY coords. Returns (min_x, min_y, max_x, max_y) or None."""
    bb = [float('inf'), float('inf'), float('-inf'), float('-inf')]
    def walk(g, pworld):
        if isinstance(g, DB.GeometryInstance):
            child_world = pworld.Multiply(g.Transform)
            for c in g.GetSymbolGeometry():
                walk(c, child_world)
            return
        pts = None
        try:
            if isinstance(g, DB.Line):
                pts = [g.GetEndPoint(0), g.GetEndPoint(1)]
            elif isinstance(g, DB.PolyLine):
                pts = list(g.GetCoordinates())
            elif isinstance(g, DB.Arc):
                if g.IsBound:
                    pts = [g.GetEndPoint(0), g.GetEndPoint(1), g.Center]
                else:
                    cen = g.Center; r = g.Radius
                    pts = [cen,
                           DB.XYZ(cen.X+r, cen.Y, cen.Z),
                           DB.XYZ(cen.X-r, cen.Y, cen.Z),
                           DB.XYZ(cen.X, cen.Y+r, cen.Z),
                           DB.XYZ(cen.X, cen.Y-r, cen.Z)]
        except: pts = None
        if pts:
            for p in pts:
                try:
                    wp = pworld.OfPoint(p)
                    if wp.X < bb[0]: bb[0] = wp.X
                    if wp.Y < bb[1]: bb[1] = wp.Y
                    if wp.X > bb[2]: bb[2] = wp.X
                    if wp.Y > bb[3]: bb[3] = wp.Y
                except: pass
    for c in block.GetSymbolGeometry():
        walk(c, world)
    if bb[0] == float('inf'): return None
    return tuple(bb)


def _is_powered(candidates):
    """True if at least one center-square candidate sits on the lateral centerline at the back of the block,
    OR a symmetric pair exists around the centerline."""
    centered = [c for c in candidates
                if abs(c['local_offset'][1]) < POWER_BAR_LOCAL_Y_TOL
                and POWER_BAR_LOCAL_X_MIN < c['local_offset'][0] < POWER_BAR_LOCAL_X_MAX]
    if centered:
        return True
    by_lx = {}
    for c in candidates:
        lx, ly = c['local_offset']
        if POWER_BAR_LOCAL_X_MIN < lx < POWER_BAR_LOCAL_X_MAX and 0.3 < abs(ly) < 2.0:
            key = round(lx, 1)
            by_lx.setdefault(key, []).append(c)
    for lx, lst in by_lx.items():
        ly_vals = [c['local_offset'][1] for c in lst]
        if len(lst) >= 2 and any(v > 0 for v in ly_vals) and any(v < 0 for v in ly_vals):
            return True
    return False


def find_powered_gym_blocks(doc, view, layer='A-N-GYM EQUIPMENT', building_bounds=None):
    """Scan the linked CAD in `view` for blocks on `layer` and return placement records:
        [{'block_xy': (x, y), 'north_center': (x, y), 'bb': (mnx, mny, mxx, mxy), 'world_transform': Transform}, ...]
    Each record corresponds to one POWERED block. The 'north_center' is the recommended
    receptacle placement point (north edge center of the block's bbox in world coords).
    """
    cad = None
    for inst in DB.FilteredElementCollector(doc, view.Id).OfClass(DB.ImportInstance):
        cad = inst; break
    if cad is None:
        return []
    opts = DB.Options(); opts.View = view
    geom = cad.get_Geometry(opts)
    out = []
    def walk(g, parent_world, depth=0):
        if not isinstance(g, DB.GeometryInstance):
            return
        own = g.Transform
        world = parent_world.Multiply(own) if parent_world else own
        wo = world.Origin
        try:
            gs = doc.GetElement(g.GraphicsStyleId)
            ln = gs.GraphicsStyleCategory.Name if gs else None
        except: ln = None
        if ln == layer:
            ok = True
            if building_bounds is not None:
                bnd = building_bounds
                ok = (bnd[0] <= wo.X <= bnd[2] and bnd[1] <= wo.Y <= bnd[3])
            if ok:
                candidates = _small_closed_polylines(g, world)
                if _is_powered(candidates):
                    bb = _block_bbox(g, world)
                    if bb is not None:
                        cx = (bb[0]+bb[2])/2.0
                        out.append({
                            'block_xy': (wo.X, wo.Y),
                            'north_center': (cx, bb[3]),
                            'bb': bb,
                            'world_transform': world,
                        })
        if depth < 3:
            for c in g.GetSymbolGeometry():
                walk(c, world, depth+1)
    for g in geom:
        walk(g, None, 0)
    return out
