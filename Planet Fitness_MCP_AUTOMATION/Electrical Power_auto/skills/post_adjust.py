# post_adjust.py
# After initial profile-driven placement, adjust fixtures to comply with learned rules:
# - Snap to nearest LONG wall in the correct direction (per ROOM_DIRECTION dict)
# - Apply same-side rule (don't put on wall opposite to equipment)
# - Apply min wall-to-equipment clearance
# - Apply flush offset (0.1 ft inward from wall face)
# - Use corrected rotation formula (face into room)
# Generic — works for any project; takes room→direction map as input.

import math
import clr
from Autodesk.Revit import DB

_NAME = clr.GetClrType(DB.Element).GetProperty('Name')

# Default direction-per-room-pattern. Generic; override per project if needed.
DEFAULT_ROOM_DIRECTION = {
    'RED WAVE':   'south',
    'CRYO':       'east',     # cryolounge beds on east wall (per Charlotte feedback)
    'HYDRO':      'west',     # hydromassage beds on west wall
    'SAUNA':      'east',
    'SPRAY':      'west',
    'IT ROOM':    'east',
    'RECEPTION':  'east',
    'VESTIBULE':  'south',
    'BLACK CARD': 'north',
    'STRENGTH':   'north',
    'FREE WEIGHTS': 'north',
    'STORAGE':    'south',
    'TOILET':     'east',
}


def get_wall_segs(doc, view):
    cad = None
    for inst in DB.FilteredElementCollector(doc, view.Id).OfClass(DB.ImportInstance):
        cad = inst; break
    if cad is None: return []
    opts = DB.Options(); opts.View = view
    WALL_LAYERS = set([sc.Name for sc in cad.Category.SubCategories if 'WALL' in sc.Name.upper()])
    geom = cad.get_Geometry(opts)
    segs = []
    def layer_of(g):
        try:
            gs = doc.GetElement(g.GraphicsStyleId)
            return gs.GraphicsStyleCategory.Name if gs else None
        except: return None
    def walk(g, pworld):
        if isinstance(g, DB.GeometryInstance):
            cw = pworld.Multiply(g.Transform) if pworld else g.Transform
            for c in g.GetSymbolGeometry(): walk(c, cw)
            return
        ln = layer_of(g)
        if ln not in WALL_LAYERS: return
        try:
            if isinstance(g, DB.Line):
                p0 = pworld.OfPoint(g.GetEndPoint(0)); p1 = pworld.OfPoint(g.GetEndPoint(1))
                segs.append({'p0':(p0.X,p0.Y),'p1':(p1.X,p1.Y),'len':math.hypot(p1.X-p0.X,p1.Y-p0.Y)})
            elif isinstance(g, DB.PolyLine):
                pts = list(g.GetCoordinates())
                for i in range(len(pts)-1):
                    p0 = pworld.OfPoint(pts[i]); p1 = pworld.OfPoint(pts[i+1])
                    segs.append({'p0':(p0.X,p0.Y),'p1':(p1.X,p1.Y),'len':math.hypot(p1.X-p0.X,p1.Y-p0.Y)})
        except: pass
    for g in geom: walk(g, None)
    return segs


def make_walls(segs, min_len):
    out = []
    for w in segs:
        if w['len'] <= min_len: continue
        wdx, wdy = w['p1'][0]-w['p0'][0], w['p1'][1]-w['p0'][1]; wlen = w['len']
        orient = 'horizontal' if abs(wdx)/wlen > 0.95 else ('vertical' if abs(wdy)/wlen > 0.95 else 'other')
        out.append({'p0':w['p0'],'p1':w['p1'],'len':wlen,'orient':orient})
    return out


def dist_pt_seg(px, py, x1, y1, x2, y2):
    dx, dy = x2-x1, y2-y1
    if dx==0 and dy==0: return math.hypot(px-x1, py-y1), x1, y1
    t = max(0, min(1, ((px-x1)*dx + (py-y1)*dy)/(dx*dx+dy*dy)))
    cx, cy = x1+t*dx, y1+t*dy
    return math.hypot(px-cx, py-cy), cx, cy


def find_best_wall(room_xy, direction, walls, ha=10, hp=10):
    """Find best wall in direction. Returns (proj_x, proj_y, facing_rad, inward) or None."""
    if direction == 'east':   req='vertical';   in_dir = lambda dx,dy: dx>0.3 and abs(dy)<hp and dx<ha
    elif direction == 'west': req='vertical';   in_dir = lambda dx,dy: dx<-0.3 and abs(dy)<hp and -dx<ha
    elif direction == 'north':req='horizontal'; in_dir = lambda dx,dy: dy>0.3 and abs(dx)<hp and dy<ha
    else:                     req='horizontal'; in_dir = lambda dx,dy: dy<-0.3 and abs(dx)<hp and -dy<ha
    best = None; best_score = None
    for w in walls:
        if w['orient'] != req: continue
        d_room, proj_x, proj_y = dist_pt_seg(room_xy[0], room_xy[1], w['p0'][0], w['p0'][1], w['p1'][0], w['p1'][1])
        if not in_dir(proj_x-room_xy[0], proj_y-room_xy[1]): continue
        if abs(proj_x-room_xy[0]) > ha or abs(proj_y-room_xy[1]) > ha: continue
        score = d_room - 0.4*w['len']
        if best_score is None or score < best_score:
            wdx,wdy=w['p1'][0]-w['p0'][0],w['p1'][1]-w['p0'][1]; wlen=w['len']; wdx/=wlen; wdy/=wlen
            perp1=(-wdy,wdx); perp2=(wdy,-wdx)
            tor=(room_xy[0]-proj_x, room_xy[1]-proj_y)
            inward = perp1 if perp1[0]*tor[0]+perp1[1]*tor[1]>0 else perp2
            facing_rad = math.atan2(inward[1], inward[0]) - math.pi/2
            best = {'proj':(proj_x,proj_y),'facing_rad':facing_rad,'inward':inward}
            best_score = score
    return best


def get_rooms_map(doc, view):
    rooms = []
    for sp in DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType():
        try:
            nm = sp.LookupParameter('Name'); name = nm.AsString() if nm else ''
            loc = sp.Location
            if not isinstance(loc, DB.LocationPoint): continue
            rooms.append({'name': name or '', 'xy': (loc.Point.X, loc.Point.Y)})
        except: pass
    return rooms


def find_room_at(rooms, x, y, max_dist=15):
    """Return room whose center is nearest to (x,y) within max_dist, or None."""
    best = None; best_d = None
    for r in rooms:
        d = math.hypot(r['xy'][0]-x, r['xy'][1]-y)
        if d > max_dist: continue
        if best_d is None or d < best_d:
            best = r; best_d = d
    return best


def snap_fixture_to_wall(doc, view, elem, direction, walls, ha=10, hp=10,
                          inward_offset=0.1, max_snap_distance=5.0):
    """Snap a fixture to the wall in 'direction' relative to its current position.

    Hard constraint (USER-REQUESTED): if no wall in `direction` is within
    `max_snap_distance` feet of the fixture, DO NOT snap — leave the fixture
    where the profile placed it. This protects rooms whose nearest wall in
    the chosen direction is far away (like BCS Beauty Angel where the BCS
    north wall is ~9 ft from the equipment anchor row).

    Returns True if snapped, False if skipped or no candidate wall found.
    """
    loc = elem.Location
    if not isinstance(loc, DB.LocationPoint): return False
    cur = loc.Point
    target = find_best_wall((cur.X, cur.Y), direction, walls, ha=ha, hp=hp)
    if target is None: return False
    # Distance from fixture to its projected position on the wall
    snap_dist = math.hypot(target['proj'][0] - cur.X, target['proj'][1] - cur.Y)
    if snap_dist > max_snap_distance:
        return False   # too far — leave at profile position
    new_x = target['proj'][0] + inward_offset*target['inward'][0]
    new_y = target['proj'][1] + inward_offset*target['inward'][1]
    if abs(new_x - cur.X) < 0.05 and abs(new_y - cur.Y) < 0.05:
        return False
    DB.ElementTransformUtils.MoveElement(doc, elem.Id, DB.XYZ(new_x - cur.X, new_y - cur.Y, 0))
    cur_rot = loc.Rotation
    rot_delta = target['facing_rad'] - cur_rot
    if abs(rot_delta) > 0.01:
        new_pt = DB.XYZ(new_x, new_y, cur.Z)
        axis = DB.Line.CreateBound(new_pt, DB.XYZ(new_x, new_y, new_pt.Z+1))
        DB.ElementTransformUtils.RotateElement(doc, elem.Id, axis, rot_delta)
    return True
