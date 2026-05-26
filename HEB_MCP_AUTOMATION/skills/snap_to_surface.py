# SKILL: snap_to_surface  (post-placement correction)
# Seats placed receptacles flush against the nearest target SURFACE
# (wall face or equipment bounding-box face) from the Oakmont equipment/arch
# link, fixing the "juts out" defect from long-reach equipment anchors.
# Keeps Z; re-aims device facing to the surface normal (away from surface).
#
# Globals:
#   BA_SNAP_APPLY   (default False -> dry-run report only)
#   BA_TGT_TITLE    (default "OkmntProfCorr")
#   BA_LINK         link doc substring holding walls/equip (default "Oakmont_v24_HEB_ARCH")
#   BA_SNAP_ADIST   only snap devices whose anchor_dist_ft > this (default 3.0)
#   BA_SNAP_MAX     skip if nearest surface farther than this ft (default 9.0)
# Writes: data\snap_report.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math

APPLY    = globals().get("BA_SNAP_APPLY", False)
TGT      = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
LINKSUB  = globals().get("BA_LINK", "Oakmont_v24_HEB_ARCH")
ADIST    = globals().get("BA_SNAP_ADIST", 3.0)
SNAP_MAX = globals().get("BA_SNAP_MAX", 9.0)

app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
bv  = get_view(tgt)
li  = next(x for x in DB.FilteredElementCollector(tgt, bv.Id)
           if isinstance(x, DB.RevitLinkInstance) and x.GetLinkDocument()
           and LINKSUB in x.GetLinkDocument().Title)
ld  = li.GetLinkDocument(); tf = li.GetTotalTransform()

inv = bv.CropBox.Transform.Inverse; cb = bv.CropBox
def near_bakery(p, pad=10):
    q = inv.OfPoint(p)
    return cb.Min.X-pad <= q.X <= cb.Max.X+pad and cb.Min.Y-pad <= q.Y <= cb.Max.Y+pad

# --- candidate surfaces near the bakery ---
walls = []   # (a, b, halfwidth)  world-space centerline segment + half thickness
for w in DB.FilteredElementCollector(ld).OfClass(DB.Wall):
    try:
        c = w.Location.Curve
        a = tf.OfPoint(c.GetEndPoint(0)); b = tf.OfPoint(c.GetEndPoint(1))
        if not (near_bakery(a) or near_bakery(b)): continue
        walls.append((a, b, (w.Width or 0.5)/2.0))
    except: pass

boxes = []   # (minX,minY,maxX,maxY) world axis-aligned (link is unrotated)
for bic in (DB.BuiltInCategory.OST_SpecialityEquipment,
            DB.BuiltInCategory.OST_MechanicalEquipment):
    for e in DB.FilteredElementCollector(ld).OfCategory(bic).WhereElementIsNotElementType():
        try:
            bb = e.get_BoundingBox(None)
            if bb is None: continue
            mn = tf.OfPoint(bb.Min); mx = tf.OfPoint(bb.Max)
            cx = (mn.X+mx.X)/2.0; cy = (mn.Y+mx.Y)/2.0
            if not near_bakery(DB.XYZ(cx, cy, mn.Z)): continue
            boxes.append((min(mn.X,mx.X), min(mn.Y,mx.Y),
                          max(mn.X,mx.X), max(mn.Y,mx.Y)))
        except: pass

def proj_seg(p, a, b):
    abx, aby = b.X-a.X, b.Y-a.Y
    L2 = abx*abx + aby*aby
    if L2 < 1e-9: return a.X, a.Y, 0.0
    t = ((p.X-a.X)*abx + (p.Y-a.Y)*aby)/L2
    t = max(0.0, min(1.0, t))
    return a.X+t*abx, a.Y+t*aby, t

def nearest_surface(p):
    best = None  # (dist, snapX, snapY, kind, nx, ny)
    for (a, b, hw) in walls:
        cxp, cyp, _ = proj_seg(p, a, b)
        vx, vy = p.X-cxp, p.Y-cyp
        d = math.hypot(vx, vy)
        if d < 1e-6:  # on centerline; push out along wall normal
            dx, dy = b.X-a.X, b.Y-a.Y
            nx, ny = -dy, dx
            nlen = math.hypot(nx, ny) or 1.0
            nx, ny = nx/nlen, ny/nlen
        else:
            nx, ny = vx/d, vy/d
        face_d = abs(d - hw)               # distance from p to the wall face
        sx, sy = cxp + nx*hw, cyp + ny*hw  # point on the room-side face
        if best is None or face_d < best[0]:
            best = (face_d, sx, sy, "wall", nx, ny)
    for (x0, y0, x1, y1) in boxes:
        inside = (x0 <= p.X <= x1) and (y0 <= p.Y <= y1)
        # distance to each of the 4 vertical faces
        cands = [(abs(p.X-x0), x0, max(y0,min(p.Y,y1)), -1, 0),
                 (abs(p.X-x1), x1, max(y0,min(p.Y,y1)),  1, 0),
                 (abs(p.Y-y0), max(x0,min(p.X,x1)), y0, 0, -1),
                 (abs(p.Y-y1), max(x0,min(p.X,x1)), y1, 0,  1)]
        fd, fx, fy, nx, ny = min(cands, key=lambda z: z[0])
        if not inside:
            fd = math.hypot(p.X-fx, p.Y-fy)
        if best is None or fd < best[0]:
            best = (fd, fx, fy, "equip", float(nx), float(ny))
    return best

pm = json.load(io.open(DATA + r"\place_relative_map.json"))["devices"]
rep = {x["id"]: x for x in json.load(io.open(DATA + r"\place_relative_report.json"))["rows"]}

targets = []
for sid, tid in pm.items():
    info = rep.get(int(sid), {})
    fam = info.get("fam", "")
    ad  = info.get("anchor_dist_ft", 0)
    if "Receptacle" not in (fam or ""):       # receptacles only
        continue
    if ad <= ADIST:                            # only long-reach (unreliable) ones
        continue
    targets.append((sid, int(tid)))

moves = []
t = None
if APPLY:
    t = DB.Transaction(tgt, "BAKERY snap receptacles to surface"); t.Start()
applied = skipped = 0
try:
    for sid, tid in targets:
        e = tgt.GetElement(DB.ElementId(tid))
        if e is None: skipped += 1; continue
        p = e.Location.Point
        ns = nearest_surface(p)
        if ns is None or ns[0] > SNAP_MAX:
            skipped += 1
            moves.append({"id": int(sid), "status": "no-surface-within-max",
                          "dist": None}); continue
        fd, sx, sy, kind, nx, ny = ns
        rec = {"id": int(sid), "kind": kind, "move_ft": round(fd, 3),
               "from": [round(p.X,3), round(p.Y,3)],
               "to": [round(sx,3), round(sy,3)]}
        if APPLY:
            DB.ElementTransformUtils.MoveElement(
                tgt, e.Id, DB.XYZ(sx-p.X, sy-p.Y, 0.0))
            # re-aim: face away from the surface (normal nx,ny)
            try:
                fo = e.FacingOrientation
                cur = math.atan2(fo.Y, fo.X)
                des = math.atan2(ny, nx)
                dth = des - cur
                if abs(((dth+math.pi) % (2*math.pi)) - math.pi) > 0.05:
                    q = e.Location.Point
                    ax = DB.Line.CreateBound(q, DB.XYZ(q.X, q.Y, q.Z+1))
                    DB.ElementTransformUtils.RotateElement(tgt, e.Id, ax, dth)
                    rec["reaimed_deg"] = round(math.degrees(dth), 1)
            except: pass
            applied += 1
        moves.append(rec)
    if APPLY: t.Commit()
except Exception as ex:
    if APPLY and t: t.RollBack()
    print("SNAP ABORTED, rolled back:", ex); raise

write_json("snap_report.json", {"target": tgt.Title, "link": ld.Title,
           "apply": APPLY, "candidates_walls": len(walls),
           "candidates_equip_boxes": len(boxes),
           "device_count": len(targets), "moves": moves})
print("=== SNAP %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
print("wall segs:%d equip boxes:%d | receptacles targeted (anchor_dist>%.1f): %d"
      % (len(walls), len(boxes), ADIST, len(targets)))
for m in moves:
    print("  dev %s -> %s move=%s ft %s" % (
        m["id"], m.get("kind","?"), m.get("move_ft"),
        ("reaim %s deg" % m.get("reaimed_deg")) if "reaimed_deg" in m else ""))
if APPLY: print("applied=%d skipped=%d" % (applied, skipped))
print("written: data\\snap_report.json")
