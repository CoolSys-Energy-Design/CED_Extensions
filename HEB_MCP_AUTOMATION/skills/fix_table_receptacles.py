# SKILL: fix_table_receptacles  (corrective)
# For the equip-snapped receptacles whose facing was wrongly re-aimed:
#  - ensure device XY is within its nearest linked TABLE bbox (clamp w/ inset)
#  - set facing to OUTWARD-FROM-NEAREST-WALL (from wall toward table center),
#    matching the correct devices.
# Globals: BA_FIX_APPLY (default False), BA_TGT_TITLE, BA_LINK,
#          BA_FIX_IDS (list[int]; default = snap_report 'equip' devices)
# Writes: data\fix_table_report.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math, re

APPLY   = globals().get("BA_FIX_APPLY", False)
TGT     = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
LINKSUB = globals().get("BA_LINK", "Oakmont_v24_HEB_ARCH")
INSET   = 0.3   # ft kept inside table bbox

app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
bv  = get_view(tgt)
li  = next(x for x in DB.FilteredElementCollector(tgt, bv.Id)
           if isinstance(x, DB.RevitLinkInstance) and x.GetLinkDocument()
           and LINKSUB in x.GetLinkDocument().Title)
ld  = li.GetLinkDocument(); tf = li.GetTotalTransform()
pm  = json.load(io.open(DATA + r"\place_relative_map.json"))["devices"]
# original equipment-relative (pre-snap) positions = design spacing along table
ORIG = {x["id"]: x.get("tgt_dev_xyz")
        for x in json.load(io.open(DATA + r"\place_relative_report.json"))["rows"]}

IDS = globals().get("BA_FIX_IDS",
        [m["id"] for m in json.load(io.open(DATA + r"\snap_report.json"))["moves"]
         if m.get("kind") == "equip"])

inv = bv.CropBox.Transform.Inverse; cb = bv.CropBox
def near(p, pad=10):
    q = inv.OfPoint(p); return cb.Min.X-pad<=q.X<=cb.Max.X+pad and cb.Min.Y-pad<=q.Y<=cb.Max.Y+pad

TBL = re.compile(r"table|worktop|worktable|smartlever|workstation", re.I)
tables = []
for e in DB.FilteredElementCollector(ld).OfCategory(
        DB.BuiltInCategory.OST_SpecialityEquipment).WhereElementIsNotElementType():
    fs = ld.GetElement(e.GetTypeId())
    fam = nm(fs.Family) if fs and hasattr(fs, 'Family') else ""
    if not fam or not TBL.search(fam): continue
    bb = e.get_BoundingBox(None)
    if not bb: continue
    mn = tf.OfPoint(bb.Min); mx = tf.OfPoint(bb.Max)
    cx, cy = (mn.X+mx.X)/2, (mn.Y+mx.Y)/2
    if not near(DB.XYZ(cx, cy, mn.Z)): continue
    tables.append((fam, min(mn.X,mx.X), min(mn.Y,mx.Y), max(mn.X,mx.X), max(mn.Y,mx.Y)))

walls = []
for w in DB.FilteredElementCollector(ld).OfClass(DB.Wall):
    try:
        c = w.Location.Curve
        a = tf.OfPoint(c.GetEndPoint(0)); b = tf.OfPoint(c.GetEndPoint(1))
        if near(a) or near(b): walls.append((a, b))
    except: pass

def nearest_table(p):
    best = None
    for (fam, x0, y0, x1, y1) in tables:
        cx, cy = (x0+x1)/2, (y0+y1)/2
        d = math.hypot(p.X-cx, p.Y-cy)
        if best is None or d < best[0]: best = (d, fam, x0, y0, x1, y1)
    return best

def outward_normal(p, tcx, tcy):
    """unit normal of nearest wall, signed to point from wall toward table center."""
    best = None
    for (a, b) in walls:
        abx, aby = b.X-a.X, b.Y-a.Y
        L2 = abx*abx + aby*aby
        if L2 < 1e-9: continue
        t = max(0.0, min(1.0, ((p.X-a.X)*abx + (p.Y-a.Y)*aby)/L2))
        qx, qy = a.X+t*abx, a.Y+t*aby
        d = math.hypot(p.X-qx, p.Y-qy)
        if best is None or d < best[0]:
            nx, ny = -aby, abx
            nl = math.hypot(nx, ny) or 1.0
            nx, ny = nx/nl, ny/nl
            if (tcx-qx)*nx + (tcy-qy)*ny < 0: nx, ny = -nx, -ny
            best = (d, nx, ny)
    return (best[1], best[2]) if best else None

moves = []
t = None
if APPLY:
    t = DB.Transaction(tgt, "BAKERY fix table receptacles"); t.Start()
try:
    for did in IDS:
        e = tgt.GetElement(DB.ElementId(int(pm[str(did)])))
        if e is None:
            moves.append({"id": did, "status": "gone"}); continue
        p = e.Location.Point
        nt = nearest_table(p)
        if not nt:
            moves.append({"id": did, "status": "no-table"}); continue
        _, fam, x0, y0, x1, y1 = nt
        tcx, tcy = (x0+x1)/2, (y0+y1)/2
        # facing is already correct (user-confirmed). Seat the device against
        # the table's WALL-SIDE edge (opposite the facing direction), and keep
        # the source design spacing on the along-table axis.
        fo = e.FacingOrientation
        fx = 1.0 if fo.X > 0.5 else (-1.0 if fo.X < -0.5 else 0.0)
        fy = 1.0 if fo.Y > 0.5 else (-1.0 if fo.Y < -0.5 else 0.0)
        orig = ORIG.get(did) or [p.X, p.Y]
        if fy != 0.0:                       # faces +/-Y -> depth on Y, along X
            ny = (y1 - INSET) if fy < 0 else (y0 + INSET)   # wall-side Y edge
            nx = min(max(orig[0], x0+INSET), x1-INSET)       # design spacing on X
        elif fx != 0.0:                     # faces +/-X -> depth on X, along Y
            nx = (x1 - INSET) if fx < 0 else (x0 + INSET)
            ny = min(max(orig[1], y0+INSET), y1-INSET)
        else:                               # unknown facing -> clamp
            nx = min(max(p.X, x0+INSET), x1-INSET)
            ny = min(max(p.Y, y0+INSET), y1-INSET)
        rec = {"id": did, "table": fam[:34],
               "facing": [round(fo.X,2), round(fo.Y,2)],
               "pos_from": [round(p.X,3), round(p.Y,3)],
               "pos_to": [round(nx,3), round(ny,3)],
               "move_ft": round(math.hypot(nx-p.X, ny-p.Y), 3)}
        if APPLY and rec["move_ft"] > 1e-4:
            DB.ElementTransformUtils.MoveElement(tgt, e.Id,
                DB.XYZ(nx-p.X, ny-p.Y, 0.0))
        moves.append(rec)
    if APPLY: t.Commit()
except Exception as ex:
    if APPLY and t: t.RollBack()
    print("FIX ABORTED, rolled back:", ex); raise

write_json("fix_table_report.json", {"target": tgt.Title, "apply": APPLY,
           "tables_considered": len(tables), "walls": len(walls), "moves": moves})
print("=== FIX TABLE RECEPTACLES %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
for m in moves:
    print("  dev %s table=%s facing=%s  %s -> %s  move=%s ft" % (
        m["id"], m.get("table"), m.get("facing"),
        m.get("pos_from"), m.get("pos_to"), m.get("move_ft")))
print("written: data\\fix_table_report.json")
