# SKILL: place_wires  (device-relative; runs AFTER devices/keynotes/textnotes)
# Each source wire polyline is rigidly re-based on the placed device it
# connects to:  tgtV = tgtDev + Rz(dtheta)*(srcV - srcDev)
# dtheta = tgt_device_facing - src_device_facing.  Anchor device = the wire's
# connected fixture if placed, else nearest placed device to vertex[0].
# Graphical Arc wires (no connectors), matching the source homerun look.
# Globals: BA_W_APPLY (default False), BA_W_REBUILD, BA_TGT_TITLE
# Writes: data\place_wires_report.json ; updates place_relative_map.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
from System.Collections.Generic import List
import math

APPLY   = globals().get("BA_W_APPLY", False)
REBUILD = globals().get("BA_W_REBUILD", False)
ARROW   = globals().get("BA_W_ARROW", True)     # draw homerun arrowhead at panel end
ARROW_L = globals().get("BA_W_ARROW_LEN", 0.55) # ft (model), arrowhead barb length
TGT     = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
tv  = get_view(tgt)

wires = json.load(io.open(DATA + r"\wires.json"))["elements"]
hosts = {r["id"]: r for r in json.load(io.open(DATA + r"\host_elements.json"))["elements"]
         if r["category"] == "Electrical Fixtures"}
MAP = DATA + r"\place_relative_map.json"
pm  = json.load(io.open(MAP)); pm.setdefault("wires", {}); pm.setdefault("wire_arrows", {})
dev_map = pm["devices"]
src_fx = [(sid, hosts[int(sid)]) for sid in dev_map if int(sid) in hosts]

def tgt_dev_pose(src_id):
    e = tgt.GetElement(DB.ElementId(int(dev_map[src_id])))
    if e is None: return None
    p = e.Location.Point
    fo = getattr(e, 'FacingOrientation', None)
    return p, (math.atan2(fo.Y, fo.X) if fo else 0.0)

def anchor_for(w):
    """connected placed fixture, else nearest placed fixture to vertex[0]."""
    for c in w.get("connected_to", []):
        sid = str(c.get("id"))
        if sid in dev_map and int(sid) in hosts:
            return sid
    v0 = w["vertices"][0]
    best = None
    for sid, r in src_fx:
        d = (v0[0]-r["world_xyz"][0])**2 + (v0[1]-r["world_xyz"][1])**2
        if best is None or d < best[0]: best = (d, sid)
    return best[1] if best else None

# wire-type cache (match source name, create if absent)
wt_cache = {}
def wire_type(name):
    if name in wt_cache: return wt_cache[name]
    wid = None
    for wt in DB.FilteredElementCollector(tgt).OfClass(DB.Electrical.WireType):
        if nm(wt) == name: wid = wt.Id; break
    if wid is None:
        try: wid = DB.Electrical.WireType.Create(tgt, name).Id
        except:
            any_wt = DB.FilteredElementCollector(tgt).OfClass(DB.Electrical.WireType).FirstElement()
            wid = any_wt.Id if any_wt else None
    wt_cache[name] = wid
    return wid

# solid filled-region type for arrowheads (prefer a solid/black fill)
_FR_TYPE = None
for _frt in DB.FilteredElementCollector(tgt).OfClass(DB.FilledRegionType):
    if _FR_TYPE is None: _FR_TYPE = _frt.Id
    _n = (nm(_frt) or "").lower()
    if "solid" in _n or "black" in _n or "fill" in _n:
        _FR_TYPE = _frt.Id; break

def make_arrowhead(end_pt, prev_pt):
    """Solid filled-triangle homerun arrowhead at end_pt, opening back along
    (prev_pt -> end_pt). Returns [filledRegionId] (or [] on failure)."""
    dx, dy = end_pt.X - prev_pt.X, end_pt.Y - prev_pt.Y
    L = math.hypot(dx, dy)
    if L < 1e-6 or _FR_TYPE is None: return []
    ux, uy = dx / L, dy / L                  # wire direction (toward tip)
    z = end_pt.Z
    corners = []
    for ang in (math.radians(155.0), math.radians(-155.0)):
        c, s = math.cos(ang), math.sin(ang)
        bx = ux * c - uy * s
        by = ux * s + uy * c
        corners.append(DB.XYZ(end_pt.X + bx * ARROW_L,
                              end_pt.Y + by * ARROW_L, z))
    tip = DB.XYZ(end_pt.X, end_pt.Y, z)
    pts = [tip, corners[0], corners[1]]
    try:
        loop = DB.CurveLoop()
        for i in range(3):
            a = pts[i]; b = pts[(i + 1) % 3]
            if a.DistanceTo(b) < 1e-7: return []
            loop.Append(DB.Line.CreateBound(a, b))
        loops = List[DB.CurveLoop](); loops.Add(loop)
        fr = DB.FilledRegion.Create(tgt, _FR_TYPE, tv.Id, loops)
        return [fr.Id.IntegerValue]
    except Exception:
        return []

rows = []
placed = skipped = removed = arrows = 0
tg = DB.TransactionGroup(tgt, "BAKERY wires (device-relative)")
if APPLY: tg.Start()
t = DB.Transaction(tgt, "place wires")
if APPLY: t.Start()
try:
    if APPLY and REBUILD:
        for wid0, tid0 in list(pm["wires"].items()):
            el = tgt.GetElement(DB.ElementId(int(tid0)))
            if el is not None:
                try: tgt.Delete(el.Id); removed += 1
                except: pass
        for wid0, aids in list(pm["wire_arrows"].items()):
            for aid in aids:
                el = tgt.GetElement(DB.ElementId(int(aid)))
                if el is not None:
                    try: tgt.Delete(el.Id)
                    except: pass
        pm["wires"] = {}; pm["wire_arrows"] = {}
    for w in wires:
        wid = str(w["id"])
        if wid in pm["wires"]:
            skipped += 1; continue
        verts = w.get("vertices") or []
        if len(verts) < 2:
            rows.append({"id": w["id"], "status": "degenerate"}); skipped += 1; continue
        sid = anchor_for(w)
        if sid is None:
            rows.append({"id": w["id"], "status": "no-anchor"}); skipped += 1; continue
        pose = tgt_dev_pose(sid)
        if pose is None:
            rows.append({"id": w["id"], "status": "tgt-device-missing"}); skipped += 1; continue
        P, tang = pose
        sr = hosts[int(sid)]; sw = sr["world_xyz"]
        sf = sr.get("facing") or [1.0, 0.0]
        dth = tang - math.atan2(sf[1], sf[0])
        c, s = math.cos(dth), math.sin(dth)
        tv_pts = []
        for v in verts:
            ox, oy = v[0]-sw[0], v[1]-sw[1]
            tv_pts.append(DB.XYZ(P.X + c*ox - s*oy,
                                 P.Y + s*ox + c*oy,
                                 P.Z + (v[2]-sw[2])))
        # bind the wire start to the device's electrical connector so the
        # wire joins that device's circuit (-> Panel/Circuits populate)
        startc = None
        de = tgt.GetElement(DB.ElementId(int(dev_map[str(sid)])))
        try:
            mm = de.MEPModel
            if mm and mm.ConnectorManager:
                for cc in mm.ConnectorManager.Connectors:
                    try:
                        if cc.Domain == DB.Domain.DomainElectrical:
                            startc = cc; break
                    except: pass
        except: pass
        if startc is not None:
            o = startc.Origin
            shift = DB.XYZ(o.X - tv_pts[0].X, o.Y - tv_pts[0].Y, o.Z - tv_pts[0].Z)
            tv_pts = [DB.XYZ(q.X+shift.X, q.Y+shift.Y, q.Z+shift.Z) for q in tv_pts]
        rec = {"id": w["id"], "anchor_src_dev": int(sid),
               "panel": w.get("circuit", {}).get("panel"),
               "bound": bool(startc), "nverts": len(tv_pts), "status": "ok"}
        if APPLY:
            wtid = wire_type(w.get("wire_type") or "THWN")
            if wtid is None:
                rec["status"] = "no-wire-type"; rows.append(rec); skipped += 1; continue
            vlist = List[DB.XYZ]()
            for p in tv_pts: vlist.Add(p)
            try:
                wn = DB.Electrical.Wire.Create(tgt, wtid, tv.Id,
                        DB.Electrical.WiringType.Arc, vlist, startc, None)
            except Exception as e:
                try:
                    wn = DB.Electrical.Wire.Create(tgt, wtid, tv.Id,
                            DB.Electrical.WiringType.Arc, vlist, None, None)
                    rec["bound"] = False
                except Exception as e2:
                    rec["status"] = "create-fail:%s" % e2
                    rows.append(rec); skipped += 1; continue
            pm["wires"][wid] = wn.Id.IntegerValue
            placed += 1
            if ARROW and len(tv_pts) >= 2:
                # panel/home end = vertex farthest from the anchor device P
                d0 = (tv_pts[0].X-P.X)**2 + (tv_pts[0].Y-P.Y)**2
                dN = (tv_pts[-1].X-P.X)**2 + (tv_pts[-1].Y-P.Y)**2
                if dN >= d0: e_pt, p_pt = tv_pts[-1], tv_pts[-2]
                else:        e_pt, p_pt = tv_pts[0], tv_pts[1]
                aids = make_arrowhead(e_pt, p_pt)
                if aids:
                    pm["wire_arrows"][wid] = aids
                    arrows += 1
        rows.append(rec)
    if APPLY:
        t.Commit(); tg.Assimilate()
except Exception as ex:
    if APPLY:
        t.RollBack()
        if tg.GetStatus() == DB.TransactionStatus.Started: tg.RollBack()
    print("WIRES ABORTED, rolled back:", ex); raise

if APPLY:
    with io.open(MAP, 'w', encoding='utf-8') as f:
        f.write(json.dumps(pm, indent=1))
write_json("place_wires_report.json",
           {"target": tgt.Title, "apply": APPLY, "count": len(wires),
            "placed": placed, "skipped": skipped, "removed": removed, "rows": rows})
print("=== WIRES %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
print("wires:%d placed=%d arrowheads=%d skipped=%d removed=%d"
      % (len(wires), placed, arrows, skipped, removed))
for r in rows:
    if r["status"] != "ok": print("  !", r["id"], r["status"])
print("written: data\\place_wires_report.json")
