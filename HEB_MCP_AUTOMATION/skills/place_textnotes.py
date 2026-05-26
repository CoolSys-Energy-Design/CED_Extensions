# SKILL: place_textnotes  (device-relative; runs AFTER keynotes)
# Each source Text Note is anchored to the placed device it sits nearest to:
#   tgt = tgt_device_pos + Rz(dtheta) * (src_textnote_coord - src_device_pos)
# dtheta = tgt_device_facing - src_device_facing.  View-specific.
# Globals: BA_TN_APPLY (default False -> dry-run), BA_TGT_TITLE
# Writes: data\place_textnotes_report.json ; updates place_relative_map.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math, os

APPLY   = globals().get("BA_TN_APPLY", False)
REBUILD = globals().get("BA_TN_REBUILD", False)   # delete+replace existing
TGT     = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
tv  = get_view(tgt)

tns   = json.load(io.open(DATA + r"\textnotes.json"))["elements"]
hosts = {r["id"]: r for r in json.load(io.open(DATA + r"\host_elements.json"))["elements"]
         if r["category"] == "Electrical Fixtures"}
MAP = DATA + r"\place_relative_map.json"
pm  = json.load(io.open(MAP))
pm.setdefault("textnotes", {})
dev_map = pm["devices"]
src_fx = [(sid, hosts[int(sid)]) for sid in dev_map if int(sid) in hosts]

# resolve a target TextNoteType (match source name, else any, else duplicate)
tnt_by_name = {}
default_tnt = None
for tt in DB.FilteredElementCollector(tgt).OfClass(DB.TextNoteType):
    nmt = nm(tt)
    tnt_by_name[nmt] = tt.Id
    if default_tnt is None: default_tnt = tt.Id

def tgt_dev_pose(src_id):
    e = tgt.GetElement(DB.ElementId(int(dev_map[src_id])))
    if e is None: return None
    p = e.Location.Point
    fo = getattr(e, 'FacingOrientation', None)
    return p, (math.atan2(fo.Y, fo.X) if fo else 0.0)

def map_point(px, py, pz):
    """device-relative map of an arbitrary source point via nearest placed
    source fixture (used for leader End/Elbow so they hit the right element)."""
    best = None
    for sid, r in src_fx:
        d = (px-r["world_xyz"][0])**2 + (py-r["world_xyz"][1])**2
        if best is None or d < best[0]: best = (d, sid, r)
    if not best: return None
    _, sid, sr = best
    pose = tgt_dev_pose(sid)
    if pose is None: return None
    P, tang = pose
    sf = sr.get("facing") or [1.0, 0.0]
    dth = tang - math.atan2(sf[1], sf[0])
    ox, oy = px-sr["world_xyz"][0], py-sr["world_xyz"][1]
    c, s = math.cos(dth), math.sin(dth)
    return DB.XYZ(P.X + c*ox - s*oy, P.Y + s*ox + c*oy,
                  P.Z + (pz - sr["world_xyz"][2]))

rows = []
placed = skipped = 0
tg = DB.TransactionGroup(tgt, "BAKERY text notes (device-relative)")
if APPLY: tg.Start()
t = DB.Transaction(tgt, "place text notes")
if APPLY: t.Start()
removed = 0
try:
    if APPLY and REBUILD:
        for nid0, tid0 in list(pm["textnotes"].items()):
            el = tgt.GetElement(DB.ElementId(int(tid0)))
            if el is not None:
                try: tgt.Delete(el.Id); removed += 1
                except: pass
        pm["textnotes"] = {}
    for tn in tns:
        nid = str(tn["id"])
        if nid in pm["textnotes"]:
            skipped += 1; continue
        cw = tn["coord"]
        best = None
        for sid, r in src_fx:
            d = (cw[0]-r["world_xyz"][0])**2 + (cw[1]-r["world_xyz"][1])**2
            if best is None or d < best[0]: best = (d, sid, r)
        if not best:
            rows.append({"id": tn["id"], "status": "no-anchor"}); skipped += 1; continue
        _, sid, sr = best
        pose = tgt_dev_pose(sid)
        if pose is None:
            rows.append({"id": tn["id"], "status": "tgt-device-missing"}); skipped += 1; continue
        P, tang = pose
        sf = sr.get("facing") or [1.0, 0.0]
        dth = tang - math.atan2(sf[1], sf[0])
        ox, oy = cw[0]-sr["world_xyz"][0], cw[1]-sr["world_xyz"][1]
        c, s = math.cos(dth), math.sin(dth)
        nx = P.X + c*ox - s*oy
        ny = P.Y + s*ox + c*oy
        nz = P.Z + (cw[2]-sr["world_xyz"][2])
        txt = tn.get("text") or ""
        rec = {"id": tn["id"], "anchor_src_dev": int(sid),
               "text": txt.replace("\r", " ").strip()[:40],
               "pos": [round(nx,3), round(ny,3)], "status": "ok"}
        if APPLY:
            tnt = tnt_by_name.get(tn.get("textnote_type")) or default_tnt
            try:
                note = DB.TextNote.Create(tgt, tv.Id, DB.XYZ(nx, ny, nz), txt, tnt)
            except Exception as e:
                try:   # fallback: minimal overload
                    note = DB.TextNote.Create(tgt, tv.Id, DB.XYZ(nx, ny, nz),
                                              txt or " ", tnt)
                except Exception as e2:
                    rec["status"] = "create-fail:%s" % e2
                    rows.append(rec); skipped += 1; continue
            # width
            w = tn.get("width")
            if w:
                try: note.Width = float(w)
                except: pass
            # NOTE: text kept HORIZONTAL (no rotation) for readability.
            # Rebuild leaders so they point at the placed target element.
            nleg = 0
            for L in (tn.get("leaders") or []):
                es, eb = L.get("end"), L.get("elbow")
                if not es: continue
                Et = map_point(es[0], es[1], es[2])
                if Et is None: continue
                side = (DB.TextNoteLeaderTypes.TNLT_STRAIGHT_R
                        if Et.X >= nx else DB.TextNoteLeaderTypes.TNLT_STRAIGHT_L)
                try:
                    ld = note.AddLeader(side)
                    ld.End = Et
                    if L.get("has_elbow") and eb:
                        Bt = map_point(eb[0], eb[1], eb[2])
                        if Bt is not None:
                            try: ld.Elbow = Bt
                            except: pass
                    nleg += 1
                except Exception as le:
                    rec["leader_err"] = str(le)
            rec["leaders"] = nleg
            pm["textnotes"][nid] = note.Id.IntegerValue
            placed += 1
        rows.append(rec)
    if APPLY:
        t.Commit(); tg.Assimilate()
except Exception as ex:
    if APPLY:
        t.RollBack()
        if tg.GetStatus() == DB.TransactionStatus.Started: tg.RollBack()
    print("TEXTNOTES ABORTED, rolled back:", ex); raise

if APPLY:
    with io.open(MAP, 'w', encoding='utf-8') as f:
        f.write(json.dumps(pm, indent=1))
write_json("place_textnotes_report.json",
           {"target": tgt.Title, "apply": APPLY, "count": len(tns),
            "placed": placed, "skipped": skipped, "rows": rows})
print("=== TEXT NOTES %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
print("textnotes:%d placed=%d skipped=%d removed(rebuild)=%d leaders=%d"
      % (len(tns), placed, skipped, removed,
         sum(r.get("leaders", 0) for r in rows)))
for r in rows:
    if r["status"] != "ok": print("  !", r["id"], r["status"])
print("samples:", [(r["text"], r["pos"]) for r in rows[:4] if r["status"]=="ok"])
print("written: data\\place_textnotes_report.json")
