# SKILL: place_wire_tags  (runs AFTER wires + fixture tags)
# Recreates each Wire Tag as an IndependentTag referencing the PLACED target
# wire (mapped via place_relative_map.wires). Head = tgt_wire_vertex[idx]
# + Rz(dtheta)*offset_from_wire, where dtheta is the wire's anchor-device
# facing delta (from place_wires_report). Leaders rebuilt.
# Globals: BA_WT_APPLY (default False), BA_WT_REBUILD, BA_TGT_TITLE
# Writes: data\place_wire_tags_report.json ; updates place_relative_map.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math

APPLY   = globals().get("BA_WT_APPLY", False)
REBUILD = globals().get("BA_WT_REBUILD", False)
TGT     = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
tv  = get_view(tgt)

wtags = json.load(io.open(DATA + r"\wire_tags.json"))["elements"]
hosts = {r["id"]: r for r in json.load(io.open(DATA + r"\host_elements.json"))["elements"]}
wrep  = {r["id"]: r for r in json.load(io.open(DATA + r"\place_wires_report.json"))["rows"]}
MAP = DATA + r"\place_relative_map.json"
pm  = json.load(io.open(MAP)); pm.setdefault("wire_tags", {})
dev_map = pm["devices"]; wire_map = pm["wires"]

tt_by_ft = {}; tt_by_fam = {}; any_tt = None   # (fam,type)->id ; fam->id
for et in DB.FilteredElementCollector(tgt).OfClass(DB.FamilySymbol):
    try:
        if et.Category and et.Category.Name == "Wire Tags":
            f = nm(et.Family); ty = nm(et)
            tt_by_ft[(f, ty)] = et.Id
            tt_by_fam[f] = et.Id
            if any_tt is None: any_tt = et.Id
    except: pass
def pick_tagtype(wt):
    return (tt_by_ft.get((wt.get("tag_family"), wt.get("type")))
            or tt_by_fam.get(wt.get("tag_family")) or any_tt)
tagtypes = tt_by_fam   # kept for the summary print

def wire_dtheta(src_wire_id):
    """facing delta of the wire's anchor device (how the wire was rotated)."""
    r = wrep.get(src_wire_id)
    if not r or "anchor_src_dev" not in r: return 0.0
    sd = str(r["anchor_src_dev"])
    if sd not in dev_map: return 0.0
    de = tgt.GetElement(DB.ElementId(int(dev_map[sd])))
    sr = hosts.get(int(sd))
    if de is None or sr is None: return 0.0
    fo = getattr(de, 'FacingOrientation', None)
    tang = math.atan2(fo.Y, fo.X) if fo else 0.0
    sf = sr.get("facing") or [1.0, 0.0]
    return tang - math.atan2(sf[1], sf[0])

def tag_create(typeId, ref, addLeader, pnt):
    try:
        return DB.IndependentTag.Create(tgt, typeId, tv.Id, ref, addLeader,
                 DB.TagOrientation.Horizontal, pnt)
    except:
        it = DB.IndependentTag.Create(tgt, tv.Id, ref, addLeader,
                 DB.TagMode.TM_ADDBY_CATEGORY, DB.TagOrientation.Horizontal, pnt)
        try: it.ChangeTypeId(typeId)
        except: pass
        return it

rows = []
placed = skipped = removed = 0
tg = DB.TransactionGroup(tgt, "BAKERY wire tags")
if APPLY: tg.Start()
t = DB.Transaction(tgt, "place wire tags")
if APPLY: t.Start()
try:
    if APPLY and REBUILD:
        for k, tid in list(pm["wire_tags"].items()):
            el = tgt.GetElement(DB.ElementId(int(tid)))
            if el is not None:
                try: tgt.Delete(el.Id); removed += 1
                except: pass
        pm["wire_tags"] = {}
    for wt in wtags:
        wid = str(wt["id"])
        if wid in pm["wire_tags"]:
            skipped += 1; continue
        swire = str(wt.get("tagged_id"))
        if swire not in wire_map:
            rows.append({"id": wt["id"], "status": "wire-not-placed"}); skipped += 1; continue
        we = tgt.GetElement(DB.ElementId(int(wire_map[swire])))
        if we is None:
            rows.append({"id": wt["id"], "status": "tgt-wire-missing"}); skipped += 1; continue
        try: nv = we.NumberOfVertices
        except: nv = 0
        if nv < 1:
            rows.append({"id": wt["id"], "status": "wire-no-verts"}); skipped += 1; continue
        idx = wt.get("near_vertex_index") or 0
        idx = max(0, min(int(idx), nv - 1))
        vtx = we.GetVertex(idx)
        off = wt.get("offset_from_wire") or [0.0, 0.0, 0.0]
        dth = wire_dtheta(int(swire))
        c, s = math.cos(dth), math.sin(dth)
        hx = vtx.X + c*off[0] - s*off[1]
        hy = vtx.Y + s*off[0] + c*off[1]
        hz = vtx.Z + off[2]
        rec = {"id": wt["id"], "tagged_wire": int(swire),
               "text": wt.get("tag_text"), "status": "ok"}
        if APPLY:
            ttid = pick_tagtype(wt)
            if ttid is None:
                rec["status"] = "no-tag-type"; rows.append(rec); skipped += 1; continue
            try:
                it = tag_create(ttid, DB.Reference(we),
                                bool(wt.get("has_leader")), DB.XYZ(hx, hy, hz))
            except Exception as e:
                rec["status"] = "create-fail:%s" % e; rows.append(rec); skipped += 1; continue
            try: it.TagHeadPosition = DB.XYZ(hx, hy, hz)
            except: pass
            if wt.get("has_leader") and wt.get("leader_end"):
                le = wt["leader_end"]
                lx = vtx.X + c*(le[0]-(wt["head"][0]-off[0])) - s*(le[1]-(wt["head"][1]-off[1]))
                ly = vtx.Y + s*(le[0]-(wt["head"][0]-off[0])) + c*(le[1]-(wt["head"][1]-off[1]))
                try:
                    it.LeaderEndCondition = DB.LeaderEndCondition.Free
                    it.SetLeaderEnd(DB.Reference(we), DB.XYZ(lx, ly, vtx.Z))
                except: pass
            pm["wire_tags"][wid] = it.Id.IntegerValue
            placed += 1
        rows.append(rec)
    if APPLY:
        t.Commit(); tg.Assimilate()
except Exception as ex:
    if APPLY:
        t.RollBack()
        if tg.GetStatus() == DB.TransactionStatus.Started: tg.RollBack()
    print("WIRE TAGS ABORTED, rolled back:", ex); raise

if APPLY:
    with io.open(MAP, 'w', encoding='utf-8') as f:
        f.write(json.dumps(pm, indent=1))
write_json("place_wire_tags_report.json",
           {"target": tgt.Title, "apply": APPLY, "count": len(wtags),
            "placed": placed, "skipped": skipped, "removed": removed, "rows": rows})
print("=== WIRE TAGS %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
print("tags:%d placed=%d skipped=%d removed=%d  tag_types=%s"
      % (len(wtags), placed, skipped, removed, list(tagtypes.keys())))
for r in rows:
    if r["status"] != "ok": print("  !", r["id"], r["status"])
print("written: data\\place_wire_tags_report.json")
