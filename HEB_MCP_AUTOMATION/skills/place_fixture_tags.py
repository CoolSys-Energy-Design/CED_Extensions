# SKILL: place_fixture_tags  (runs AFTER wires)
# Recreates each Electrical Fixture Tag as an IndependentTag referencing the
# PLACED target device (mapped from source via place_relative_map.devices).
# Head = tgt_device_pos + Rz(dtheta)*offset_from_host.  Leaders rebuilt.
# Globals: BA_FT_APPLY (default False), BA_FT_REBUILD, BA_TGT_TITLE
# Writes: data\place_fixture_tags_report.json ; updates place_relative_map.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math

APPLY   = globals().get("BA_FT_APPLY", False)
REBUILD = globals().get("BA_FT_REBUILD", False)
TGT     = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
tv  = get_view(tgt)

ftags = json.load(io.open(DATA + r"\fixture_tags.json"))["elements"]
hosts = {r["id"]: r for r in json.load(io.open(DATA + r"\host_elements.json"))["elements"]}
MAP = DATA + r"\place_relative_map.json"
pm  = json.load(io.open(MAP)); pm.setdefault("fixture_tags", {})
dev_map = pm["devices"]

# target tag types (category Electrical Fixture Tags) by family name
tagtypes = {}; any_tt = None
for et in DB.FilteredElementCollector(tgt).OfClass(DB.FamilySymbol):
    try:
        if et.Category and et.Category.Name == "Electrical Fixture Tags":
            tagtypes[nm(et.Family)] = et.Id
            if any_tt is None: any_tt = et.Id
    except: pass

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
tg = DB.TransactionGroup(tgt, "BAKERY fixture tags")
if APPLY: tg.Start()
t = DB.Transaction(tgt, "place fixture tags")
if APPLY: t.Start()
try:
    if APPLY and REBUILD:
        for k, tid in list(pm["fixture_tags"].items()):
            el = tgt.GetElement(DB.ElementId(int(tid)))
            if el is not None:
                try: tgt.Delete(el.Id); removed += 1
                except: pass
        pm["fixture_tags"] = {}
    for ft in ftags:
        fid = str(ft["id"])
        if fid in pm["fixture_tags"]:
            skipped += 1; continue
        sdev = str(ft.get("tagged_id"))
        if sdev not in dev_map:
            rows.append({"id": ft["id"], "status": "device-not-placed"}); skipped += 1; continue
        de = tgt.GetElement(DB.ElementId(int(dev_map[sdev])))
        sr = hosts.get(int(sdev))
        if de is None or sr is None:
            rows.append({"id": ft["id"], "status": "tgt-device-missing"}); skipped += 1; continue
        P = de.Location.Point
        fo = getattr(de, 'FacingOrientation', None)
        tang = math.atan2(fo.Y, fo.X) if fo else 0.0
        sf = sr.get("facing") or [1.0, 0.0]
        dth = tang - math.atan2(sf[1], sf[0])
        off = ft.get("offset_from_host") or [0.0, 0.0, 0.0]
        c, s = math.cos(dth), math.sin(dth)
        hx = P.X + c*off[0] - s*off[1]
        hy = P.Y + s*off[0] + c*off[1]
        hz = P.Z + off[2]
        rec = {"id": ft["id"], "tagged_dev": int(sdev),
               "text": ft.get("tag_text"), "status": "ok"}
        if APPLY:
            ttid = tagtypes.get(ft.get("tag_family")) or any_tt
            if ttid is None:
                rec["status"] = "no-tag-type"; rows.append(rec); skipped += 1; continue
            try:
                it = tag_create(ttid, DB.Reference(de),
                                bool(ft.get("has_leader")), DB.XYZ(hx, hy, hz))
            except Exception as e:
                rec["status"] = "create-fail:%s" % e; rows.append(rec); skipped += 1; continue
            try: it.TagHeadPosition = DB.XYZ(hx, hy, hz)
            except: pass
            if ft.get("has_leader") and ft.get("leader_end"):
                le = ft["leader_end"]
                lx = P.X + c*(le[0]-sr["world_xyz"][0]) - s*(le[1]-sr["world_xyz"][1])
                ly = P.Y + s*(le[0]-sr["world_xyz"][0]) + c*(le[1]-sr["world_xyz"][1])
                try:
                    it.LeaderEndCondition = DB.LeaderEndCondition.Free
                    it.SetLeaderEnd(DB.Reference(de), DB.XYZ(lx, ly, P.Z))
                except: pass
            pm["fixture_tags"][fid] = it.Id.IntegerValue
            placed += 1
        rows.append(rec)
    if APPLY:
        t.Commit(); tg.Assimilate()
except Exception as ex:
    if APPLY:
        t.RollBack()
        if tg.GetStatus() == DB.TransactionStatus.Started: tg.RollBack()
    print("FIXTURE TAGS ABORTED, rolled back:", ex); raise

if APPLY:
    with io.open(MAP, 'w', encoding='utf-8') as f:
        f.write(json.dumps(pm, indent=1))
write_json("place_fixture_tags_report.json",
           {"target": tgt.Title, "apply": APPLY, "count": len(ftags),
            "placed": placed, "skipped": skipped, "removed": removed, "rows": rows})
print("=== FIXTURE TAGS %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
print("tags:%d placed=%d skipped=%d removed=%d  tag_types=%s"
      % (len(ftags), placed, skipped, removed, list(tagtypes.keys())))
for r in rows:
    if r["status"] != "ok": print("  !", r["id"], r["status"])
print("written: data\\place_fixture_tags_report.json")
