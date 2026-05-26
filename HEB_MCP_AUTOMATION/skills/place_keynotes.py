# SKILL: place_keynotes  (device-relative; generalizes across projects)
# Each GA_Keynote Symbol_CED is anchored to the placed device it labels:
#   tgt_keynote = tgt_device_pos + Rz(dtheta) * (src_keynote - src_device)
# where dtheta = tgt_device_facing - src_device_facing. Robust because the
# devices were already placed/corrected. View-specific annotation.
# Globals: BA_KN_APPLY (default False -> dry-run), BA_TGT_TITLE
# Writes: data\place_keynotes_report.json ; updates place_relative_map.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math, os

APPLY = globals().get("BA_KN_APPLY", False)
TGT   = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
tv  = get_view(tgt)

kns   = json.load(io.open(DATA + r"\keynotes.json"))["elements"]
hosts = {r["id"]: r for r in json.load(io.open(DATA + r"\host_elements.json"))["elements"]
         if r["category"] == "Electrical Fixtures"}
MAP = DATA + r"\place_relative_map.json"
pm  = json.load(io.open(MAP))
pm.setdefault("keynotes", {})
dev_map = pm["devices"]   # src device id (str) -> tgt element id

# source fixtures that actually got placed (anchors for keynotes)
src_fx = [(sid, hosts[int(sid)]) for sid in dev_map if int(sid) in hosts]

# target family symbols for GA_Keynote Symbol_CED
symv = {}; ftypes = {}
for fs in DB.FilteredElementCollector(tgt).OfClass(DB.FamilySymbol):
    try:
        f = fs.Family.Name
        tn = fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except: continue
    if f == "GA_Keynote Symbol_CED":
        symv[(f, tn)] = fs.Id
        ftypes.setdefault(f, []).append(fs.Id)

def tgt_dev_pose(src_id):
    e = tgt.GetElement(DB.ElementId(int(dev_map[src_id])))
    if e is None: return None
    p = e.Location.Point
    fo = getattr(e, 'FacingOrientation', None)
    a = math.atan2(fo.Y, fo.X) if fo else 0.0
    return p, a

rows = []
placed = skipped = dups = 0
tg = DB.TransactionGroup(tgt, "BAKERY keynotes (device-relative)")
if APPLY: tg.Start()
t = DB.Transaction(tgt, "place keynotes")
if APPLY: t.Start()
try:
    for k in kns:
        kid = str(k["id"])
        if kid in pm["keynotes"]:
            skipped += 1; continue
        kw = k["world_xyz"]
        # nearest placed source fixture to this keynote
        best = None
        for sid, r in src_fx:
            d = (kw[0]-r["world_xyz"][0])**2 + (kw[1]-r["world_xyz"][1])**2
            if best is None or d < best[0]: best = (d, sid, r)
        if not best:
            rows.append({"id": k["id"], "status": "no-anchor-device"}); skipped += 1; continue
        _, sid, sr = best
        pose = tgt_dev_pose(sid)
        if pose is None:
            rows.append({"id": k["id"], "status": "tgt-device-missing"}); skipped += 1; continue
        P, tang = pose
        sf = sr.get("facing") or [1.0, 0.0]
        sang = math.atan2(sf[1], sf[0])
        dth = tang - sang
        ox, oy = kw[0]-sr["world_xyz"][0], kw[1]-sr["world_xyz"][1]
        c, s = math.cos(dth), math.sin(dth)
        nx = P.X + c*ox - s*oy
        ny = P.Y + s*ox + c*oy
        nz = P.Z + (kw[2]-sr["world_xyz"][2])
        fam, typ = k["family"], k["type"]
        rec = {"id": k["id"], "anchor_src_dev": int(sid),
               "kv": k["params"].get("Keynote Value"),
               "pos": [round(nx,3), round(ny,3)], "status": "ok"}
        if APPLY:
            if (fam, typ) not in symv:
                sibs = ftypes.get(fam)
                if not sibs:
                    rec["status"] = "no-family"; rows.append(rec); skipped += 1; continue
                nsym = tgt.GetElement(sibs[0]).Duplicate(typ)
                symv[(fam, typ)] = nsym.Id; ftypes[fam].append(nsym.Id); dups += 1
            fsym = tgt.GetElement(symv[(fam, typ)])
            if not fsym.IsActive: fsym.Activate()
            try:
                fi = tgt.Create.NewFamilyInstance(DB.XYZ(nx, ny, nz), fsym, tv)
            except Exception as e:
                rec["status"] = "place-fail:%s" % e; rows.append(rec); skipped += 1; continue
            krot = (k.get("rotation") or 0.0) + dth
            if abs(krot) > 1e-4:
                try:
                    q = fi.Location.Point
                    ax = DB.Line.CreateBound(q, DB.XYZ(q.X, q.Y, q.Z+1))
                    DB.ElementTransformUtils.RotateElement(tgt, fi.Id, ax, krot)
                except: pass
            for pn in ("Keynote Value", "Symbol Text", "Show Symbol Text"):
                v = k["params"].get(pn)
                if v is None: continue
                try:
                    pr = fi.LookupParameter(pn)
                    if pr and not pr.IsReadOnly:
                        if pr.StorageType == DB.StorageType.Integer: pr.Set(int(v))
                        else: pr.Set(str(v))
                except: pass
            try:
                cp = fi.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                if cp and not cp.IsReadOnly: cp.Set("[BA:knsrc%s]" % kid)
            except: pass
            pm["keynotes"][kid] = fi.Id.IntegerValue
            placed += 1
        rows.append(rec)
    if APPLY:
        t.Commit(); tg.Assimilate()
except Exception as ex:
    if APPLY:
        t.RollBack()
        if tg.GetStatus() == DB.TransactionStatus.Started: tg.RollBack()
    print("KEYNOTES ABORTED, rolled back:", ex); raise

if APPLY:
    with io.open(MAP, 'w', encoding='utf-8') as f:
        f.write(json.dumps(pm, indent=1))
write_json("place_keynotes_report.json",
           {"target": tgt.Title, "apply": APPLY, "count": len(kns),
            "placed": placed, "skipped": skipped, "rows": rows})
print("=== KEYNOTES %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
print("keynotes:%d placed=%d skipped=%d dup_types=%d" % (len(kns), placed, skipped, dups))
bad = [r for r in rows if r["status"] != "ok"]
if bad:
    for r in bad: print("  !", r["id"], r["status"])
print("written: data\\place_keynotes_report.json")
