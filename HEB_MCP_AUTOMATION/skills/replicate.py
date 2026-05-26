# SKILL: replicate  (writes into the TARGET project)
# Phased replication of the BAKERY power design.
#   PHASE "A" : place Electrical Fixtures (devices). Panels are NOT duplicated.
#   PHASE "B" : place GA_Keynote Symbol_CED annotations.
#   PHASE "C" : place wires (vertex polylines).
#   PHASE "D" : circuiting (existing panels only).
# Set PHASE below. Source may be the active doc; target is written cross-doc.
# Idempotent: stamps Comments with [BA:src<id>] and skips already-placed.
# Writes/updates: data\replication_map.json , data\replication_report.txt
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
from System.Collections.Generic import List

PHASE = globals().get("BA_PHASE", "A")   # set BA_PHASE before exec to pick phase
TGT_TITLE = "CED HEB Test Run_MEPR_R24"
SRC_EQUIP_LINK = "Equip"            # source equipment link doc title contains
TGT_EQUIP_LINK = "Carrollton_v24"   # target equipment link doc title contains
TGT_LEVEL = "Level 1 - Sheet view"
TAGP = DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
MARKP = DB.BuiltInParameter.ALL_MODEL_MARK

app = doc.Application
tgt = next(d for d in app.Documents if d.Title == TGT_TITLE)
src = doc if doc.Title.startswith("RunUpdateProfiles") else \
      next(d for d in app.Documents if d.Title.startswith("RunUpdateProfiles"))

def equip_tf(d, part):
    v = get_view(d)
    for li in DB.FilteredElementCollector(d, v.Id):
        if isinstance(li, DB.RevitLinkInstance):
            ld = li.GetLinkDocument()
            if ld and part in ld.Title:
                return li.GetTotalTransform()
    raise Exception("equip link not found in %s" % d.Title)

T = equip_tf(tgt, TGT_EQUIP_LINK).Multiply(equip_tf(src, SRC_EQUIP_LINK).Inverse)
def XF(xyz):  # source world -> target world
    p = T.OfPoint(DB.XYZ(xyz[0], xyz[1], xyz[2]))
    return p

# symbol map (family,type)->FamilySymbol id ; duplicate missing types from sibling
sym = {}
fam_types = {}
for fsym in DB.FilteredElementCollector(tgt).OfClass(DB.FamilySymbol):
    try:
        f = fsym.Family.Name
        t = fsym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except: continue
    sym[(f, t)] = fsym.Id
    fam_types.setdefault(f, []).append((t, fsym.Id))

tgt_lvl = next(L for L in DB.FilteredElementCollector(tgt).OfClass(DB.Level)
               if L.Name == TGT_LEVEL)

import os
MAP_PATH = DATA + r"\replication_map.json"
rep_map = {}
if os.path.exists(MAP_PATH):
    try: rep_map = json.load(io.open(MAP_PATH))
    except: rep_map = {}
rep_map.setdefault("devices", {})
rep_map.setdefault("panels", {})
rep_map.setdefault("flags", [])

report = []
def log(s):
    report.append(s); print(s)

hosts = json.load(io.open(DATA + r"\host_elements.json"))["elements"]

if PHASE == "A":
    # resolve panels (map by Panel Name to existing target panels; never duplicate)
    tgt_panels = {}
    for ee in DB.FilteredElementCollector(tgt).OfCategory(
            DB.BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType():
        p = ee.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        if p and p.AsString():
            tgt_panels[p.AsString()] = ee.Id.IntegerValue
    placed = skipped = dup_types = 0
    t = DB.Transaction(tgt, "BAKERY auto - Phase A devices")
    t.Start()
    try:
        for r in hosts:
            sid = str(r["id"])
            if r["category"] == "Electrical Equipment":
                pn = (r["params"].get("Panel Name") or "").strip()
                if pn in tgt_panels:
                    rep_map["panels"][pn] = tgt_panels[pn]
                    log("PANEL %s -> existing target panel id %d (not duplicated)" % (pn, tgt_panels[pn]))
                else:
                    rep_map["flags"].append("PANEL %s missing in target (src %s) - skipped, circuiting deferred" % (pn, sid))
                    log("FLAG: panel %s missing in target - skipped" % pn)
                continue
            if sid in rep_map["devices"]:
                skipped += 1; continue
            fam, typ = r["family"], r["type"]
            key = (fam, typ)
            if key not in sym:
                # duplicate from a sibling type of same family
                sibs = fam_types.get(fam)
                if not sibs:
                    rep_map["flags"].append("NO FAMILY %s in target (src %s)" % (fam, sid)); skipped += 1; continue
                base = tgt.GetElement(sibs[0][1])
                newsym = base.Duplicate(typ)
                sym[key] = newsym.Id
                fam_types[fam].append((typ, newsym.Id))
                dup_types += 1
                log("DUP TYPE created: %s / %s (from %s)" % (fam, typ, sibs[0][0]))
            fsym = tgt.GetElement(sym[key])
            if not fsym.IsActive: fsym.Activate()
            p = XF(r["world_xyz"])
            try:
                fi = tgt.Create.NewFamilyInstance(p, fsym, tgt_lvl,
                       DB.Structure.StructuralType.NonStructural)
            except Exception as e:
                rep_map["flags"].append("PLACE FAIL src %s %s/%s: %s" % (sid, fam, typ, e))
                skipped += 1; continue
            rot = r.get("rotation")
            if rot:
                ax = DB.Line.CreateBound(p, DB.XYZ(p.X, p.Y, p.Z + 1))
                try: DB.ElementTransformUtils.RotateElement(tgt, fi.Id, ax, rot)
                except: pass
            try:
                cp = fi.get_Parameter(TAGP)
                old = ""
                sc = r["params"].get("Comments")
                if sc: old = sc + " "
                if cp and not cp.IsReadOnly: cp.Set(old + "[BA:src%s]" % sid)
            except: pass
            try:
                mp = fi.get_Parameter(MARKP)
                sm = r["params"].get("Mark")
                if mp and not mp.IsReadOnly and sm: mp.Set(sm)
            except: pass
            rep_map["devices"][sid] = fi.Id.IntegerValue
            placed += 1
        t.Commit()
    except Exception as e:
        t.RollBack()
        log("PHASE A ABORTED, rolled back: %s" % e)
        raise
    log("PHASE A done: placed=%d skipped=%d dup_types=%d" % (placed, skipped, dup_types))

if PHASE == "B":
    rep_map.setdefault("keynotes", {})
    kns = json.load(io.open(DATA + r"\keynotes.json"))["elements"]
    tv = get_view(tgt)
    placed = skipped = 0
    t = DB.Transaction(tgt, "BAKERY auto - Phase B keynotes")
    t.Start()
    try:
        for r in kns:
            sid = str(r["id"])
            if sid in rep_map["keynotes"]:
                skipped += 1; continue
            fam, typ = r["family"], r["type"]
            key = (fam, typ)
            if key not in sym:
                sibs = fam_types.get(fam)
                if not sibs:
                    rep_map["flags"].append("NO KEYNOTE FAMILY %s (src %s)" % (fam, sid))
                    skipped += 1; continue
                newsym = tgt.GetElement(sibs[0][1]).Duplicate(typ)
                sym[key] = newsym.Id; fam_types[fam].append((typ, newsym.Id))
                log("DUP KEYNOTE TYPE: %s / %s" % (fam, typ))
            fsym = tgt.GetElement(sym[key])
            if not fsym.IsActive: fsym.Activate()
            p = XF(r["world_xyz"])
            try:
                fi = tgt.Create.NewFamilyInstance(p, fsym, tv)
            except Exception as e:
                rep_map["flags"].append("KEYNOTE PLACE FAIL src %s: %s" % (sid, e))
                skipped += 1; continue
            rot = r.get("rotation")
            if rot:
                ax = DB.Line.CreateBound(p, DB.XYZ(p.X, p.Y, p.Z + 1))
                try: DB.ElementTransformUtils.RotateElement(tgt, fi.Id, ax, rot)
                except: pass
            for pname in ("Keynote Value", "Symbol Text", "Show Symbol Text"):
                try:
                    sv = r["params"].get(pname)
                    if sv is None: continue
                    pr = fi.LookupParameter(pname)
                    if pr and not pr.IsReadOnly:
                        if pr.StorageType == DB.StorageType.Integer: pr.Set(int(sv))
                        else: pr.Set(str(sv))
                except: pass
            rep_map["keynotes"][sid] = fi.Id.IntegerValue
            placed += 1
        t.Commit()
    except Exception as e:
        t.RollBack(); log("PHASE B ABORTED, rolled back: %s" % e); raise
    log("PHASE B done: placed=%d skipped=%d" % (placed, skipped))

if PHASE == "C":
    rep_map.setdefault("wires", {})
    wires = json.load(io.open(DATA + r"\wires.json"))["elements"]
    tv = get_view(tgt)
    # resolve / create wire type
    wt_id = None
    want = wires[0]["wire_type"] if wires else "THWN"
    for wt in DB.FilteredElementCollector(tgt).OfClass(DB.Electrical.WireType):
        try:
            if DB.Element.Name.GetValue(wt) == want: wt_id = wt.Id; break
        except: pass
    placed = skipped = 0
    t = DB.Transaction(tgt, "BAKERY auto - Phase C wires")
    t.Start()
    try:
        if wt_id is None:
            nwt = DB.Electrical.WireType.Create(tgt, want)
            wt_id = nwt.Id
            log("CREATED WireType '%s' id %d" % (want, wt_id.IntegerValue))
        for r in wires:
            sid = str(r["id"])
            if sid in rep_map["wires"]:
                skipped += 1; continue
            pts = [XF(v) for v in r["vertices"]]
            vlist = List[DB.XYZ]()
            for p in pts: vlist.Add(p)
            try:
                w = DB.Electrical.Wire.Create(tgt, wt_id, tv.Id,
                        DB.Electrical.WiringType.Arc, vlist, None, None)
            except Exception as e:
                rep_map["flags"].append("WIRE FAIL src %s: %s" % (sid, e))
                skipped += 1; continue
            rep_map["wires"][sid] = w.Id.IntegerValue
            placed += 1
        t.Commit()
    except Exception as e:
        t.RollBack(); log("PHASE C ABORTED, rolled back: %s" % e); raise
    log("PHASE C done: placed=%d skipped=%d" % (placed, skipped))

if PHASE == "D":
    rep_map.setdefault("circuits", {})
    OK_PANELS = {"BA", "RC"}            # existing target branch panels only
    # target panel elements by name
    tp = {}
    for ee in DB.FilteredElementCollector(tgt).OfCategory(
            DB.BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType():
        pn = ee.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        if pn and pn.AsString() in OK_PANELS:
            tp[pn.AsString()] = ee
    # group placed target devices by source circuit sys_id (panel in OK_PANELS)
    groups = {}   # sys_id -> {"panel":p, "ids":[targetElementId]}
    for r in hosts:
        sid = str(r["id"])
        if sid not in rep_map["devices"]: continue
        for c in r.get("circuits", []):
            pnl = c.get("panel"); sys = c.get("sys_id")
            if pnl in OK_PANELS and sys:
                g = groups.setdefault(sys, {"panel": pnl, "ids": []})
                g["ids"].append(DB.ElementId(rep_map["devices"][sid]))
    # source circuit descriptions, keyed by source sys_id (str)
    desc = {}
    for c in json.load(io.open(DATA + r"\circuits.json"))["circuits"]:
        desc[str(c["sys_id"])] = c
    DESC_MAP = [  # (param name, source circuits.json key)
        ("Load Name",                "load_name"),
        ("Schedule Circuit Notes",   "schedule_notes"),
        ("CKT_Load Name_CEDT",       "ckt_load_name_cedt"),
        ("CKT_Schedule Notes_CEDT",  "ckt_schedule_notes_cedt"),
    ]
    def set_desc(es, sysstr):
        d = desc.get(sysstr)
        if not d: return None
        for pname, k in DESC_MAP:
            v = d.get(k)
            if v in (None, ""): continue
            try:
                pr = es.LookupParameter(pname)
                if pr and not pr.IsReadOnly and pr.StorageType == DB.StorageType.String:
                    pr.Set(v)
            except: pass
        return d.get("load_name")

    made = failed = updated = 0
    t = DB.Transaction(tgt, "BAKERY auto - Phase D circuiting + descriptions")
    t.Start()
    try:
        for sys, g in groups.items():
            sstr = str(sys)
            rec = rep_map["circuits"].get(sstr)
            es = None
            if rec:
                es = tgt.GetElement(DB.ElementId(int(rec["tgt_sys_id"])))
            if es is not None:
                # refresh description on the already-created circuit (idempotent)
                ln = set_desc(es, sstr)
                rec["load_name"] = ln
                updated += 1
                continue
            # not present (new, or undone in Revit) -> (re)create
            if rec:
                del rep_map["circuits"][sstr]
            ids = List[DB.ElementId]()
            for i in g["ids"]: ids.Add(i)
            try:
                es = DB.Electrical.ElectricalSystem.Create(
                        tgt, ids, DB.Electrical.ElectricalSystemType.PowerCircuit)
                es.SelectPanel(tp[g["panel"]])
                ln = set_desc(es, sstr)
                cn = es.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
                rep_map["circuits"][sstr] = {
                    "tgt_sys_id": es.Id.IntegerValue, "panel": g["panel"],
                    "auto_ckt": cn.AsString() if cn else None,
                    "load_name": ln, "n_devices": len(g["ids"])}
                made += 1
            except Exception as e:
                rep_map["flags"].append("CIRCUIT FAIL src_sys %s panel %s: %s"
                                        % (sys, g["panel"], e))
                failed += 1
        t.Commit()
    except Exception as e:
        t.RollBack(); log("PHASE D ABORTED, rolled back: %s" % e); raise
    log("PHASE D done: created=%d desc-updated=%d failed=%d (FS/other panels not circuited)"
        % (made, updated, failed))

with io.open(MAP_PATH, 'w', encoding='utf-8') as f:
    f.write(json.dumps(rep_map, indent=1))
with io.open(DATA + r"\replication_report.txt", 'a', encoding='utf-8') as f:
    f.write("\n=== PHASE %s ===\n" % PHASE + "\n".join(report) + "\n")
print("map:", MAP_PATH)
