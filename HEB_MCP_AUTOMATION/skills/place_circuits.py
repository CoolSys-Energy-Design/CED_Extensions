# SKILL: place_circuits  (Oakmont; runs AFTER textnotes, BEFORE wires)
# Groups placed devices by their SOURCE circuit (sys_id) and creates one
# ElectricalSystem per group, assigned to the matching target panel by name.
# Auto circuit numbers; copies source Load Name / Schedule Notes / CED mirror.
# Globals: BA_C_APPLY (default False), BA_C_REBUILD, BA_TGT_TITLE,
#          BA_C_PANELS (default {"BA","RC","FS"})
# Writes: data\place_circuits_report.json ; updates place_relative_map.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
from System.Collections.Generic import List

APPLY   = globals().get("BA_C_APPLY", False)
REBUILD = globals().get("BA_C_REBUILD", False)
TGT     = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)

hosts = json.load(io.open(DATA + r"\host_elements.json"))["elements"]
desc  = {str(c["sys_id"]): c for c in
         json.load(io.open(DATA + r"\circuits.json"))["circuits"]}
MAP = DATA + r"\place_relative_map.json"
pm  = json.load(io.open(MAP)); pm.setdefault("circuits", {})
dev_map = pm["devices"]

# panels that actually exist in the target, by name
tgt_panels = {}
for ee in DB.FilteredElementCollector(tgt).OfCategory(
        DB.BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType():
    p = ee.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
    if p and p.AsString():
        tgt_panels[p.AsString()] = ee

# OK_PAN = panels the source devices are on AND that exist in target
# (override with BA_C_PANELS if explicitly provided)
_override = globals().get("BA_C_PANELS")
if _override:
    OK_PAN = set(_override)
else:
    src_pan = set()
    for r in hosts:
        if str(r["id"]) in dev_map:
            for c in r.get("circuits", []):
                if c.get("panel"): src_pan.add(c["panel"])
    OK_PAN = set(p for p in src_pan if p in tgt_panels)
print("circuit panels in scope:", sorted(OK_PAN))
tp = dict((k, v) for k, v in tgt_panels.items() if k in OK_PAN)

# group placed devices by source circuit sys_id (panel in OK_PAN)
groups = {}
for r in hosts:
    sid = str(r["id"])
    if sid not in dev_map: continue
    for c in r.get("circuits", []):
        pnl, sys = c.get("panel"), c.get("sys_id")
        if pnl in OK_PAN and sys:
            g = groups.setdefault(str(sys), {"panel": pnl, "ids": []})
            g["ids"].append(DB.ElementId(int(dev_map[sid])))

DESC_MAP = [("Load Name", "load_name"),
            ("Schedule Circuit Notes", "schedule_notes"),
            ("CKT_Load Name_CEDT", "ckt_load_name_cedt"),
            ("CKT_Schedule Notes_CEDT", "ckt_schedule_notes_cedt")]
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

import re as _re
def _amps(s):
    if s is None: return None
    m = _re.search(r"[\d.]+", str(s))
    return float(m.group(0)) if m else None

def set_rating_poles(es, sysstr):
    """Set breaker rating / poles / frame from the SOURCE circuit so target
    circuits aren't all defaulted to 20 A / 1-pole."""
    d = desc.get(sysstr)
    if not d: return
    # poles first (affects how the panel slot is taken)
    try:
        pv = d.get("poles")
        if pv not in (None, ""):
            pr = es.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES)
            if pr and not pr.IsReadOnly: pr.Set(int(_amps(pv)))
    except: pass
    for bip, key in ((DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM, "rating"),
                     (DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM, "frame")):
        try:
            a = _amps(d.get(key))
            if a is None: continue
            pr = es.get_Parameter(bip)
            if pr and not pr.IsReadOnly: pr.Set(float(a))
        except: pass

made = updated = failed = 0
rows = []
tg = DB.TransactionGroup(tgt, "BAKERY circuits (Oakmont)")
if APPLY: tg.Start()
t = DB.Transaction(tgt, "circuit devices")
if APPLY: t.Start()
try:
    if APPLY and REBUILD:
        for k, rec in list(pm["circuits"].items()):
            es = tgt.GetElement(DB.ElementId(int(rec["tgt_sys_id"])))
            if es is not None:
                try: tgt.Delete(es.Id)
                except: pass
        pm["circuits"] = {}
    for sysstr, g in groups.items():
        rec = pm["circuits"].get(sysstr)
        es = tgt.GetElement(DB.ElementId(int(rec["tgt_sys_id"]))) if rec else None
        if es is not None:
            ln = set_desc(es, sysstr); set_rating_poles(es, sysstr)
            rec["load_name"] = ln; updated += 1; continue
        if rec: del pm["circuits"][sysstr]
        if not APPLY:
            rows.append({"src_sys": sysstr, "panel": g["panel"],
                         "n_dev": len(g["ids"]), "status": "would-create"}); continue
        ids = List[DB.ElementId]()
        for i in g["ids"]: ids.Add(i)
        try:
            es = DB.Electrical.ElectricalSystem.Create(
                    tgt, ids, DB.Electrical.ElectricalSystemType.PowerCircuit)
            es.SelectPanel(tp[g["panel"]])
            ln = set_desc(es, sysstr)
            set_rating_poles(es, sysstr)
            pm["circuits"][sysstr] = {"tgt_sys_id": es.Id.IntegerValue,
                "panel": g["panel"], "load_name": ln, "n_devices": len(g["ids"])}
            made += 1
        except Exception as e:
            rows.append({"src_sys": sysstr, "panel": g["panel"],
                         "status": "FAIL:%s" % e}); failed += 1
    if APPLY:
        t.Commit(); tg.Assimilate()   # commit regenerates the model
except Exception as ex:
    if APPLY:
        t.RollBack()
        if tg.GetStatus() == DB.TransactionStatus.Started: tg.RollBack()
    print("CIRCUITS ABORTED, rolled back:", ex); raise

if APPLY:
    with io.open(MAP, 'w', encoding='utf-8') as f:
        f.write(json.dumps(pm, indent=1))
write_json("place_circuits_report.json",
           {"target": tgt.Title, "apply": APPLY, "groups": len(groups),
            "made": made, "updated": updated, "failed": failed, "rows": rows})
print("=== CIRCUITS %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
bp = {}
for g in groups.values(): bp[g["panel"]] = bp.get(g["panel"], 0) + 1
print("groups:%d by panel=%s | made=%d updated=%d failed=%d"
      % (len(groups), bp, made, updated, failed))
for r in rows:
    if str(r.get("status","")).startswith("FAIL"): print("  !", r)
print("written: data\\place_circuits_report.json")
