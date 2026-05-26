# SKILL: collect_circuits
# Builds the bakery circuit list from system ids referenced by host devices +
# wires (host_elements.json, wires.json). For each ElectricalSystem: panel,
# circuit number, rating, load name, member element ids/categories.
# Writes: data\circuits.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
# Always read from the SOURCE project regardless of which doc is active
doc = next(d for d in doc.Application.Documents
           if d.Title.startswith("RunUpdateProfiles"))
print("collecting circuits from:", doc.Title)

sysids=set()
for r in json.load(io.open(DATA+r"\host_elements.json"))["elements"]:
    for c in r.get("circuits",[]):
        if c.get("sys_id"): sysids.add(c["sys_id"])
for r in json.load(io.open(DATA+r"\wires.json"))["elements"]:
    s=r["circuit"].get("sys_id")
    if s: sysids.add(s)

def gp(el,bip):
    try:
        p=el.get_Parameter(bip)
        if p:
            return p.AsString() or p.AsValueString()
    except: pass
    return None

recs=[]
for sid in sorted(sysids):
    s=doc.GetElement(DB.ElementId(sid))
    if s is None: continue
    members=[]
    try:
        for e in s.Elements:
            members.append({"id":e.Id.IntegerValue,
                            "cat":e.Category.Name if e.Category else None})
    except: pass
    base=None
    try:
        be=s.BaseEquipment
        if be: base={"id":be.Id.IntegerValue,"name":nm(be)}
    except: pass
    recs.append({
        "sys_id": sid,
        "panel": gp(s,DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM),
        "ckt": gp(s,DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER),
        "rating": gp(s,DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM),
        "frame": gp(s,DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM),
        "load_name": gp(s,DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME),
        "schedule_notes": (lambda p: (p.AsString() if p else None))(s.LookupParameter("Schedule Circuit Notes")),
        "ckt_load_name_cedt": (lambda p: (p.AsString() if p else None))(s.LookupParameter("CKT_Load Name_CEDT")),
        "ckt_schedule_notes_cedt": (lambda p: (p.AsString() if p else None))(s.LookupParameter("CKT_Schedule Notes_CEDT")),
        "load_classification": gp(s,DB.BuiltInParameter.RBS_ELEC_LOAD_CLASSIFICATION),
        "voltage": gp(s,DB.BuiltInParameter.RBS_ELEC_VOLTAGE),
        "poles": gp(s,DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES),
        "wire_type": gp(s,DB.BuiltInParameter.RBS_ELEC_WIRE_TYPE),
        "base_equipment": base,
        "members": members,
    })

path = write_json("circuits.json", {"view":VIEW_NAME,"source_doc":doc.Title,
                   "count":len(recs),"circuits":recs})
print("circuits scoped to %r:" % VIEW_NAME, len(recs))
bp={}
for r in recs: bp[r["panel"] or "?"]=bp.get(r["panel"] or "?",0)+1
for k in sorted(bp,key=lambda z:-bp[z]): print("  panel %-8s %d" % (k,bp[k]))
if recs:
    r=recs[0]; print("sample:",{k:r[k] for k in ("panel","ckt","rating","load_name","voltage","poles")},
          "members",len(r["members"]))
print("written:", path)
