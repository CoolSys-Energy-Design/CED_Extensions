# SKILL: collect_host_elements
# Collects Electrical Fixtures + Electrical Equipment inside the BAKERY crop:
# full params, location, facing/hand, host, level, circuits, AND the
# nearest linked equipment anchor + relative offset vector / orientation.
# Requires linked_elements.json (run collect_linked_elements first).
# Writes: data\host_elements.json
import json, io, math
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())  # nm, type_name

LINKED = DATA + r"\linked_elements.json"   # VIEW_NAME, DATA from _lib
OUT = DATA + r"\host_elements.json"
HOST_CATS = ("Electrical Fixtures", "Electrical Equipment")
ANCHOR_CATS = ("Specialty Equipment", "Mechanical Equipment", "Generic Models",
               "Plumbing Fixtures", "Lighting Fixtures", "Casework",
               "Food Service Equipment", "Electrical Equipment")

def nm(e):
    try:
        v = DB.Element.Name.GetValue(e)
        if v: return v
    except: pass
    try: return getattr(e, 'Name', '?')
    except: return '?'

anchors = json.load(io.open(LINKED, 'r', encoding='utf-8'))["elements"]
anchors = [a for a in anchors if a["category"] in ANCHOR_CATS and a.get("world_xyz")]

bv = next(v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
          if getattr(v, 'Name', '') == VIEW_NAME and not v.IsTemplate)
cb = bv.CropBox; inv = cb.Transform.Inverse
def in_crop(p):
    q = inv.OfPoint(p)
    return (cb.Min.X-0.5 <= q.X <= cb.Max.X+0.5 and cb.Min.Y-0.5 <= q.Y <= cb.Max.Y+0.5)

def params_of(el):
    d = {}
    for p in el.Parameters:
        try:
            n = p.Definition.Name
            st = p.StorageType
            if st == DB.StorageType.String: d[n] = p.AsString()
            elif st == DB.StorageType.Double: d[n] = round(p.AsDouble(), 5)
            elif st == DB.StorageType.Integer: d[n] = p.AsInteger()
            elif st == DB.StorageType.ElementId: d[n] = p.AsElementId().IntegerValue
        except: pass
    return d

def wparams_of(el):
    """Writable INSTANCE parameter values, with metadata so they can be
    re-applied in a different project. ElementId-valued params are skipped
    (ids are not portable across projects)."""
    out = []
    for p in el.Parameters:
        try:
            if p.IsReadOnly: continue
            st = p.StorageType
            if st == DB.StorageType.ElementId: continue   # not portable
            v = None
            if st == DB.StorageType.String:   v = p.AsString()
            elif st == DB.StorageType.Double: v = p.AsDouble()
            elif st == DB.StorageType.Integer:v = p.AsInteger()
            else: continue
            if v is None or v == "": continue
            bip = None
            try:
                d = p.Definition
                if isinstance(d, DB.InternalDefinition):
                    b = d.BuiltInParameter
                    if b != DB.BuiltInParameter.INVALID: bip = str(b)
            except: pass
            out.append({"n": p.Definition.Name,
                        "st": str(st),
                        "v": (round(v, 6) if st == DB.StorageType.Double else v),
                        "bip": bip,
                        "shared": bool(getattr(p, "IsShared", False))})
        except: pass
    return out

def nearest_anchor(x, y):
    best = None; bd = 1e18
    for a in anchors:
        ax, ay = a["world_xyz"][0], a["world_xyz"][1]
        dd = (ax-x)**2 + (ay-y)**2
        if dd < bd: bd = dd; best = a
    if best is None: return None
    return {"anchor_id": best["id"], "anchor_family": best.get("family"),
            "anchor_type": best.get("type"), "anchor_mark": best.get("mark"),
            "anchor_category": best["category"],
            "dxy_ft": [round(x-best["world_xyz"][0],4), round(y-best["world_xyz"][1],4)],
            "dist_ft": round(math.sqrt(bd), 4)}

recs = []
for el in DB.FilteredElementCollector(doc, bv.Id).WhereElementIsNotElementType():
    try: cat = el.Category.Name if el.Category else None
    except: cat = None
    if cat not in HOST_CATS: continue
    L = el.Location
    if not isinstance(L, DB.LocationPoint): continue
    p = L.Point
    if not in_crop(p): continue
    fs = doc.GetElement(el.GetTypeId())
    fo = getattr(el, 'FacingOrientation', None)
    ho = getattr(el, 'HandOrientation', None)
    cir = []
    try:
        mm = el.MEPModel
        syss = mm.GetElectricalSystems() if mm else None
        if syss:
            for s in syss:
                pn = s.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
                cn = s.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
                cir.append({"sys_id": s.Id.IntegerValue,
                            "panel": pn.AsString() if pn else None,
                            "ckt": cn.AsString() if cn else None})
    except: pass
    host = None
    try:
        h = getattr(el, 'Host', None)
        if h: host = {"id": h.Id.IntegerValue,
                       "cat": h.Category.Name if h.Category else None}
    except: pass
    recs.append({
        "id": el.Id.IntegerValue, "category": cat,
        "family": nm(fs.Family) if fs and hasattr(fs,'Family') else None,
        "type": type_name(el),
        "world_xyz": [round(p.X,4), round(p.Y,4), round(p.Z,4)],
        "facing": [round(fo.X,4), round(fo.Y,4)] if fo else None,
        "hand": [round(ho.X,4), round(ho.Y,4)] if ho else None,
        "rotation": round(L.Rotation, 6) if hasattr(L, 'Rotation') else None,
        "level": nm(doc.GetElement(el.LevelId)) if el.LevelId and el.LevelId.IntegerValue > 0 else None,
        "level_elev": (lambda L0: (round(L0.Elevation, 5) if L0 else 0.0))(
            doc.GetElement(el.LevelId) if el.LevelId and el.LevelId.IntegerValue > 0 else None),
        "host": host, "circuits": cir,
        "rel_to_equipment": nearest_anchor(p.X, p.Y),
        "params": params_of(el),
        "wparams": wparams_of(el),
    })

with io.open(OUT, 'w', encoding='utf-8') as f:
    f.write(json.dumps({"view": VIEW_NAME, "source_doc": doc.Title,
                        "count": len(recs), "elements": recs}, indent=1))
bycat = {}
for r in recs:
    k = r["category"] + " | " + (r["family"] or "?")
    bycat[k] = bycat.get(k, 0) + 1
print("host elements in crop:", len(recs))
for k in sorted(bycat, key=lambda z: -bycat[z]): print("  %-45s %d" % (k, bycat[k]))
print("written:", OUT)
