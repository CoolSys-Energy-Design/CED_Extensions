# SKILL: collect_keynotes
# Collects GA_Keynote Symbol_CED (Generic Annotation) instances in the BAKERY
# crop: location, rotation, all params (keynote value/text), leader, and
# nearest host device + nearest equipment anchor (keynotes label devices).
# Requires linked_elements.json + host_elements.json.
# Writes: data\keynotes.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
FAM = "GA_Keynote Symbol_CED"

anchors = [a for a in json.load(io.open(DATA+r"\linked_elements.json"))["elements"]
           if a["category"] in ("Specialty Equipment","Mechanical Equipment",
           "Generic Models","Plumbing Fixtures","Lighting Fixtures","Electrical Equipment")
           and a.get("world_xyz")]
hosts = json.load(io.open(DATA+r"\host_elements.json"))["elements"]

def nearest(x, y, pool, kx, ky):
    best=None; bd=1e18
    for a in pool:
        px = a[kx][0] if isinstance(a[kx],list) else a[kx]
        py = a[kx][1] if isinstance(a[kx],list) else a[ky]
        d=(px-x)**2+(py-y)**2
        if d<bd: bd=d; best=a
    return best, math.sqrt(bd) if best else None

bv = get_view(doc)
in_crop = crop_test(bv)
recs=[]
for el in DB.FilteredElementCollector(doc, bv.Id).WhereElementIsNotElementType():
    try: cat = el.Category.Name if el.Category else None
    except: cat=None
    if cat != "Generic Annotations": continue
    fs = doc.GetElement(el.GetTypeId())
    fam = nm(fs.Family) if fs and hasattr(fs,'Family') else None
    if fam != FAM: continue
    L = el.Location
    if not isinstance(L, DB.LocationPoint): continue
    p = L.Point
    if not in_crop(p): continue
    # leader info
    leader=None
    try:
        if getattr(el,'HasLeader',False):
            leader=True
    except: pass
    na = ne = None; nad=ned=None
    if anchors:
        a,da = nearest(p.X,p.Y,anchors,"world_xyz","world_xyz")
        if a: na={"id":a["id"],"family":a.get("family"),"category":a["category"],
                  "dxy_ft":[round(p.X-a["world_xyz"][0],4),round(p.Y-a["world_xyz"][1],4)],
                  "dist_ft":round(da,4)}
    h,dh = nearest(p.X,p.Y,hosts,"world_xyz","world_xyz")
    if h: ne={"id":h["id"],"family":h.get("family"),
              "dxy_ft":[round(p.X-h["world_xyz"][0],4),round(p.Y-h["world_xyz"][1],4)],
              "dist_ft":round(dh,4)}
    recs.append({
        "id": el.Id.IntegerValue,
        "family": fam, "type": type_name(el),
        "world_xyz": [round(p.X,4),round(p.Y,4),round(p.Z,4)],
        "rotation": round(L.Rotation,6) if hasattr(L,'Rotation') else None,
        "owner_view": el.OwnerViewId.IntegerValue if el.OwnerViewId else None,
        "has_leader": leader,
        "rel_to_equipment": na, "rel_to_device": ne,
        "params": params_of(el),
    })

path = write_json("keynotes.json", {"view":VIEW_NAME,"source_doc":doc.Title,
                   "family":FAM,"count":len(recs),"elements":recs})
print("keynotes in crop:", len(recs))
if recs:
    print("sample params keys:", list(recs[0]["params"].keys()))
    for r in recs[:3]:
        print(" id",r["id"],"type",r["type"],"xy",r["world_xyz"][:2],
              "->dev",r["rel_to_device"]["family"] if r["rel_to_device"] else None,
              "params:", {k:v for k,v in r["params"].items() if v not in (None,"")})
print("written:", path)
