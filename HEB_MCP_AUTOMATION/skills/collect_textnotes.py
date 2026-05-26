# SKILL: collect_textnotes
# Collects DB.TextNote elements in the BAKERY view crop: text, type, world
# coord, width, rotation, alignment, leaders, plus nearest host device and
# nearest equipment anchor (for relative placement later).
# Requires linked_elements.json + host_elements.json.
# Writes: data\textnotes.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
# Always read from the SOURCE project regardless of which doc is active
doc = next(d for d in doc.Application.Documents
           if d.Title.startswith("RunUpdateProfiles"))
print("collecting textnotes from:", doc.Title)

anchors = [a for a in json.load(io.open(DATA+r"\linked_elements.json"))["elements"]
           if a["category"] in ("Specialty Equipment","Mechanical Equipment",
           "Generic Models","Plumbing Fixtures","Lighting Fixtures","Electrical Equipment")
           and a.get("world_xyz")]
hosts = json.load(io.open(DATA+r"\host_elements.json"))["elements"]

def nearest(x, y, pool):
    best=None; bd=1e18
    for a in pool:
        w=a.get("world_xyz")
        if not w: continue
        d=(w[0]-x)**2+(w[1]-y)**2
        if d<bd: bd=d; best=a
    return best, (math.sqrt(bd) if best else None)

bv = get_view(doc)
in_crop = crop_test(bv)
recs=[]
for tn in DB.FilteredElementCollector(doc, bv.Id).OfClass(DB.TextNote):
    try: p = tn.Coord
    except: continue
    if p is None or not in_crop(p): continue
    try: txt = tn.Text
    except: txt = None
    try: width = round(tn.Width, 5)
    except: width = None
    try: rot = round(tn.RotationAngle, 6)
    except: rot = None
    try: halign = str(tn.HorizontalAlignment)
    except: halign = None
    leaders = None
    leader_geo = []
    try:
        lds = tn.GetLeaders()
        leaders = bool(lds)
        for ld in (lds or []):
            g = {}
            for attr in ("Anchor", "End", "Elbow"):
                try:
                    v = getattr(ld, attr)
                    g[attr.lower()] = [round(v.X,4), round(v.Y,4), round(v.Z,4)]
                except: g[attr.lower()] = None
            try: g["has_elbow"] = bool(ld.HasElbow)
            except: g["has_elbow"] = (g.get("elbow") is not None)
            leader_geo.append(g)
    except: pass
    try: tnt = nm(doc.GetElement(tn.GetTypeId()))
    except: tnt = None
    na=ne=None
    a,da = nearest(p.X,p.Y,anchors)
    if a: na={"id":a["id"],"family":a.get("family"),"category":a["category"],
              "dxy_ft":[round(p.X-a["world_xyz"][0],4),round(p.Y-a["world_xyz"][1],4)],
              "dist_ft":round(da,4)}
    h,dh = nearest(p.X,p.Y,hosts)
    if h: ne={"id":h["id"],"family":h.get("family"),
              "dxy_ft":[round(p.X-h["world_xyz"][0],4),round(p.Y-h["world_xyz"][1],4)],
              "dist_ft":round(dh,4)}
    recs.append({
        "id": tn.Id.IntegerValue,
        "text": txt,
        "type": type_name(tn),
        "textnote_type": tnt,
        "coord": [round(p.X,4),round(p.Y,4),round(p.Z,4)],
        "width": width,
        "rotation": rot,
        "h_align": halign,
        "has_leaders": leaders,
        "leaders": leader_geo,
        "owner_view": tn.OwnerViewId.IntegerValue if tn.OwnerViewId else None,
        "rel_to_equipment": na,
        "rel_to_device": ne,
    })

path = write_json("textnotes.json", {"view":VIEW_NAME,"source_doc":doc.Title,
                   "count":len(recs),"elements":recs})
print("textnotes in crop:", len(recs))
for r in recs[:3]:
    t = (r["text"] or "").replace("\r"," ").replace("\n"," ").strip()
    print(" id",r["id"],"text:",repr(t[:60]),
          "->dev",r["rel_to_device"]["family"] if r["rel_to_device"] else None)
print("written:", path)
