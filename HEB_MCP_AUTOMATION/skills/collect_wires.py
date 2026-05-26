# SKILL: collect_wires
# Collects Wires with any vertex inside the BAKERY crop: vertex polyline
# ("vectors"), wire type, size params, owning circuit (panel/ckt), and the
# elements at the wire endpoints (connectivity).
# Writes: data\wires.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())

bv = get_view(doc)
in_crop = crop_test(bv)
recs=[]
for w in DB.FilteredElementCollector(doc, bv.Id).WhereElementIsNotElementType():
    try: cat = w.Category.Name if w.Category else None
    except: cat=None
    if cat != "Wires": continue
    try: nv = w.NumberOfVertices
    except: continue
    verts=[]
    for i in range(nv):
        v=w.GetVertex(i); verts.append([round(v.X,4),round(v.Y,4),round(v.Z,4)])
    if not any(in_crop(DB.XYZ(x,y,z)) for x,y,z in verts): continue
    # owning circuit / system
    panel=ckt=sysid=None
    try:
        ms = w.MEPSystem
        if ms:
            sysid = ms.Id.IntegerValue
            pn = ms.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_PANEL_PARAM)
            cn = ms.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
            panel = pn.AsString() if pn else None
            ckt = cn.AsString() if cn else None
    except: pass
    # endpoint connectivity
    conn=[]
    try:
        cm = w.ConnectorManager
        if cm:
            for c in cm.Connectors:
                for r in c.AllRefs:
                    o = r.Owner
                    if o and o.Id != w.Id:
                        conn.append({"id":o.Id.IntegerValue,
                                     "cat":o.Category.Name if o.Category else None})
    except: pass
    wt = doc.GetElement(w.GetTypeId())
    recs.append({
        "id": w.Id.IntegerValue,
        "wire_type": nm(wt),
        "num_vertices": nv,
        "vertices": verts,
        "circuit": {"sys_id":sysid,"panel":panel,"ckt":ckt},
        "connected_to": conn,
        "params": params_of(w, only=("Size","Wire Size","Wire Type","Hot Conductors",
                  "Neutral Conductors","Ground Conductors","Comments","Mark")),
    })

path = write_json("wires.json", {"view":VIEW_NAME,"source_doc":doc.Title,
                   "count":len(recs),"elements":recs})
print("wires in crop:", len(recs))
bypanel={}
for r in recs:
    k=r["circuit"]["panel"] or "<none>"; bypanel[k]=bypanel.get(k,0)+1
for k in sorted(bypanel,key=lambda z:-bypanel[z]): print("  panel %-8s %d" % (k,bypanel[k]))
if recs:
    s=recs[0]; print("sample:", s["id"], s["wire_type"], "nv",s["num_vertices"],
                      "ckt",s["circuit"], "conn",s["connected_to"])
print("written:", path)
