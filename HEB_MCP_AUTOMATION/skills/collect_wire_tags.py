# SKILL: collect_wire_tags
# Collects DB.IndependentTag (category "Wire Tags") in the BAKERY view crop:
# head xyz, tag text, type/family, tagged Wire id + category/wire-type,
# leader geo, orientation/rotation, and offset from the nearest wire vertex.
# Writes: data\wire_tags.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
# Always read from the SOURCE project regardless of which doc is active
_alldocs = list(doc.Application.Documents)
doc = next(d for d in _alldocs if d.Title.startswith("RunUpdateProfiles"))
print("collecting wire tags from:", doc.Title)

bv = get_view(doc)
in_crop = crop_test(bv)
recs = []
for tag in DB.FilteredElementCollector(doc, bv.Id).OfClass(DB.IndependentTag):
    try: cat = tag.Category.Name if tag.Category else None
    except: cat = None
    if cat != "Wire Tags": continue
    try: head = tag.TagHeadPosition
    except: head = None
    if head is None or not in_crop(head): continue

    # tagged element id (robust)
    tid = None
    try:
        ids = list(tag.GetTaggedLocalElementIds())
        if ids: tid = ids[0].IntegerValue
    except: pass
    if tid is None:
        try: tid = tag.TaggedLocalElementId.IntegerValue
        except: pass
    tcat = twtype = None
    te = None
    if tid is not None:
        try:
            te = doc.GetElement(DB.ElementId(tid))
            if te is not None:
                try: tcat = te.Category.Name if te.Category else None
                except: tcat = None
                try: twtype = type_name(te)
                except: twtype = None
        except: pass

    try: ttext = tag.TagText
    except: ttext = None
    try: tfamname = nm(doc.GetElement(tag.GetTypeId()).Family)
    except: tfamname = None

    has_leader = False
    try: has_leader = bool(tag.HasLeader)
    except: pass
    leader_end = leader_elbow = lec = None
    if has_leader:
        try: lec = str(tag.LeaderEndCondition)
        except: lec = None
        try:
            refs = list(tag.GetTaggedReferences())
            if refs:
                try:
                    v = tag.GetLeaderEnd(refs[0])
                    leader_end = [round(v.X,4), round(v.Y,4), round(v.Z,4)]
                except: leader_end = None
                try:
                    v = tag.GetLeaderElbow(refs[0])
                    leader_elbow = [round(v.X,4), round(v.Y,4), round(v.Z,4)]
                except: leader_elbow = None
        except: pass

    try: orientation = str(tag.TagOrientation)
    except: orientation = None
    try: rotation = round(tag.RotationAngle, 6)
    except: rotation = None

    # offset from the tagged wire's nearest vertex (in XY) to the tag head
    offset = None
    near_idx = None
    nverts = None
    try:
        if te is not None and hasattr(te, "NumberOfVertices"):
            nverts = te.NumberOfVertices
            best_d = None
            best_v = None
            for i in range(nverts):
                v = te.GetVertex(i)
                d = (v.X-head.X)**2 + (v.Y-head.Y)**2
                if best_d is None or d < best_d:
                    best_d = d; best_v = v; near_idx = i
            if best_v is not None:
                offset = [round(head.X-best_v.X,4), round(head.Y-best_v.Y,4),
                          round(head.Z-best_v.Z,4)]
    except:
        offset = None

    recs.append({
        "id": tag.Id.IntegerValue,
        "tagged_id": tid,
        "tagged_category": tcat,
        "tagged_wire_type": twtype,
        "head": [round(head.X,4), round(head.Y,4), round(head.Z,4)],
        "tag_text": ttext,
        "type": type_name(tag),
        "tag_family": tfamname,
        "has_leader": has_leader,
        "leader_end": leader_end,
        "leader_elbow": leader_elbow,
        "leader_end_condition": lec,
        "orientation": orientation,
        "rotation": rotation,
        "offset_from_wire": offset,
        "near_vertex_index": near_idx,
        "wire_num_vertices": nverts,
    })

path = write_json("wire_tags.json", {"view":VIEW_NAME,"source_doc":doc.Title,
                   "count":len(recs),"elements":recs})
print("wire tags in crop:", len(recs))
for r in recs[:3]:
    print(" id",r["id"],"text",repr(r["tag_text"]),
          "tagged_id",r["tagged_id"],"has_leader",r["has_leader"])
print("written:", path)
