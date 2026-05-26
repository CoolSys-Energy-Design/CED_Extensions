# SKILL: collect_fixture_tags
# Collects DB.IndependentTag (category "Electrical Fixture Tags") in the
# BAKERY view crop: head xyz, tag text, type/family, tagged element id +
# category/family, leader geo, orientation/rotation, and offset from host.
# Writes: data\fixture_tags.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
# Always read from the SOURCE project regardless of which doc is active
_alldocs = list(doc.Application.Documents)
doc = next(d for d in _alldocs if d.Title.startswith("RunUpdateProfiles"))
print("collecting fixture tags from:", doc.Title)

bv = get_view(doc)
in_crop = crop_test(bv)
recs = []
for tag in DB.FilteredElementCollector(doc, bv.Id).OfClass(DB.IndependentTag):
    try: cat = tag.Category.Name if tag.Category else None
    except: cat = None
    if cat != "Electrical Fixture Tags": continue
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
    tcat = tfam = None
    te = None
    if tid is not None:
        try:
            te = doc.GetElement(DB.ElementId(tid))
            if te is not None:
                try: tcat = te.Category.Name if te.Category else None
                except: tcat = None
                try:
                    fs = doc.GetElement(te.GetTypeId())
                    tfam = nm(fs.Family) if fs and hasattr(fs, 'Family') else None
                except: tfam = None
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

    offset = None
    try:
        if te is not None:
            L = te.Location
            if isinstance(L, DB.LocationPoint):
                hp = L.Point
                offset = [round(head.X-hp.X,4), round(head.Y-hp.Y,4),
                          round(head.Z-hp.Z,4)]
    except: offset = None

    recs.append({
        "id": tag.Id.IntegerValue,
        "tagged_id": tid,
        "tagged_category": tcat,
        "tagged_family": tfam,
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
        "offset_from_host": offset,
    })

path = write_json("fixture_tags.json", {"view":VIEW_NAME,"source_doc":doc.Title,
                   "count":len(recs),"elements":recs})
print("fixture tags in crop:", len(recs))
for r in recs[:3]:
    print(" id",r["id"],"text",repr(r["tag_text"]),
          "tagged_family",r["tagged_family"],"has_leader",r["has_leader"])
print("written:", path)
