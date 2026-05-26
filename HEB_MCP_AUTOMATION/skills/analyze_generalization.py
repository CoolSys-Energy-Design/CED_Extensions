# SKILL: analyze_generalization  (READ-ONLY, no model changes)
# Tests equipment-relative placement vs the rigid equipment-link transform.
# For each source device: rigid_pos = T(device).  equiprel_pos = matched
# TARGET equipment world + recorded device->equipment offset.  Where the
# anchor equipment MOVED between the two equipment links, equiprel tracks it
# and rigid does not -> quantifies generalization benefit.
# Writes: data\generalization_report.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())

app = doc.Application
tgt = next(d for d in app.Documents if d.Title == "CED HEB Test Run_MEPR_R24")
src = next(d for d in app.Documents if d.Title.startswith("RunUpdateProfiles"))
ANCHOR_CATS = ("Specialty Equipment","Mechanical Equipment","Generic Models",
               "Plumbing Fixtures","Lighting Fixtures","Casework",
               "Food Service Equipment","Electrical Equipment")

def equip_link(d, part):
    v = get_view(d)
    for li in DB.FilteredElementCollector(d, v.Id):
        if isinstance(li, DB.RevitLinkInstance):
            ld = li.GetLinkDocument()
            if ld and part in ld.Title:
                return li.GetTotalTransform(), ld
    return None, None

stf, sld = equip_link(src, "Equip")
ttf, tld = equip_link(tgt, "Carrollton_v24")
T = ttf.Multiply(stf.Inverse)   # rigid source->target

def anchors_world(ld, tf):
    """family/type -> list of world XY (only ANCHOR cats, point-located)."""
    by = {}
    for e in DB.FilteredElementCollector(ld).WhereElementIsNotElementType():
        try: cat = e.Category.Name if e.Category else None
        except: cat = None
        if cat not in ANCHOR_CATS: continue
        L = e.Location
        if not isinstance(L, DB.LocationPoint): continue
        fsym = ld.GetElement(e.GetTypeId())
        fam = nm(fsym.Family) if fsym and hasattr(fsym, 'Family') else None
        typ = type_name(e)
        w = tf.OfPoint(L.Point)
        by.setdefault((fam, typ), []).append((round(w.X, 4), round(w.Y, 4)))
    return by

S = anchors_world(sld, stf)
Tg = anchors_world(tld, ttf)
# clean 1:1 matchable equipment (unique family/type in both)
matchable = set(k for k in S if k in Tg and len(S[k]) == 1 and len(Tg[k]) == 1)

hosts = json.load(io.open(DATA + r"\host_elements.json"))["elements"]
rows = []
moved = same = unmatched = 0
for r in hosts:
    rel = r.get("rel_to_equipment")
    if not rel: unmatched += 1; continue
    key = (rel.get("anchor_family"), rel.get("anchor_type"))
    dxy = rel["dxy_ft"]
    dw = r["world_xyz"]
    rp = T.OfPoint(DB.XYZ(dw[0], dw[1], dw[2]))          # rigid target pos
    rigid = (round(rp.X, 4), round(rp.Y, 4))
    if key not in matchable:
        unmatched += 1
        rows.append({"id": r["id"], "family": r["family"], "anchor": key,
                     "status": "anchor-not-1to1", "rigid": rigid})
        continue
    sxy = S[key][0]; txy = Tg[key][0]                    # src/tgt anchor world
    # how far this anchor equipment moved (target vs rigid-mapped source)
    sx_rig = T.OfPoint(DB.XYZ(sxy[0], sxy[1], 0))
    equip_move = ((txy[0]-sx_rig.X)**2 + (txy[1]-sx_rig.Y)**2) ** 0.5
    equiprel = (round(txy[0]+dxy[0], 4), round(txy[1]+dxy[1], 4))
    delta = ((equiprel[0]-rigid[0])**2 + (equiprel[1]-rigid[1])**2) ** 0.5
    st = "EQUIP-MOVED" if equip_move > 0.25 else "stationary"
    if equip_move > 0.25: moved += 1
    else: same += 1
    rows.append({"id": r["id"], "family": r["family"], "anchor": key,
                 "status": st, "equip_move_ft": round(equip_move, 3),
                 "rigid": rigid, "equiprel": equiprel,
                 "rigid_vs_equiprel_ft": round(delta, 3)})

rows.sort(key=lambda z: -z.get("equip_move_ft", 0))
write_json("generalization_report.json", {
    "matchable_equipment_families": len(matchable),
    "devices_total": len(hosts),
    "devices_on_moved_equipment": moved,
    "devices_on_stationary_equipment": same,
    "devices_unmatched_anchor": unmatched,
    "rows": rows})

print("matchable equip families:", len(matchable))
print("devices: moved-equip=%d stationary=%d unmatched=%d" % (moved, same, unmatched))
print("-- devices whose anchor equipment MOVED (equiprel follows it, rigid misses by ~equip_move) --")
for z in rows:
    if z["status"] == "EQUIP-MOVED":
        print("  id %d %-30s move=%.2fft  rigid_vs_equiprel=%.2fft"
              % (z["id"], z["family"], z["equip_move_ft"], z["rigid_vs_equiprel_ft"]))
print("written: data\\generalization_report.json")
