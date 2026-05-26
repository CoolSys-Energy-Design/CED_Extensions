# SKILL: place_relative  (equipment-RELATIVE replication; generalizes across buildings)
# For each source device, find its source anchor equipment (rel_to_equipment),
# locate the SAME family/type equipment instance in the TARGET equipment link,
# and place the device at  target_anchor_world + Rz(dtheta)*offset , with
# device rotation adjusted by the anchor's facing delta.
#
# Config via globals before exec:
#   BA_DRYRUN      (default True)  -> analysis only, no model changes
#   BA_TGT_TITLE   substring of target project title
#   BA_SRC_EQUIP   source equip link doc title substring   (default "Equip")
#   BA_TGT_EQUIP   target equip link doc title substring
#   BA_TGT_LEVEL   target level name for placement
# Writes: data\place_relative_report.json  (+ map when not dry-run)
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math

DRYRUN     = globals().get("BA_DRYRUN", True)
TGT_TITLE  = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
SRC_EQUIP  = globals().get("BA_SRC_EQUIP", "Equip")
TGT_EQUIP  = globals().get("BA_TGT_EQUIP", "Oakmont_v24_HEB_ARCH")
TGT_LEVEL  = globals().get("BA_TGT_LEVEL", "Level 1 - Sheet View")

app = doc.Application
tgt = next(d for d in app.Documents if TGT_TITLE in d.Title)
src = next(d for d in app.Documents if d.Title.startswith("RunUpdateProfiles"))

def equip_link(host, part):
    v = get_view(host)
    for li in DB.FilteredElementCollector(host, v.Id):
        if isinstance(li, DB.RevitLinkInstance):
            ld = li.GetLinkDocument()
            if ld and part in ld.Title:
                return li, ld, li.GetTotalTransform()
    return None, None, None

sli, sld, stf = equip_link(src, SRC_EQUIP)
tli, tld, ttf = equip_link(tgt, TGT_EQUIP)
if not sld or not tld:
    raise Exception("equip link missing: src=%s tgt=%s" % (bool(sld), bool(tld)))

def facing_angle(e):
    fo = getattr(e, 'FacingOrientation', None)
    if fo is None: return 0.0
    return math.atan2(fo.Y, fo.X)

ANCHOR_BIC = (DB.BuiltInCategory.OST_SpecialityEquipment,
              DB.BuiltInCategory.OST_MechanicalEquipment)

def eq_index(ld, tf):
    """element id -> {w:[x,y,z], ang, guid, fam}  for equipment instances."""
    idx = {}
    base = math.atan2(tf.BasisX.Y, tf.BasisX.X)
    for bic in ANCHOR_BIC:
        for e in DB.FilteredElementCollector(ld).OfCategory(bic).WhereElementIsNotElementType():
            L = e.Location
            if not isinstance(L, DB.LocationPoint): continue
            w = tf.OfPoint(L.Point)
            g = None
            try:
                gp = e.get_Parameter(DB.BuiltInParameter.IFC_GUID)
                g = gp.AsString() if gp else None
            except: pass
            fsym = ld.GetElement(e.GetTypeId())
            idx[e.Id.IntegerValue] = {
                "w": [w.X, w.Y, w.Z], "ang": facing_angle(e) + base, "guid": g,
                "fam": nm(fsym.Family) if fsym and hasattr(fsym, 'Family') else None}
    return idx

SIDX = eq_index(sld, stf)   # Carrollton equipment, by id
TIDX = eq_index(tld, ttf)   # Oakmont equipment, by id (same ids = same equipment)

def nearest_src(x, y):
    """Nearest source equipment whose SAME id also exists in the target
    equipment link (i.e. a transferable, shared anchor)."""
    best = None; bd = 1e18
    for eid, v in SIDX.items():
        if eid not in TIDX:          # only shared equipment can be an anchor
            continue
        d = (v["w"][0]-x)**2 + (v["w"][1]-y)**2
        if d < bd: bd = d; best = eid
    return best, math.sqrt(bd) if best is not None else None

hosts = json.load(io.open(DATA + r"\host_elements.json"))["elements"]
rows = []
n_uni = n_multi = n_missing = n_norel = 0
for r in hosts:
    if r["category"] != "Electrical Fixtures":   # devices only for the test
        continue
    dwx = r["world_xyz"]
    aid, adist = nearest_src(dwx[0], dwx[1])
    if aid is None:
        n_norel += 1
        rows.append({"id": r["id"], "fam": r["family"], "status": "no-src-equip"})
        continue
    sa = SIDX[aid]
    ta = TIDX.get(aid)
    off = [dwx[0]-sa["w"][0], dwx[1]-sa["w"][1], dwx[2]-sa["w"][2]]
    rec = {"id": r["id"], "fam": r["family"], "type": r["type"],
           "anchor_id": aid, "anchor_fam": sa["fam"],
           "anchor_dist_ft": round(adist, 2),
           "src_anchor_w": [round(c,3) for c in sa["w"]],
           "offset_ft": [round(c,3) for c in off],
           "src_dev_rot": r.get("rotation")}
    if ta is None:
        n_missing += 1; rec["status"] = "anchor-id-not-in-target"
        rows.append(rec); continue
    if sa["guid"] and ta["guid"] and sa["guid"] != ta["guid"]:
        n_missing += 1; rec["status"] = "guid-mismatch"
        rows.append(rec); continue
    n_uni += 1; rec["status"] = "matched"
    dth = ta["ang"] - sa["ang"]
    c, s = math.cos(dth), math.sin(dth)
    tx = ta["w"][0] + c*off[0] - s*off[1]
    ty = ta["w"][1] + s*off[0] + c*off[1]
    tz = ta["w"][2] + off[2]
    rec["tgt_anchor_w"] = [round(v,3) for v in ta["w"]]
    rec["dtheta_deg"] = round(math.degrees(dth), 2)
    rec["tgt_dev_xyz"] = [round(tx,3), round(ty,3), round(tz,3)]
    rec["tgt_dev_rot"] = (r.get("rotation") or 0.0) + dth
    rec["equip_shift_ft"] = round(((ta["w"][0]-sa["w"][0])**2 +
                                   (ta["w"][1]-sa["w"][1])**2) ** 0.5, 2)
    rows.append(rec)

shifts = [x["equip_shift_ft"] for x in rows if "equip_shift_ft" in x]
write_json("place_relative_report.json", {
    "target": tgt.Title, "src_equip": sld.Title, "tgt_equip": tld.Title,
    "dryrun": DRYRUN,
    "devices_total": sum(1 for r in hosts if r["category"]=="Electrical Fixtures"),
    "unique_match": n_uni, "ambiguous": n_multi,
    "anchor_missing": n_missing, "no_anchor": n_norel,
    "equip_shift_ft_min": round(min(shifts),2) if shifts else None,
    "equip_shift_ft_max": round(max(shifts),2) if shifts else None,
    "equip_shift_ft_avg": round(sum(shifts)/len(shifts),2) if shifts else None,
    "rows": rows})

if not DRYRUN:
    from System.Collections.Generic import List
    import os, re
    MAP = DATA + r"\place_relative_map.json"
    TRIG    = globals().get("BA_TRIG", 3.0)        # anchor_dist that triggers a fix
    SRC_TBL = globals().get("BA_SRC_TBL_MAX", 4.0) # device "sits on" a table if within
    SNAP_MX = globals().get("BA_SNAP_MAX", 9.0)
    TBL = re.compile(r"table|worktop|worktable|smartlever|workstation", re.I)
    pm = {}
    if os.path.exists(MAP):
        try: pm = json.load(io.open(MAP))
        except: pm = {}
    pm.setdefault("devices", {})
    symv = {}; ftypes = {}
    for fsx in DB.FilteredElementCollector(tgt).OfClass(DB.FamilySymbol):
        try:
            f = fsx.Family.Name
            tn = fsx.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        except: continue
        symv[(f, tn)] = fsx.Id
        ftypes.setdefault(f, []).append(fsx.Id)
    lvl = next(L for L in DB.FilteredElementCollector(tgt).OfClass(DB.Level)
               if L.Name == TGT_LEVEL)
    hostsById = {r["id"]: r for r in hosts}

    sba = math.atan2(stf.BasisX.Y, stf.BasisX.X)
    tba = math.atan2(ttf.BasisX.Y, ttf.BasisX.X)
    def ang_of(e, base):
        fo = getattr(e, 'FacingOrientation', None)
        return (math.atan2(fo.Y, fo.X) + base) if fo else base
    def table_list(ld, tfm, base):
        out = []
        for e in DB.FilteredElementCollector(ld).OfCategory(
                DB.BuiltInCategory.OST_SpecialityEquipment).WhereElementIsNotElementType():
            L = e.Location
            if not isinstance(L, DB.LocationPoint): continue
            fs = ld.GetElement(e.GetTypeId())
            fam = nm(fs.Family) if fs and hasattr(fs, 'Family') else ""
            if not fam or not TBL.search(fam): continue
            out.append({"fam": fam, "type": type_name(e),
                        "w": tfm.OfPoint(L.Point), "ang": ang_of(e, base)})
        return out
    src_tables = table_list(sld, stf, sba)
    tgt_tables = table_list(tld, ttf, tba)

    bvw = get_view(tgt); inc = crop_test(bvw, pad=10)
    walls = []
    for w in DB.FilteredElementCollector(tld).OfClass(DB.Wall):
        try:
            c = w.Location.Curve
            a = ttf.OfPoint(c.GetEndPoint(0)); b = ttf.OfPoint(c.GetEndPoint(1))
            if inc(a) or inc(b): walls.append((a, b, (w.Width or 0.5)/2.0))
        except: pass
    boxes = []
    for bic in (DB.BuiltInCategory.OST_SpecialityEquipment,
                DB.BuiltInCategory.OST_MechanicalEquipment):
        for e in DB.FilteredElementCollector(tld).OfCategory(bic).WhereElementIsNotElementType():
            try:
                bb = e.get_BoundingBox(None)
                if bb is None: continue
                mn = ttf.OfPoint(bb.Min); mx = ttf.OfPoint(bb.Max)
                cx, cy = (mn.X+mx.X)/2, (mn.Y+mx.Y)/2
                if not inc(DB.XYZ(cx, cy, mn.Z)): continue
                boxes.append((min(mn.X,mx.X), min(mn.Y,mx.Y),
                              max(mn.X,mx.X), max(mn.Y,mx.Y)))
            except: pass

    def table_pose(r):
        """exact on-table position via same-family table local frame, or None."""
        dw = r["world_xyz"]
        best = None
        for st in src_tables:
            d = math.hypot(dw[0]-st["w"].X, dw[1]-st["w"].Y)
            if best is None or d < best[0]: best = (d, st)
        if not best or best[0] > SRC_TBL: return None
        st = best[1]
        cand = None
        for tt in tgt_tables:
            if tt["fam"] != st["fam"]: continue
            d = math.hypot(tt["w"].X-st["w"].X, tt["w"].Y-st["w"].Y)
            if cand is None or d < cand[0]: cand = (d, tt)
        if not cand: return None
        tt = cand[1]; dth = tt["ang"] - st["ang"]
        c, s = math.cos(dth), math.sin(dth)
        ox, oy, oz = dw[0]-st["w"].X, dw[1]-st["w"].Y, dw[2]-st["w"].Z
        return (tt["w"].X + c*ox - s*oy, tt["w"].Y + s*ox + c*oy,
                tt["w"].Z + oz, dth, "table:" + st["fam"][:24])

    def snap_pose(px, py):
        """nearest wall/equip face point + outward normal, or None."""
        best = None
        for (a, b, hw) in walls:
            abx, aby = b.X-a.X, b.Y-a.Y
            L2 = abx*abx + aby*aby
            if L2 < 1e-9: continue
            tt = max(0.0, min(1.0, ((px-a.X)*abx + (py-a.Y)*aby)/L2))
            cx, cy = a.X+tt*abx, a.Y+tt*aby
            d = math.hypot(px-cx, py-cy)
            if d < 1e-6:
                nx, ny = -aby, abx; nl = math.hypot(nx, ny) or 1.0
                nx, ny = nx/nl, ny/nl
            else:
                nx, ny = (px-cx)/d, (py-cy)/d
            fd = abs(d-hw)
            if best is None or fd < best[0]:
                best = (fd, cx+nx*hw, cy+ny*hw, nx, ny)
        for (x0, y0, x1, y1) in boxes:
            inside = (x0 <= px <= x1) and (y0 <= py <= y1)
            cs = [(abs(px-x0), x0, max(y0,min(py,y1)), -1.0, 0.0),
                  (abs(px-x1), x1, max(y0,min(py,y1)),  1.0, 0.0),
                  (abs(py-y0), max(x0,min(px,x1)), y0, 0.0, -1.0),
                  (abs(py-y1), max(x0,min(px,x1)), y1, 0.0,  1.0)]
            fd, fx, fy, nx, ny = min(cs, key=lambda z: z[0])
            if not inside: fd = math.hypot(px-fx, py-fy)
            if best is None or fd < best[0]:
                best = (fd, fx, fy, nx, ny)
        return best

    # ---- writable-param application policy (user-confirmed) ----
    # Skip: identity/unique, worksharing/env, placement geometry, the CED
    # cross-ref linker, AND all circuit-derived params (CKT_*, computed
    # circuit metrics, conduit/wire sizing) -> those are recomputed by the
    # circuiting step. Everything else writable (loads, symbol/visibility
    # config) is applied so placed devices carry appropriate parameters.
    WP_SKIP_EXACT = set([
        "IfcGUID", "Workset", "Export to IFC", "Edited by", "Mark",
        "Comments", "Element_Linker",
        "Offset from Host", "Elevation from Level",
        "ACA Z Rotation", "ACA Y", "ACA Z",
        "Circuit Ampacity_CED", "Circuit Load Current_CED",
        "Conduit Fill Percentage_CED", "Voltage Drop Percentage_CED",
        "Conduit and Wire Size_CEDT", "Conduit Size_CEDT", "Conduit Type_CEDT",
        "Wire Material_CEDT", "Wire Insulation_CEDT",
        "Wire Temparature Rating_CEDT", "Wire Size_CEDT",
        # view/project DISPLAY toggles -- depend on the host project's view
        # setup, NOT portable. Copying source "Visible in Floor Plan"=0 hid
        # the receptacle symbols in the FloorPlan-type callout view.
        "Visible in Floor Plan", "Visible in Callout", "Visible in Section",
        "Visible in Elevation", "1/4\" Plan",
        "Appears in Schedule", "Schedule Sort Order",
    ])
    def apply_wparams(fi, rec):
        n_ok = 0
        for w in (rec.get("wparams") or []):
            nm_ = w.get("n")
            if (not nm_ or nm_ in WP_SKIP_EXACT or nm_.startswith("CKT_")
                    or nm_.startswith("Visible in ")):
                continue
            try:
                pr = fi.LookupParameter(nm_)
                if pr is None or pr.IsReadOnly: continue
                st = w.get("st", ""); v = w.get("v")
                if v is None: continue
                if st.endswith("String"):
                    if pr.StorageType == DB.StorageType.String: pr.Set(str(v)); n_ok += 1
                elif st.endswith("Integer"):
                    if pr.StorageType == DB.StorageType.Integer: pr.Set(int(v)); n_ok += 1
                elif st.endswith("Double"):
                    if pr.StorageType == DB.StorageType.Double: pr.Set(float(v)); n_ok += 1
            except: pass
        return n_ok

    placed = skipped = dups = n_tbl = n_snap = 0
    wp_total = 0
    tg = DB.TransactionGroup(tgt, "BAKERY Phase A (equip-relative + table/wall fix)")
    tg.Start()
    t = DB.Transaction(tgt, "place + correct")
    t.Start()
    try:
        for x in rows:
            if x["status"] != "matched": skipped += 1; continue
            sid = str(x["id"])
            if sid in pm["devices"]: skipped += 1; continue
            r = hostsById[x["id"]]
            fam, typ = r["family"], r["type"]
            if (fam, typ) not in symv:
                sibs = ftypes.get(fam)
                if not sibs: skipped += 1; continue
                ns = tgt.GetElement(sibs[0]).Duplicate(typ)
                symv[(fam, typ)] = ns.Id; ftypes[fam].append(ns.Id); dups += 1
            fsx = tgt.GetElement(symv[(fam, typ)])
            if not fsx.IsActive: fsx.Activate()

            method = "equip"; reaim = None
            pos = list(x["tgt_dev_xyz"]); rr = x.get("tgt_dev_rot") or 0.0
            if x.get("anchor_dist_ft", 0) > TRIG:
                tp = table_pose(r)
                if tp:
                    pos = [tp[0], tp[1], tp[2]]
                    rr = (r.get("rotation") or 0.0) + tp[3]
                    method = tp[4]; n_tbl += 1
                else:
                    sp = snap_pose(pos[0], pos[1])
                    if sp and sp[0] <= SNAP_MX:
                        pos[0], pos[1] = sp[1], sp[2]
                        reaim = math.atan2(sp[4], sp[3])
                        method = "snap"; n_snap += 1

            # Re-base Z to TARGET level: preserve mounting height AFF, not
            # absolute Z. Source/target Level-1 elevations differ (e.g.
            # source=0, Oakmont=100); equipment-relative XY is independent of
            # this, but Z must follow the TARGET level so devices land within
            # the plan view's view range and render correctly.
            src_aff = r["world_xyz"][2] - float(r.get("level_elev", 0.0) or 0.0)
            p = DB.XYZ(pos[0], pos[1], lvl.Elevation + src_aff)
            try:
                fi = tgt.Create.NewFamilyInstance(
                        p, fsx, lvl, DB.Structure.StructuralType.NonStructural)
            except Exception as e:
                x["place_error"] = str(e); skipped += 1; continue
            if rr:
                ax = DB.Line.CreateBound(p, DB.XYZ(p.X, p.Y, p.Z+1))
                try: DB.ElementTransformUtils.RotateElement(tgt, fi.Id, ax, rr)
                except: pass
            if reaim is not None:
                try:
                    fo = fi.FacingOrientation
                    cur = math.atan2(fo.Y, fo.X)
                    dd = (reaim - cur + math.pi) % (2*math.pi) - math.pi
                    if abs(dd) > 0.05:
                        q = fi.Location.Point
                        ax = DB.Line.CreateBound(q, DB.XYZ(q.X, q.Y, q.Z+1))
                        DB.ElementTransformUtils.RotateElement(tgt, fi.Id, ax, dd)
                except: pass
            wp_total += apply_wparams(fi, r)
            try:
                cp = fi.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                if cp and not cp.IsReadOnly:
                    cp.Set("[BA:src%s][%s]" % (sid, method))
            except: pass
            pm["devices"][sid] = fi.Id.IntegerValue
            placed += 1
        t.Commit()
        tg.Assimilate()
    except Exception as e:
        t.RollBack()
        if tg.GetStatus() == DB.TransactionStatus.Started: tg.RollBack()
        print("PHASE A ABORTED, rolled back:", e); raise
    with io.open(MAP, 'w', encoding='utf-8') as _f:
        _f.write(json.dumps(pm, indent=1))
    print(">>> PHASE A placed=%d (equip + %d table-anchored + %d wall-snap) "
          "skipped=%d dup_types=%d wparams_set=%d"
          % (placed, n_tbl, n_snap, skipped, dups, wp_total))
    print("    single undo step; map=%s" % MAP)

print("=== EQUIPMENT-RELATIVE MATCH ANALYSIS (dryrun=%s) ===" % DRYRUN)
print("target:", tgt.Title)
print("src equip link:", sld.Title, "| tgt equip link:", tld.Title)
print("devices(fixtures):", sum(1 for r in hosts if r["category"]=="Electrical Fixtures"))
print("  unique anchor match : %d" % n_uni)
print("  ambiguous (multi)   : %d" % n_multi)
print("  anchor missing      : %d" % n_missing)
print("  device had no anchor : %d" % n_norel)
if shifts:
    print("equipment shift Carrollton->Oakmont (ft): min=%.2f avg=%.2f max=%.2f"
          % (min(shifts), sum(shifts)/len(shifts), max(shifts)))
print("written: data\\place_relative_report.json")
