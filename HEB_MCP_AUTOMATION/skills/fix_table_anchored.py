# SKILL: fix_table_anchored  (precise per-table correction)
# For receptacles that sit ON a table: identify the table in the SOURCE
# (nearest table-like equipment), and because that table is the SAME model
# element (shared id) in the TARGET equipment link, reproduce the device's
# EXACT offset in the table's local frame -> precise along-table position.
#   target = tgtTableLoc + Rz(tgtAng) * Rz(-srcAng) * (srcDevWorld - srcTableLoc)
# Globals: BA_FIX_APPLY (default False), BA_TGT_TITLE, BA_SRC_EQUIP,
#          BA_TGT_EQUIP, BA_FIX_IDS (default = fix_table_report ids)
# Writes: data\fix_table_anchored_report.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import math, re

APPLY    = globals().get("BA_FIX_APPLY", False)
TGT      = globals().get("BA_TGT_TITLE", "OkmntProfCorr")
SRC_EQ   = globals().get("BA_SRC_EQUIP", "Equip")
TGT_EQ   = globals().get("BA_TGT_EQUIP", "Oakmont_v24_HEB_ARCH")
TBL = re.compile(r"table|worktop|worktable|smartlever|workstation", re.I)

app = doc.Application
tgt = next(d for d in app.Documents if TGT in d.Title)
src = next(d for d in app.Documents if d.Title.startswith("RunUpdateProfiles"))

def eqlink(host, sub):
    v = get_view(host)
    for x in DB.FilteredElementCollector(host, v.Id):
        if isinstance(x, DB.RevitLinkInstance) and x.GetLinkDocument() \
           and sub in x.GetLinkDocument().Title:
            return x.GetLinkDocument(), x.GetTotalTransform()
    return None, None
sld, stf = eqlink(src, SRC_EQ)
tld, ttf = eqlink(tgt, TGT_EQ)
sbase = math.atan2(stf.BasisX.Y, stf.BasisX.X)
tbase = math.atan2(ttf.BasisX.Y, ttf.BasisX.X)

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
        out.append({"id": e.Id.IntegerValue, "fam": fam, "type": type_name(e),
                    "w": tfm.OfPoint(L.Point), "ang": ang_of(e, base)})
    return out

src_tables = table_list(sld, stf, sbase)
tgt_tables = table_list(tld, ttf, tbase)
def match_tgt_table(stbl):
    """same-family target table nearest to the source table's world pos."""
    best = None
    for tt in tgt_tables:
        if tt["fam"] != stbl["fam"]: continue
        d = math.hypot(tt["w"].X - stbl["w"].X, tt["w"].Y - stbl["w"].Y)
        if best is None or d < best[0]:
            if tt["type"] == stbl["type"] or best is None or d < best[0]:
                best = (d, tt)
    return best

host = {r["id"]: r for r in json.load(io.open(DATA + r"\host_elements.json"))["elements"]}
pm   = json.load(io.open(DATA + r"\place_relative_map.json"))["devices"]
IDS  = globals().get("BA_FIX_IDS",
        [m["id"] for m in json.load(io.open(DATA + r"\fix_table_report.json"))["moves"]])

moves = []
t = None
if APPLY:
    t = DB.Transaction(tgt, "BAKERY fix table-anchored"); t.Start()
try:
    for did in IDS:
        r = host.get(did)
        if not r: moves.append({"id": did, "status": "no-src-rec"}); continue
        dw = r["world_xyz"]
        # source table the device sits on = nearest table-like in source
        best = None
        for st in src_tables:
            d = math.hypot(dw[0]-st["w"].X, dw[1]-st["w"].Y)
            if best is None or d < best[0]: best = (d, st)
        if not best:
            moves.append({"id": did, "status": "no-src-table"}); continue
        sd, stbl = best
        fam = stbl["fam"]; sw = stbl["w"]; sang = stbl["ang"]
        mt = match_tgt_table(stbl)
        if not mt:
            moves.append({"id": did, "status": "no-matching-tgt-table",
                          "table": fam[:34]}); continue
        tdist, tt = mt
        tw = tt["w"]; tang = tt["ang"]
        dth = tang - sang
        c, s = math.cos(dth), math.sin(dth)
        ox, oy, oz = dw[0]-sw.X, dw[1]-sw.Y, dw[2]-sw.Z
        nx = tw.X + c*ox - s*oy
        ny = tw.Y + s*ox + c*oy
        e = tgt.GetElement(DB.ElementId(int(pm[str(did)])))
        p = e.Location.Point
        # desired facing = source device facing rotated by table delta
        sf = r.get("facing") or [0.0, -1.0]
        des = math.atan2(sf[1], sf[0]) + dth
        fo = e.FacingOrientation
        cur = math.atan2(fo.Y, fo.X)
        rot = (des - cur + math.pi) % (2*math.pi) - math.pi
        rec = {"id": did, "table": fam[:34], "src_table_dist": round(sd, 2),
               "dtheta_deg": round(math.degrees(dth), 1),
               "pos_from": [round(p.X, 3), round(p.Y, 3)],
               "pos_to": [round(nx, 3), round(ny, 3)],
               "move_ft": round(math.hypot(nx-p.X, ny-p.Y), 3),
               "rot_deg": round(math.degrees(rot), 1)}
        # position only -- rotation already user-confirmed correct
        if APPLY and rec["move_ft"] > 1e-4:
            DB.ElementTransformUtils.MoveElement(tgt, e.Id,
                DB.XYZ(nx-p.X, ny-p.Y, 0.0))
        rec["status"] = "ok"
        moves.append(rec)
    if APPLY: t.Commit()
except Exception as ex:
    if APPLY and t: t.RollBack()
    print("ABORTED, rolled back:", ex); raise

write_json("fix_table_anchored_report.json",
           {"target": tgt.Title, "apply": APPLY, "moves": moves})
print("=== FIX TABLE-ANCHORED %s ===" % ("APPLIED" if APPLY else "DRY-RUN"))
for m in moves:
    print("  dev %s [%s] table=%s d=%s dth=%s  %s->%s move=%s rot=%s" % (
        m["id"], m.get("status"), m.get("table"), m.get("src_table_dist"),
        m.get("dtheta_deg"), m.get("pos_from"), m.get("pos_to"),
        m.get("move_ft"), m.get("rot_deg")))
print("written: data\\fix_table_anchored_report.json")
