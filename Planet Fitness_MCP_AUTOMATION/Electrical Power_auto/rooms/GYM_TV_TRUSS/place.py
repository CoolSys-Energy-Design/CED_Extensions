# Re-execute placement for GYM TV TRUSS.
# Places 20 Quad Wall - TV receps (10 anchor positions x 2 receps each), facing south.

import math, json, clr
from Autodesk.Revit.DB import UnitUtils
clr.AddReference("System")
clr.AddReference("System.IO")
from System.IO import File

ROOT = r"c:\CED_Extensions\Planet Fitness_MCP_AUTOMATION\Electrical Power_auto"
ROOM_DIR = ROOT + r"\rooms\GYM_TV_TRUSS"

_NAME = clr.GetClrType(DB.Element).GetProperty('Name')
def el_name(el):
    try: return _NAME.GetValue(el, None) or ""
    except: return ""

def find_sym(fam, typ):
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements():
        if s.Family.Name == fam and el_name(s) == typ: return s
    return None

def to_internal(elem, pname, val):
    p = elem.LookupParameter(pname)
    if p is None: return None
    spec = p.Definition.GetDataType()
    du = doc.GetUnits().GetFormatOptions(spec).GetUnitTypeId()
    return UnitUtils.ConvertToInternalUnits(float(val), du)

bm = json.loads(File.ReadAllText(ROOM_DIR + r"\block_map.json"))
sym = find_sym(bm["fixture_family"], bm["fixture_type"])
if not sym.IsActive:
    t = DB.Transaction(doc, "act"); t.Start(); sym.Activate(); doc.Regenerate(); t.Commit()

lvls = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
level = next(L for L in lvls if "L1 - Finished Floor" in el_name(L))
Y = float(bm["anchor_y"])
DX = float(bm["x_offset_in_per_pair"]) / 12.0
rot = math.pi if bm.get("rotation_rad_pi") else 0.0

placed = []
t = DB.Transaction(doc, "Place GYM TV TRUSS")
t.Start()
try:
    for ax in bm["anchor_xs"]:
        for sign, gid in zip([+1, -1], bm["pair_dev_group_ids"]):
            pt = DB.XYZ(ax + sign*DX, Y, level.Elevation)
            inst = doc.Create.NewFamilyInstance(pt, sym, level, DB.Structure.StructuralType.NonStructural)
            try:
                h = inst.LookupParameter("Elevation from Level")
                if h and not h.IsReadOnly: h.Set(float(bm["mount_height_in"])/12.0)
            except: pass
            if rot != 0:
                axis = DB.Line.CreateBound(pt, DB.XYZ(pt.X, pt.Y, pt.Z + 1))
                inst.Location.Rotate(axis, rot)
            params = dict(bm["params"]); params["dev-Group ID"] = gid
            for k, v in params.items():
                p = inst.LookupParameter(k)
                if p is None or p.IsReadOnly: continue
                st = p.StorageType
                if st == DB.StorageType.String: p.Set(str(v))
                elif st == DB.StorageType.Integer: p.Set(int(v))
                elif st == DB.StorageType.Double:
                    iv = to_internal(inst, k, v)
                    if iv is not None: p.Set(iv)
            placed.append({"id": inst.Id.IntegerValue, "x": pt.X, "y": pt.Y, "dev_group": gid})
    t.Commit()
    print("Placed: %d TV TRUSS receps" % len(placed))
except Exception as e:
    t.RollBack()
    print("FAILED: %s" % e)
    raise

File.WriteAllText(ROOM_DIR + r"\manifest.json", json.dumps({"placed": placed}, indent=2))
