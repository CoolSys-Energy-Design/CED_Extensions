# Generic convenience-recep placer.
# Reads block_map.json from the same folder, places one Duplex Wall recep per
# entry in convenience_recep_plan.receps, with rotation per wall.
# Usage: open with execute_revit_code on the Charlotte Central E101 view.

import math, json, clr
from Autodesk.Revit.DB import UnitUtils
clr.AddReference("System")
clr.AddReference("System.IO")
from System.IO import File

# CHANGE THIS PATH for each room copy:
ROOM_DIR = r"c:\CED_Extensions\Planet Fitness_MCP_AUTOMATION\Electrical Power_auto\rooms\110_STRENGTH"

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

placed = []
t = DB.Transaction(doc, "Place convenience receps for room " + bm["room_number"])
t.Start()
try:
    for r in bm["convenience_recep_plan"]["receps"]:
        pt = DB.XYZ(r["x"], r["y"], level.Elevation)
        inst = doc.Create.NewFamilyInstance(pt, sym, level, DB.Structure.StructuralType.NonStructural)
        try:
            h = inst.LookupParameter("Elevation from Level")
            if h and not h.IsReadOnly: h.Set(float(bm["mount_height_in"])/12.0)
        except: pass
        if r["rot"] != 0:
            axis = DB.Line.CreateBound(pt, DB.XYZ(pt.X, pt.Y, pt.Z + 1))
            inst.Location.Rotate(axis, float(r["rot"]))
        STR = dict(bm["profile_str"])
        STR["CKT_Load Name_CEDT"] = "RECEPTACLE - " + bm["room_number"] + " " + bm["room_name"]
        for k, v in STR.items():
            p = inst.LookupParameter(k)
            if p and not p.IsReadOnly: p.Set(str(v))
        for k, v in bm["profile_dbl_disp"].items():
            p = inst.LookupParameter(k)
            if p and not p.IsReadOnly:
                iv = to_internal(inst, k, v)
                if iv is not None: p.Set(iv)
        placed.append({"id": inst.Id.IntegerValue, "x": r["x"], "y": r["y"], "wall": r["wall"]})
    t.Commit()
    print("Placed %d convenience receps in %s %s" % (len(placed), bm["room_number"], bm["room_name"]))
except Exception as e:
    t.RollBack()
    print("FAILED: %s" % e)
    raise

File.WriteAllText(ROOM_DIR + r"\manifest.json",
    json.dumps({"convenience_receps": placed, "bbox": bm["convenience_recep_plan"]["bbox"],
                "n_target": len(bm["convenience_recep_plan"]["receps"])}, indent=2))
