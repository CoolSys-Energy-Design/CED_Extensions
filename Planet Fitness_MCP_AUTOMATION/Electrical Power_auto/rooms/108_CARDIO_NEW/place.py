# Re-execute placement for 108 CARDIO NEW.
# Reads block_map.json + the global blocks-with-names dump, places everything in one transaction.
# Usage: feed this script content to Revit MCP execute_revit_code on the Charlotte Central E101 view.

import math, json, clr
from Autodesk.Revit.DB import UnitUtils
clr.AddReference("System")
clr.AddReference("System.IO")
from System.IO import File

ROOT = r"c:\CED_Extensions\Planet Fitness_MCP_AUTOMATION\Electrical Power_auto"
ROOM_DIR = ROOT + r"\rooms\108_CARDIO_NEW"

_NAME = clr.GetClrType(DB.Element).GetProperty('Name')
def el_name(el):
    try: return _NAME.GetValue(el, None) or ""
    except: return ""

def find_sym(fam, typ):
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements():
        if s.Family.Name == fam and el_name(s) == typ: return s
    return None

def activate(sym):
    if sym and not sym.IsActive:
        t = DB.Transaction(doc, "act"); t.Start(); sym.Activate(); doc.Regenerate(); t.Commit()

def to_internal(elem, pname, val):
    p = elem.LookupParameter(pname)
    if p is None: return None
    spec = p.Definition.GetDataType()
    du = doc.GetUnits().GetFormatOptions(spec).GetUnitTypeId()
    return UnitUtils.ConvertToInternalUnits(float(val), du)

def apply_params(inst, params):
    for k, v in params.items():
        p = inst.LookupParameter(k)
        if p is None or p.IsReadOnly: continue
        st = p.StorageType
        if st == DB.StorageType.String:
            p.Set(str(v))
        elif st == DB.StorageType.Integer:
            p.Set(int(v))
        elif st == DB.StorageType.Double:
            iv = to_internal(inst, k, v)
            if iv is not None: p.Set(iv)

bm = json.loads(File.ReadAllText(ROOM_DIR + r"\block_map.json"))
blocks_all = json.loads(File.ReadAllText(ROOT + r"\charlotte_central_blocks_v3.json"))["blocks"]
lvls = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
level = next(L for L in lvls if "L1 - Finished Floor" in el_name(L))

placed_fixtures = []
placed_keynotes = []

t = DB.Transaction(doc, "Place 108 CARDIO NEW")
t.Start()
try:
    # Fixtures from CAD blocks
    for entry in bm["blocks"]:
        sym = find_sym(entry["fixture_family"], entry["fixture_type"])
        activate(sym)
        cads = [b for b in blocks_all if b["name"] == entry["cad_name"]]
        for bk in cads:
            pt = DB.XYZ(bk["x"], bk["y"], level.Elevation)
            inst = doc.Create.NewFamilyInstance(pt, sym, level, DB.Structure.StructuralType.NonStructural)
            try:
                h = inst.LookupParameter("Elevation from Level")
                if h and not h.IsReadOnly: h.Set(float(entry["mount_height_in"])/12.0)
            except: pass
            if entry.get("rotation_rad", 0) != 0:
                axis = DB.Line.CreateBound(pt, DB.XYZ(pt.X, pt.Y, pt.Z + 1))
                inst.Location.Rotate(axis, float(entry["rotation_rad"]))
            apply_params(inst, entry["params"])
            placed_fixtures.append({"id": inst.Id.IntegerValue, "cad_name": bk["name"], "x": bk["x"], "y": bk["y"]})

    # Keynotes
    for kn in bm.get("keynotes", []):
        sym = find_sym(kn["type_family"], kn["type_name"]); activate(sym)
        for pos in kn["positions"]:
            inst = doc.Create.NewFamilyInstance(DB.XYZ(pos["x"], pos["y"], 0), sym, doc.ActiveView)
            try:
                p = inst.LookupParameter("CED-G-NOTE #")
                if p and not p.IsReadOnly: p.Set(str(kn["ced_g_note_number"]))
            except: pass
            apply_params(inst, kn["params"])
            placed_keynotes.append({"id": inst.Id.IntegerValue, "x": pos["x"], "y": pos["y"]})

    t.Commit()
    print("Placed: %d fixtures, %d keynotes" % (len(placed_fixtures), len(placed_keynotes)))
except Exception as e:
    t.RollBack()
    print("FAILED: %s" % e)
    raise

File.WriteAllText(ROOM_DIR + r"\manifest.json",
    json.dumps({"fixtures": placed_fixtures, "keynotes": placed_keynotes}, indent=2))
