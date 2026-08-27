# =============================================================================
# PF E101 INJECTOR (IronPython, runs inside Revit via MCP execute_revit_code)
# Reads manifest.json (produced by the reconstruction engine) and rebuilds the
# E101 power plan in the ACTIVE document:
#   pass 1  distribution equipment (panels / transformers / switchboard)
#   pass 2  devices (receptacles, disconnects, J-boxes, switches)
#   pass 3  circuits (ElectricalSystem per manifest, panel assigned, load named)
#   pass 4  keynote symbols on the active view
#   pass 5  panel/circuit tags on every circuited device
# The active view should be the E101 power plan; the CAD background must be the
# same X_BG placement the manifest was generated against (same origin).
# Re-runnable: pass a marker prefix in MARK to find/delete a previous injection.
# =============================================================================
import json, math, traceback
from Autodesk.Revit.DB import (FilteredElementCollector, FamilySymbol, Element,
                               Level, XYZ, Line, Transaction, ElementTransformUtils,
                               BuiltInParameter, IndependentTag, Reference,
                               TagOrientation, ElementId, View)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.DB.Electrical import ElectricalSystem, ElectricalSystemType
from System.Collections.Generic import List

MANIFEST = r"__MANIFEST_PATH__"
LEVEL_NAME = "L1 - Finished Floor"
MARK = "PFAI"          # stamped into Comments for cleanup/re-run
U = 10.7639104         # VA/V real -> Revit internal

result = {"ok": False, "stage": "load"}
try:
    man = json.load(open(MANIFEST))
    tdoc = doc
    view = uidoc.ActiveView

    # ---------- lookups ----------
    def ename(x):
        try:
            return Element.Name.__get__(x)
        except:
            return None
    symbols = {}
    for fs in FilteredElementCollector(tdoc).OfClass(FamilySymbol).ToElements():
        try:
            symbols[(fs.Family.Name, ename(fs))] = fs
        except:
            pass
    level = None
    for lv in FilteredElementCollector(tdoc).OfClass(Level).ToElements():
        if ename(lv) == LEVEL_NAME:
            level = lv
    if level is None:
        raise Exception("Level not found: " + LEVEL_NAME)

    def set_param(el, name, val, as_double=False):
        try:
            p = el.LookupParameter(name)
            if p and not p.IsReadOnly:
                if as_double:
                    p.Set(float(val))
                else:
                    p.Set(val)
                return True
        except:
            pass
        return False

    def place(fam, typ, x, y, rot_deg, elev):
        sym = symbols.get((fam, typ))
        if sym is None:
            return None, "missing symbol %s::%s" % (fam, typ)
        if not sym.IsActive:
            sym.Activate()
        pt = XYZ(x, y, 0.0)
        inst = tdoc.Create.NewFamilyInstance(pt, sym, level, StructuralType.NonStructural)
        if rot_deg:
            axis = Line.CreateBound(pt, XYZ(x, y, 10.0))
            ElementTransformUtils.RotateElement(tdoc, inst.Id, axis, math.radians(rot_deg))
        if elev:
            set_param(inst, "Elevation from Level", float(elev), as_double=True)
            try:
                bp = inst.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                if bp and not bp.IsReadOnly:
                    bp.Set(float(elev))
            except:
                pass
        set_param(inst, "Comments", MARK)
        return inst, None

    report = {"panels": 0, "devices": 0, "circuits": 0, "keynotes": 0, "tags": 0, "errors": []}

    # ---------- pass 1: distribution equipment ----------
    result["stage"] = "panels"
    t = Transaction(tdoc, "PFAI inject panels")
    t.Start()
    panel_inst = {}
    for p in man["panels"]:
        inst, err = place(p["fam"], p["typ"], p["x"], p["y"], 0, 0)
        if err:
            report["errors"].append(err)
            continue
        set_param(inst, "Panel Name", p["name"])
        panel_inst[p["name"]] = inst
        report["panels"] += 1
    t.Commit()

    # ---------- pass 2: devices ----------
    result["stage"] = "devices"
    t = Transaction(tdoc, "PFAI inject devices")
    t.Start()
    dev_inst = []
    for d in man["devices"]:
        inst, err = place(d["family"], d["type"], d["x"], d["y"], d["rotation_deg"],
                          d["elevation_from_level"])
        if err:
            report["errors"].append(err)
            dev_inst.append(None)
            continue
        if d.get("va"):
            set_param(inst, "Apparent Load Input_CED", d["va"] * U, as_double=True)
        if d.get("load_name"):
            set_param(inst, "CKT_Load Name_CEDT", d["load_name"])
        dev_inst.append(inst)
        report["devices"] += 1
    t.Commit()

    # ---------- pass 3: circuits ----------
    # Create in slot order so Revit's sequential slot assignment approximates the
    # intended numbering; exact slots can be arranged afterwards in the panel
    # schedule (or via PanelScheduleView.MoveSlotTo).
    result["stage"] = "circuits"
    def slot_key(c):
        try:
            return (c["panel"], int(str(c["circuit"]).split(",")[0]))
        except:
            return (c["panel"], 999)
    t = Transaction(tdoc, "PFAI inject circuits")
    t.Start()
    for c in sorted(man["circuits"], key=slot_key):
        mem = [dev_inst[i] for i in c.get("member_device_indices", [])
               if i < len(dev_inst) and dev_inst[i] is not None]
        if not mem:
            continue
        try:
            ids = List[ElementId]([m.Id for m in mem])
            sys = ElectricalSystem.Create(tdoc, ids, ElectricalSystemType.PowerCircuit)
            pn = panel_inst.get(c["panel"])
            if pn is not None:
                try:
                    sys.SelectPanel(pn)
                except:
                    report["errors"].append("panel assign failed %s/%s" % (c["panel"], c["circuit"]))
            try:
                sys.LoadName = c["load_name"]
            except:
                p2 = sys.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
                if p2:
                    p2.Set(c["load_name"])
            if c.get("rating_a"):
                try:
                    sys.Rating = float(c["rating_a"])
                except:
                    pass
            report["circuits"] += 1
        except Exception as ce:
            report["errors"].append("circuit %s/%s: %s" % (c["panel"], c["circuit"], str(ce)[:80]))
    t.Commit()

    # ---------- pass 4: keynotes on the active view ----------
    result["stage"] = "keynotes"
    kn_sym = symbols.get(("Manual Key Note- All Shapes", "Square"))
    t = Transaction(tdoc, "PFAI inject keynotes")
    t.Start()
    if kn_sym is not None:
        if not kn_sym.IsActive:
            kn_sym.Activate()
        for k in man["keynotes"]:
            if not k.get("num"):
                continue
            try:
                inst = tdoc.Create.NewFamilyInstance(XYZ(k["x"], k["y"], 0.0), kn_sym, view)
                set_param(inst, "CED-G-NOTE #", str(k["num"]))
                report["keynotes"] += 1
            except Exception as ke:
                report["errors"].append("keynote: " + str(ke)[:60])
    t.Commit()

    # ---------- pass 5: panel/circuit tags ----------
    result["stage"] = "tags"
    tag_type_id = None
    for (fam, typ), fs in symbols.items():
        if fam == "EF-Tag_Electrical Fixtures_CED" and typ == "Panel & Circuit Number":
            tag_type_id = fs.Id
    t = Transaction(tdoc, "PFAI inject tags")
    t.Start()
    if tag_type_id is not None:
        for d, inst in zip(man["devices"], dev_inst):
            if inst is None or not d.get("panel"):
                continue
            try:
                head = XYZ(d["x"] - 2.5, d["y"], 0.0)
                IndependentTag.Create(tdoc, tag_type_id, view.Id, Reference(inst),
                                      False, TagOrientation.Horizontal, head)
                report["tags"] += 1
            except Exception as te:
                report["errors"].append("tag: " + str(te)[:60])
    t.Commit()

    result = {"ok": True, "report": report, "errors_n": len(report["errors"]),
              "errors_head": report["errors"][:8]}
except Exception as e:
    result["err"] = str(e)
    result["tb"] = traceback.format_exc()[-900:]
print(json.dumps(result))
