# PF2 Training extraction template (IronPython, runs inside Revit via MCP execute_revit_code)
# READ-ONLY: no transactions anywhere in this script. Operates on a BACKGROUND document
# resolved by TITLE_KEY - never touches the active (template) document.
# Substitute __TITLE_KEY__ and __OUT_DIR__ per project before sending.
import json, os, math, traceback
from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, ViewSheet,
                               ViewType, XYZ, ElementId, FamilyInstance, TextNote,
                               ImportInstance, RevitLinkInstance, View, Element,
                               BuiltInParameter, CurveElement)
from Autodesk.Revit.DB.Electrical import ElectricalSystem

OUT_DIR = r"__OUT_DIR__"
TITLE_KEY = "__TITLE_KEY__"
if not os.path.exists(OUT_DIR):
    os.makedirs(OUT_DIR)
result = {"ok": False, "stage": "start"}

def pt(p):
    return [round(p.X, 3), round(p.Y, 3), round(p.Z, 3)] if p else None

def ename(x):
    try:
        return Element.Name.__get__(x)
    except:
        try:
            return x.Name
        except:
            return None

def params_of(el):
    out = {}
    for p in el.Parameters:
        try:
            n = p.Definition.Name
            st = p.StorageType.ToString()
            if st == "String":
                v = p.AsString()
            elif st == "Double":
                v = round(p.AsDouble(), 4)
            elif st == "Integer":
                v = p.AsInteger()
            elif st == "ElementId":
                v = p.AsValueString()
            else:
                v = None
            if v not in (None, ""):
                out[n] = v
        except:
            pass
    return out

try:
    app = uidoc.Application.Application
    tdoc = None
    for d2 in app.Documents:
        if (not d2.IsLinked) and TITLE_KEY.lower() in d2.Title.lower():
            tdoc = d2
    if tdoc is None:
        raise Exception("No open doc matching " + TITLE_KEY)

    data = {"doc_title": tdoc.Title}
    pi = tdoc.ProjectInformation
    data["project"] = params_of(pi)

    # ---- find E101 sheet + its power plan view ----
    result["stage"] = "sheet"
    sheets = FilteredElementCollector(tdoc).OfClass(ViewSheet).ToElements()
    e101 = None
    for s in sheets:
        num = s.SheetNumber.upper().replace(" ", "").replace("-", "").replace(".", "")
        if num == "E101":
            e101 = s; break
    if e101 is None:
        for s in sheets:
            if "POWER PLAN" in (ename(s) or "").upper():
                e101 = s; break
    plan_view = None
    sheet_info = None
    if e101 is not None:
        sheet_info = {"number": e101.SheetNumber, "name": ename(e101),
                      "params": params_of(e101), "placed_views": []}
        for vid0 in e101.GetAllPlacedViews():
            v = tdoc.GetElement(vid0)
            sheet_info["placed_views"].append({"id": vid0.IntegerValue, "name": ename(v),
                                              "type": v.ViewType.ToString()})
            if v.ViewType == ViewType.FloorPlan and plan_view is None:
                plan_view = v
        for vid0 in e101.GetAllPlacedViews():
            v = tdoc.GetElement(vid0)
            if v.ViewType == ViewType.FloorPlan and "POWER" in (ename(v) or "").upper():
                plan_view = v; break
    if plan_view is None:
        for v in FilteredElementCollector(tdoc).OfClass(View).ToElements():
            if v.ViewType == ViewType.FloorPlan and (not v.IsTemplate) and "POWER" in (ename(v) or "").upper():
                plan_view = v; break
    data["sheet"] = sheet_info
    data["all_sheets"] = [{"num": s.SheetNumber, "name": ename(s),
                           "issue": s.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE).AsString() if s.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE) else None}
                          for s in sheets]
    if plan_view is None:
        raise Exception("No E101 power plan view found")
    data["plan_view"] = {"id": plan_view.Id.IntegerValue, "name": ename(plan_view),
                         "scale": plan_view.Scale}
    vid = plan_view.Id

    # ---- MEP spaces ----
    result["stage"] = "spaces"
    spaces = []
    for sp in FilteredElementCollector(tdoc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType():
        try:
            bb = sp.get_BoundingBox(None)
            spaces.append({"id": sp.Id.IntegerValue, "number": sp.Number, "name": ename(sp),
                           "loc": pt(sp.Location.Point) if sp.Location else None,
                           "bb": [pt(bb.Min), pt(bb.Max)] if bb else None})
        except:
            pass
    data["spaces"] = spaces

    # ---- electrical fixtures MODEL-WIDE (worksets may be closed; view filter would miss them) ----
    result["stage"] = "fixtures"
    fixtures = []
    for el in FilteredElementCollector(tdoc).OfCategory(BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType():
        try:
            d = {"id": el.Id.IntegerValue,
                 "family": el.Symbol.Family.Name, "type": ename(el.Symbol),
                 "loc": pt(el.Location.Point) if el.Location and hasattr(el.Location, "Point") else None,
                 "rot": round(el.Location.Rotation, 4) if el.Location and hasattr(el.Location, "Rotation") else None,
                 "facing": pt(el.FacingOrientation) if isinstance(el, FamilyInstance) else None,
                 "params": params_of(el)}
            try:
                syss = el.MEPModel.GetElectricalSystems()
                d["circuits"] = [s.Id.IntegerValue for s in syss] if syss else []
            except:
                d["circuits"] = []
            fixtures.append(d)
        except:
            pass
    data["fixtures"] = fixtures

    # ---- electrical equipment (panels) whole model ----
    result["stage"] = "equipment"
    equip = []
    for el in FilteredElementCollector(tdoc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType():
        try:
            equip.append({"id": el.Id.IntegerValue, "family": el.Symbol.Family.Name,
                          "type": ename(el.Symbol),
                          "loc": pt(el.Location.Point) if el.Location and hasattr(el.Location, "Point") else None,
                          "params": params_of(el)})
        except:
            pass
    data["equipment"] = equip

    # ---- electrical systems (circuits) ----
    result["stage"] = "systems"
    systems = []
    for es in FilteredElementCollector(tdoc).OfClass(ElectricalSystem).ToElements():
        try:
            systems.append({"id": es.Id.IntegerValue,
                            "circuit": es.CircuitNumber,
                            "panel": es.PanelName,
                            "load_name": es.LoadName,
                            "rating": round(es.Rating, 1),
                            "volts": round(es.Voltage, 1),
                            "poles": es.PolesNumber,
                            "app_load_va": round(es.ApparentLoad, 1),
                            "members": [m.Id.IntegerValue for m in es.Elements]})
        except:
            pass
    data["systems"] = systems

    # ---- wires MODEL-WIDE with owner view recorded ----
    result["stage"] = "wires"
    wires = []
    for w in FilteredElementCollector(tdoc).OfCategory(BuiltInCategory.OST_Wire).WhereElementIsNotElementType():
        try:
            verts = []
            try:
                for i in range(w.NumberOfVertices):
                    verts.append(pt(w.GetVertex(i)))
            except:
                lc = w.Location
                if lc and hasattr(lc, "Curve"):
                    verts = [pt(lc.Curve.GetEndPoint(0)), pt(lc.Curve.GetEndPoint(1))]
            tname = None
            try:
                tel = tdoc.GetElement(w.GetTypeId())
                tname = ename(tel)
            except:
                pass
            wires.append({"id": w.Id.IntegerValue, "type": tname,
                          "ownerview": w.OwnerViewId.IntegerValue,
                          "verts": verts, "params": params_of(w)})
        except:
            pass
    data["wires"] = wires

    # ---- tags in plan view ----
    result["stage"] = "tags"
    def collect_tags(bic):
        out = []
        for t in FilteredElementCollector(tdoc).OfCategory(bic).WhereElementIsNotElementType():
            try:
                d = {"id": t.Id.IntegerValue, "text": None, "head": None, "host": None,
                     "family": None, "type": None,
                     "ownerview": t.OwnerViewId.IntegerValue}
                try: d["text"] = t.TagText
                except: pass
                try: d["head"] = pt(t.TagHeadPosition)
                except: pass
                try: d["host"] = [i.IntegerValue for i in t.GetTaggedLocalElementIds()]
                except: pass
                try:
                    sym = tdoc.GetElement(t.GetTypeId())
                    d["type"] = ename(sym)
                    d["family"] = sym.FamilyName
                except: pass
                try:
                    d["leader"] = t.HasLeader
                except: pass
                out.append(d)
            except:
                pass
        return out
    data["wire_tags"] = collect_tags(BuiltInCategory.OST_WireTags)
    data["fixture_tags"] = collect_tags(BuiltInCategory.OST_ElectricalFixtureTags)
    data["equip_tags"] = collect_tags(BuiltInCategory.OST_ElectricalEquipmentTags)

    # ---- generic annotations (keynotes) in plan view ----
    result["stage"] = "keynotes"
    notes = []
    for el in FilteredElementCollector(tdoc, vid).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType():
        try:
            notes.append({"id": el.Id.IntegerValue,
                          "family": el.Symbol.Family.Name if hasattr(el, "Symbol") else None,
                          "type": ename(el.Symbol) if hasattr(el, "Symbol") else None,
                          "loc": pt(el.Location.Point) if el.Location and hasattr(el.Location, "Point") else None,
                          "rot": round(el.Location.Rotation, 4) if el.Location and hasattr(el.Location, "Rotation") else None,
                          "params": params_of(el)})
        except:
            pass
    data["keynotes"] = notes

    # ---- text notes in plan view + sheet ----
    result["stage"] = "text"
    def collect_text(view_id):
        out = []
        for t in FilteredElementCollector(tdoc, view_id).OfClass(TextNote).ToElements():
            try:
                out.append({"loc": pt(t.Coord), "text": t.Text.strip()})
            except:
                pass
        return out
    data["text_plan"] = collect_text(vid)
    sheet_text = []
    if e101 is not None:
        sheet_text.extend(collect_text(e101.Id))
        for vid0 in e101.GetAllPlacedViews():
            if vid0 != vid:
                sheet_text.extend(collect_text(vid0))
    data["text_sheet"] = sheet_text

    # ---- detail lines (keynote leaders) in plan view ----
    result["stage"] = "leaders"
    leaders = []
    for c in FilteredElementCollector(tdoc, vid).OfClass(CurveElement).ToElements():
        try:
            crv = c.GeometryCurve
            leaders.append({"id": c.Id.IntegerValue,
                            "p0": pt(crv.GetEndPoint(0)), "p1": pt(crv.GetEndPoint(1)),
                            "style": ename(c.LineStyle) if c.LineStyle else None})
        except:
            pass
    data["leaders"] = leaders

    # ---- CAD + Revit links ----
    result["stage"] = "links"
    links = []
    for li in FilteredElementCollector(tdoc).OfClass(ImportInstance).ToElements():
        try:
            tf = li.GetTotalTransform()
            links.append({"kind": "cad", "id": li.Id.IntegerValue,
                          "name": li.Category.Name if li.Category else None,
                          "origin": pt(tf.Origin), "basisX": pt(tf.BasisX),
                          "pinned": li.Pinned})
        except:
            pass
    for li in FilteredElementCollector(tdoc).OfClass(RevitLinkInstance).ToElements():
        try:
            tf = li.GetTotalTransform()
            links.append({"kind": "rvt", "id": li.Id.IntegerValue, "name": ename(li),
                          "origin": pt(tf.Origin), "basisX": pt(tf.BasisX)})
        except:
            pass
    data["links"] = links

    with open(os.path.join(OUT_DIR, "extract.json"), "w") as f:
        json.dump(data, f)
    result = {"ok": True, "doc": tdoc.Title,
              "counts": dict((k, len(v)) for k, v in data.items() if isinstance(v, list))}
except Exception as e:
    result = {"ok": False, "stage": result.get("stage"), "err": str(e),
              "tb": traceback.format_exc()[-800:]}

with open(os.path.join(OUT_DIR, "extract_status.json"), "w") as f:
    json.dump(result, f)
print(json.dumps(result))
