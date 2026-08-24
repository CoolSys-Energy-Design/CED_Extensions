# PF2 CAD-underlay extraction (IronPython, read-only, background doc by TITLE_KEY).
# Walks ImportInstance geometry; writes cad_geom.json with compressed shapes:
#   Line  -> {"lay": layer, "l": [x0,y0,x1,y1]}
#   PolyLine/other -> {"lay": layer, "bb": [minx,miny,maxx,maxy], "c": [cx,cy], "n": npts}
import json, os, traceback
from Autodesk.Revit.DB import (FilteredElementCollector, ImportInstance, Options,
                               GeometryInstance, Line, PolyLine, Arc, Element)

OUT_DIR = r"__OUT_DIR__"
TITLE_KEY = "__TITLE_KEY__"
result = {"ok": False}
try:
    app = uidoc.Application.Application
    tdoc = None
    for d2 in app.Documents:
        if (not d2.IsLinked) and TITLE_KEY.lower() in d2.Title.lower():
            tdoc = d2
    if tdoc is None:
        raise Exception("no doc " + TITLE_KEY)

    def gsname(gobj):
        try:
            gs = tdoc.GetElement(gobj.GraphicsStyleId)
            if gs is not None:
                return gs.GraphicsStyleCategory.Name
        except:
            pass
        return None

    shapes = []
    opts = Options()
    opts.IncludeNonVisibleObjects = False
    KEEP = ("WALL", "DOOR", "GLAZ", "LEASE", "DEMISING", "DEMO", "RACEWAY",
            "EQUIP", "FURN", "TELEVISION", "ELECTRICAL", "PLUMB", "OVHD",
            "FLOR", "GYM", "SPA", "MILLWORK", "CASEWORK")
    def keep_layer(lay):
        if not lay:
            return False
        L = lay.upper()
        for k in KEEP:
            if k in L:
                return True
        return False

    def emit(crv, lay):
        if isinstance(crv, Line):
            p0 = crv.GetEndPoint(0); p1 = crv.GetEndPoint(1)
            shapes.append({"lay": lay, "l": [round(p0.X,2), round(p0.Y,2), round(p1.X,2), round(p1.Y,2)]})
        elif isinstance(crv, PolyLine):
            pts = crv.GetCoordinates()
            xs = [p.X for p in pts]; ys = [p.Y for p in pts]
            shapes.append({"lay": lay,
                           "bb": [round(min(xs),2), round(min(ys),2), round(max(xs),2), round(max(ys),2)],
                           "c": [round(sum(xs)/len(xs),2), round(sum(ys)/len(ys),2)],
                           "n": len(pts)})
        else:
            try:
                p0 = crv.GetEndPoint(0); p1 = crv.GetEndPoint(1)
                shapes.append({"lay": lay, "a": [round(p0.X,2), round(p0.Y,2), round(p1.X,2), round(p1.Y,2)]})
            except:
                pass

    def walk(geo, depth, xf):
        for g in geo:
            if isinstance(g, GeometryInstance):
                if depth < 4:
                    walk(g.GetInstanceGeometry(), depth + 1, xf)
            else:
                lay = gsname(g)
                if not keep_layer(lay):
                    continue
                if isinstance(g, (Line, PolyLine)) or hasattr(g, "GetEndPoint"):
                    emit(g, lay)

    n_imports = 0
    for li in FilteredElementCollector(tdoc).OfClass(ImportInstance).ToElements():
        try:
            geo = li.get_Geometry(opts)
            if geo is None:
                continue
            n_imports += 1
            walk(geo, 0, None)
        except:
            pass

    with open(os.path.join(OUT_DIR, "cad_geom.json"), "w") as f:
        json.dump({"imports_walked": n_imports, "shapes": shapes}, f)
    result = {"ok": True, "imports": n_imports, "shapes": len(shapes)}
except Exception as e:
    result = {"ok": False, "err": str(e), "tb": traceback.format_exc()[-600:]}
with open(os.path.join(OUT_DIR, "cad_status.json"), "w") as f:
    json.dump(result, f)
print(json.dumps(result))
