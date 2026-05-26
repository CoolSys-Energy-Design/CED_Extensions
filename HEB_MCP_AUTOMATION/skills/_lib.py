# SKILL LIB: shared helpers exec'd by the bakery_auto collectors/replicator.
# Usage inside an MCP execute_revit_code script:
#   exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import json, io, math

# Area-configurable: a runner sets BA_DATA / BA_VIEW before exec'ing skills.
# Defaults = BAKERY so existing bakery usage is unchanged.
try: DATA = BA_DATA
except NameError: DATA = r"c:\CED_Extensions\HEB_MCP_AUTOMATION\bakery_auto\data"
try: VIEW_NAME = BA_VIEW
except NameError: VIEW_NAME = "Power Callout - BAKERY - L1"

def nm(e):
    """Robust element/type name (plain .Name raises on some API types)."""
    if e is None: return None
    try:
        v = DB.Element.Name.GetValue(e)
        if v: return v
    except: pass
    for bip in (DB.BuiltInParameter.SYMBOL_NAME_PARAM,
                DB.BuiltInParameter.ALL_MODEL_TYPE_NAME,
                DB.BuiltInParameter.ELEM_TYPE_PARAM):
        try:
            p = e.get_Parameter(bip)
            if p:
                s = p.AsString() or p.AsValueString()
                if s: return s
        except: pass
    try: return getattr(e, 'Name', '?')
    except: return '?'

def type_name(inst):
    """Type/symbol name for a placed instance."""
    try:
        p = inst.get_Parameter(DB.BuiltInParameter.ELEM_TYPE_PARAM)
        if p:
            s = p.AsValueString()
            if s: return s
    except: pass
    return nm(inst.Document.GetElement(inst.GetTypeId()))

def get_view(d, name=VIEW_NAME):
    return next(v for v in DB.FilteredElementCollector(d).OfClass(DB.View)
                if getattr(v, 'Name', '') == name and not v.IsTemplate)

def crop_test(view, pad=0.5):
    cb = view.CropBox; inv = cb.Transform.Inverse
    def _in(p):
        q = inv.OfPoint(p)
        return (cb.Min.X-pad <= q.X <= cb.Max.X+pad and
                cb.Min.Y-pad <= q.Y <= cb.Max.Y+pad)
    return _in

def loc_pt(el):
    try:
        L = el.Location
        if isinstance(L, DB.LocationPoint): return L.Point
        if isinstance(L, DB.LocationCurve): return L.Curve.Evaluate(0.5, True)
    except: pass
    try:
        bb = el.get_BoundingBox(None)
        if bb: return (bb.Min + bb.Max) * 0.5
    except: pass
    return None

def params_of(el, only=None):
    d = {}
    for p in el.Parameters:
        try:
            n = p.Definition.Name
            if only and n not in only: continue
            st = p.StorageType
            if st == DB.StorageType.String: d[n] = p.AsString()
            elif st == DB.StorageType.Double: d[n] = round(p.AsDouble(), 5)
            elif st == DB.StorageType.Integer: d[n] = p.AsInteger()
            elif st == DB.StorageType.ElementId: d[n] = p.AsElementId().IntegerValue
        except: pass
    return d

def write_json(fname, payload):
    path = DATA + "\\" + fname
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(json.dumps(payload, indent=1))
    return path
