# SKILL: collect_linked_elements
# Collects linked (Arch/Equip) model elements whose location falls inside the
# BAKERY view crop box. These are the ANCHORS for equipment-relative placement.
# Run via mcp__revit__execute_revit_code with SOURCE doc active.
# Writes: c:\CED_Extensions\HEB_MCP_AUTOMATION\bakery_auto\data\linked_elements.json
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())  # DATA, VIEW_NAME
import json, io
OUT = DATA + r"\linked_elements.json"

def nm(e):
    try:
        v = DB.Element.Name.GetValue(e)
        if v: return v
    except: pass
    try: return getattr(e, 'Name', '?')
    except: return '?'

bv = next(v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
          if getattr(v, 'Name', '') == VIEW_NAME and not v.IsTemplate)
cb = bv.CropBox
inv = cb.Transform.Inverse
def in_crop(p):
    q = inv.OfPoint(p)
    return (cb.Min.X - 0.5 <= q.X <= cb.Max.X + 0.5 and
            cb.Min.Y - 0.5 <= q.Y <= cb.Max.Y + 0.5)

def loc_pt(el):
    try:
        L = el.Location
        if isinstance(L, DB.LocationPoint): return L.Point
        if isinstance(L, DB.LocationCurve):
            c = L.Curve; return c.Evaluate(0.5, True)
    except: pass
    try:
        bb = el.get_BoundingBox(None)
        if bb: return (bb.Min + bb.Max) * 0.5
    except: pass
    return None

links = [el for el in DB.FilteredElementCollector(doc, bv.Id)
         if isinstance(el, DB.RevitLinkInstance)]
records = []
for li in links:
    ld = li.GetLinkDocument()
    if not ld: continue
    tf = li.GetTotalTransform()
    lname = nm(li)
    for el in DB.FilteredElementCollector(ld).WhereElementIsNotElementType():
        try:
            cat = el.Category.Name if el.Category else None
        except: cat = None
        if not cat: continue
        lp = loc_pt(el)
        if lp is None: continue
        wp = tf.OfPoint(lp)
        if not in_crop(wp): continue
        try: fs = ld.GetElement(el.GetTypeId())
        except: fs = None
        mk = None
        try:
            mp = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
            mk = mp.AsString() if mp else None
        except: pass
        bbw = None
        try:
            bb = el.get_BoundingBox(None)
            if bb:
                mn = tf.OfPoint(bb.Min); mx = tf.OfPoint(bb.Max)
                bbw = [round(mn.X,3),round(mn.Y,3),round(mn.Z,3),
                       round(mx.X,3),round(mx.Y,3),round(mx.Z,3)]
        except: pass
        fo = getattr(el, 'FacingOrientation', None)
        records.append({
            "link": lname,
            "link_doc": ld.Title,
            "id": el.Id.IntegerValue,
            "category": cat,
            "family": nm(fs.Family) if fs and hasattr(fs, 'Family') else None,
            "type": nm(fs) if fs else None,
            "mark": mk,
            "world_xyz": [round(wp.X,3), round(wp.Y,3), round(wp.Z,3)],
            "bbox_world": bbw,
            "facing": [round(fo.X,3), round(fo.Y,3)] if fo else None,
        })

with io.open(OUT, 'w', encoding='utf-8') as f:
    f.write(json.dumps({"view": VIEW_NAME, "source_doc": doc.Title,
                        "count": len(records), "elements": records}, indent=1))

bycat = {}
for r in records: bycat[r["category"]] = bycat.get(r["category"], 0) + 1
print("linked elements in crop:", len(records))
for k in sorted(bycat, key=lambda z: -bycat[z]): print("  %-30s %d" % (k, bycat[k]))
print("written:", OUT)
