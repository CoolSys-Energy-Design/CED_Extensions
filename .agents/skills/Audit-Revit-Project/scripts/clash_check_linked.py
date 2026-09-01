# -*- coding: utf-8 -*-
# clash_check_linked.py — host-model vs LINKED-model interference check.
# (coding line above is REQUIRED: IronPython 2 refuses non-ASCII source bytes
# without it, and this file's comments contain em-dashes.)
# Runs inside Revit via execute_revit_code (IronPython 2.7: print statement,
# no f-strings, `except Exception, ex:`).
#
# READ-ONLY: collectors + geometry extraction only. No transactions.
#
# Why this exists: ElementIntersectsElementFilter only works between elements
# of the SAME document, so the in-document clash check (see
# routines\heb-qaqc\checks\clash_check.py) is blind to links. This script
# closes that gap:
#   for each loaded RevitLinkInstance:
#     for each (host category, link category) pair:
#       for each link element: transform its bounding box into host
#       coordinates (all 8 corners through GetTotalTransform) ->
#       BoundingBoxIntersectsFilter prefilter on the host collector ->
#       if any candidates: extract the link element's solids, transform them
#       with SolidUtils.CreateTransformed, and run
#       ElementIntersectsSolidFilter against the host candidates.
# Survivors are true solid intersections in host coordinates.
#
# Not covered (v1): link-vs-link clashes; clearance tolerances (an actual
# intersection is required); insulation.
#
# Before sending via execute_revit_code, the session must replace __OUT_PATH__
# with the absolute path of the output JSON (scratchpad). The MCP call will
# time out at 60 s on a big model — poll for the output file instead; any
# crash is also written there (see bottom).

import json, os

OUT_PATH = r"__OUT_PATH__"

# (label, host BuiltInCategory, link BuiltInCategory)
# Structural first — pipes/ducts through beams and columns is the headline.
# Walls/floors are OFF by default: intentional penetrations make them noisy;
# add them per-run only if the user asks.
PAIRS = [
    ("Pipes vs Linked Structural Framing",       "OST_PipeCurves", "OST_StructuralFraming"),
    ("Ducts vs Linked Structural Framing",       "OST_DuctCurves", "OST_StructuralFraming"),
    ("Cable Trays vs Linked Structural Framing", "OST_CableTray",  "OST_StructuralFraming"),
    ("Pipes vs Linked Structural Columns",       "OST_PipeCurves", "OST_StructuralColumns"),
    ("Ducts vs Linked Structural Columns",       "OST_DuctCurves", "OST_StructuralColumns"),
]

# Substring to limit which links are checked (e.g. "STRUCT"), or None for all.
LINK_NAME_FILTER = None

MIN_SOLID_VOLUME = 1e-6  # cubic feet; skips degenerate solids


def clean(s):
    # ASCII-sanitize model text — degree signs etc. crash IronPython's json.
    try:
        out = []
        for ch in s:
            o = ord(ch)
            out.append(ch if 32 <= o <= 126 else "?")
        return "".join(out)
    except Exception, ex:
        return "N/A"


def eid(element):
    try:
        return int(element.Id.Value)         # Revit 2024+
    except Exception, ex:
        return int(element.Id.IntegerValue)  # deprecated fallback


def safe_name(element):
    return clean(getattr(element, "Name", None) or "N/A")


def level_name(element, host_doc):
    try:
        lvl = host_doc.GetElement(element.LevelId)
        return clean(getattr(lvl, "Name", "N/A")) if lvl else "N/A"
    except Exception, ex:
        return "N/A"


def bbox_center(element):
    try:
        bb = element.get_BoundingBox(None)
        if not bb:
            return None
        c = (bb.Min + bb.Max) * 0.5
        return [round(c.X, 1), round(c.Y, 1), round(c.Z, 1)]
    except Exception, ex:
        return None


def solids_of(elem, opts):
    out = []
    try:
        geo = elem.get_Geometry(opts)
        if not geo:
            return out
        for g in geo:
            if isinstance(g, DB.Solid):
                if g.Volume > MIN_SOLID_VOLUME:
                    out.append(g)
            elif isinstance(g, DB.GeometryInstance):
                for g2 in g.GetInstanceGeometry():
                    if isinstance(g2, DB.Solid) and g2.Volume > MIN_SOLID_VOLUME:
                        out.append(g2)
    except Exception, ex:
        pass
    return out


def transformed_outline(bb, xform):
    # Transform all 8 bbox corners into host coordinates and take min/max —
    # transforming just Min/Max is wrong under rotation.
    pts = []
    for x in (bb.Min.X, bb.Max.X):
        for y in (bb.Min.Y, bb.Max.Y):
            for z in (bb.Min.Z, bb.Max.Z):
                pts.append(xform.OfPoint(DB.XYZ(x, y, z)))
    xs = [p.X for p in pts]
    ys = [p.Y for p in pts]
    zs = [p.Z for p in pts]
    return DB.Outline(DB.XYZ(min(xs), min(ys), min(zs)),
                      DB.XYZ(max(xs), max(ys), max(zs)))


clashes = []
errors = []
links = []
seen = {}

try:
    opts = DB.Options()
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False

    for li in (DB.FilteredElementCollector(doc)
               .OfClass(DB.RevitLinkInstance).ToElements()):
        try:
            ldoc = li.GetLinkDocument()
            iname = safe_name(li)
            if ldoc is None:
                links.append({"link": iname, "status": "unloaded - NOT checked"})
                continue
            ltitle = clean(ldoc.Title)
            if LINK_NAME_FILTER and LINK_NAME_FILTER.lower() not in ltitle.lower():
                links.append({"link": ltitle, "status": "skipped by filter"})
                continue
            links.append({"link": ltitle, "status": "checked"})
            xform = li.GetTotalTransform()

            for label, host_cat_name, link_cat_name in PAIRS:
                try:
                    host_cat = getattr(DB.BuiltInCategory, host_cat_name)
                    link_cat = getattr(DB.BuiltInCategory, link_cat_name)
                    link_elems = (DB.FilteredElementCollector(ldoc)
                                  .OfCategory(link_cat)
                                  .WhereElementIsNotElementType().ToElements())
                    for le in link_elems:
                        try:
                            bb = le.get_BoundingBox(None)
                            if not bb:
                                continue
                            outline = transformed_outline(bb, xform)
                            bb_filter = DB.BoundingBoxIntersectsFilter(outline)
                            # Keep ToElementIds() as the .NET ICollection —
                            # wrapping it in a Python list() breaks the
                            # FilteredElementCollector(doc, ids) overload
                            # ("expected ElementId, got list").
                            candidates = (DB.FilteredElementCollector(doc)
                                          .OfCategory(host_cat)
                                          .WhereElementIsNotElementType()
                                          .WherePasses(bb_filter)
                                          .ToElementIds())
                            if candidates.Count == 0:
                                continue  # cheap out before any geometry work
                            for s in solids_of(le, opts):
                                try:
                                    ts = s if xform.IsIdentity else \
                                        DB.SolidUtils.CreateTransformed(s, xform)
                                    hits = (DB.FilteredElementCollector(doc, candidates)
                                            .WherePasses(DB.ElementIntersectsSolidFilter(ts)))
                                    for h in hits:
                                        key = (eid(h), eid(li), eid(le))
                                        if key in seen:
                                            continue
                                        seen[key] = True
                                        clashes.append({
                                            "pair": label,
                                            "link": ltitle,
                                            "host_id": eid(h),
                                            "host_name": safe_name(h),
                                            "link_elem_id": eid(le),
                                            "link_elem_name": safe_name(le),
                                            "level": level_name(h, doc),
                                            "location": bbox_center(h),
                                        })
                                except Exception, ex3:
                                    errors.append("{0} solid of link elem {1}: {2}".format(
                                        label, eid(le), ex3))
                        except Exception, ex2:
                            errors.append("{0} link elem {1}: {2}".format(
                                label, eid(le), ex2))
                except Exception, ex1:
                    errors.append("{0} in link {1}: {2}".format(label, ltitle, ex1))
        except Exception, exl:
            errors.append("link instance {0}: {1}".format(eid(li), exl))
except Exception, ex_top:
    errors.append("fatal: {0}".format(ex_top))

result = {
    "check": "clash_linked",
    "model": clean(doc.Title),
    "revit_version": getattr(doc.Application, "VersionNumber", "unknown"),
    "links": links,
    "clash_count": len(clashes),
    "clashes": clashes,
    "errors": [clean(e) for e in errors],
}

f = open(OUT_PATH, "w")
f.write(json.dumps(result))
f.close()
print("linked clash check done: {0} clashes across {1} links, {2} errors -> {3}".format(
    len(clashes), len(links), len(errors), OUT_PATH))
