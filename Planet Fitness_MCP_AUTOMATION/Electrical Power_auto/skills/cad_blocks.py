# cad_blocks.py — Planet Fitness MCP skills library
# Reads block-insert positions out of a view's linked CAD with proper world-transform composition.
# IronPython 2.7, exec()'d inside Revit MCP execute_revit_code.

import math

def _layer_of(g, doc):
    try:
        gid = g.GraphicsStyleId
        if gid is None or gid.IntegerValue <= 0:
            return None
        gs = doc.GetElement(gid)
        if gs is None:
            return None
        cat = gs.GraphicsStyleCategory
        return cat.Name if cat else None
    except:
        return None


def read_cad_blocks(doc, view, wanted_layers=None, building_bounds=None, max_depth=3):
    """
    Walk every linked CAD (ImportInstance) visible in `view`, recursing GeometryInstance
    children with proper world-transform accumulation, and return one record per block insert.

    Args:
        doc, view: Revit document and view.
        wanted_layers: optional iterable of layer names. If given, only block inserts whose
            layer matches are returned.
        building_bounds: optional (min_x, min_y, max_x, max_y) tuple. If given, only blocks
            with world_xy inside the rectangle are returned.
        max_depth: max recursion depth into nested blocks. Default 3.

    Returns: list of dicts with keys:
        - layer (str)
        - world_xy (tuple of float)
        - rotation_deg (float)
        - world_origin (tuple x, y, z)
        - world_basis_x (tuple dx, dy)
    """
    from Autodesk.Revit import DB
    if wanted_layers is not None:
        wanted_layers = set(wanted_layers)

    results = []

    def walk(g, world_xform, depth):
        if not isinstance(g, DB.GeometryInstance):
            return
        own_local = g.Transform
        world = own_local if world_xform is None else world_xform.Multiply(own_local)
        own_layer = _layer_of(g, doc)
        child_layers = {}
        try:
            child_geo = g.GetSymbolGeometry()
            for child in child_geo:
                cl = _layer_of(child, doc)
                if cl:
                    child_layers[cl] = child_layers.get(cl, 0) + 1
                if isinstance(child, DB.GeometryInstance) and depth < max_depth:
                    walk(child, world, depth + 1)
        except:
            pass
        chosen = own_layer
        if not chosen and child_layers:
            chosen = max(child_layers.iteritems(), key=lambda kv: kv[1])[0]
        if chosen is None:
            return
        if wanted_layers is not None and chosen not in wanted_layers:
            return
        wo = world.Origin
        bx = world.BasisX
        x, y = wo.X, wo.Y
        if building_bounds is not None:
            mnx, mny, mxx, mxy = building_bounds
            if not (mnx <= x <= mxx and mny <= y <= mxy):
                return
        results.append({
            'layer': chosen,
            'world_xy': (x, y),
            'rotation_deg': math.degrees(math.atan2(bx.Y, bx.X)),
            'world_origin': (wo.X, wo.Y, wo.Z),
            'world_basis_x': (bx.X, bx.Y),
        })

    opts = DB.Options()
    opts.View = view
    for inst in DB.FilteredElementCollector(doc, view.Id).OfClass(DB.ImportInstance):
        geom = inst.get_Geometry(opts)
        if geom is None:
            continue
        for g in geom:
            walk(g, None, 0)

    return results


def cluster_blocks_by_y(blocks, bin_size=2.0, min_count=8):
    """Group block records into Y-bands. Returns list of (y_center, [blocks]) sorted by Y desc."""
    bins = {}
    for b in blocks:
        yb = int(round(b['world_xy'][1] / bin_size)) * bin_size
        bins.setdefault(yb, []).append(b)
    rows = [(y, lst) for y, lst in bins.items() if len(lst) >= min_count]
    rows.sort(key=lambda kv: -kv[0])
    return rows


if False:
    # Smoke-test pattern (do NOT execute — just documentation):
    view = doc.ActiveView
    blocks = read_cad_blocks(doc, view,
                             wanted_layers={'A-N-GYM EQUIPMENT'},
                             building_bounds=(40, -310, 280, -185))
    print('Got', len(blocks), 'blocks')
