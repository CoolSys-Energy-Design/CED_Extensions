# place_fixture.py — Planet Fitness MCP skills library
# Look up and place electrical fixtures (wall-hosted or work-plane).
# IronPython 2.7. Caller manages Transaction.

import clr
from Autodesk.Revit import DB
from Autodesk.Revit.DB.Structure import StructuralType

_NAME_PROP = clr.GetClrType(DB.Element).GetProperty('Name')


def el_name(el):
    if el is None:
        return ''
    try:
        return _NAME_PROP.GetValue(el, None) or ''
    except:
        return ''


def _activate(doc, sym):
    if not sym.IsActive:
        sym.Activate()
        doc.Regenerate()
    return sym


def list_fixture_types(doc, category_bic=None, search=None):
    """Return list of (family_name, type_name, symbol_id_int) for the given category
    (default OST_ElectricalFixtures), optionally filtered by case-insensitive substring."""
    if category_bic is None:
        category_bic = DB.BuiltInCategory.OST_ElectricalFixtures
    out = []
    s = search.lower() if search else None
    for sym in DB.FilteredElementCollector(doc).OfCategory(category_bic).WhereElementIsElementType():
        try:
            tn = el_name(sym)
            fn = sym.Family.Name if hasattr(sym, 'Family') and sym.Family else ''
        except:
            tn = ''; fn = ''
        if s and (s not in tn.lower() and s not in fn.lower()):
            continue
        out.append((fn, tn, sym.Id.IntegerValue))
    return out


def find_family_symbol(doc, family_name, type_name, category_bic=None):
    """Lookup a FamilySymbol by family + type name. Activates if needed. Raises ValueError if missing."""
    if category_bic is None:
        category_bic = DB.BuiltInCategory.OST_ElectricalFixtures
    for sym in DB.FilteredElementCollector(doc).OfCategory(category_bic).WhereElementIsElementType():
        try:
            fn = sym.Family.Name if hasattr(sym, 'Family') and sym.Family else ''
            tn = el_name(sym)
        except:
            continue
        if fn == family_name and tn == type_name:
            return _activate(doc, sym)
    raise ValueError('FamilySymbol not found: %s :: %s in cat %s' % (family_name, type_name, category_bic))


def find_nearest_wall(doc, point_xyz, max_dist=None):
    """Return (wall, closest_point_on_wall, distance) for the nearest Wall to point_xyz, in XY plane.
    None if none found within max_dist."""
    best = None
    best_d = None
    p = DB.XYZ(point_xyz.X, point_xyz.Y, 0)
    for w in DB.FilteredElementCollector(doc).OfClass(DB.Wall).WhereElementIsNotElementType():
        loc = w.Location
        if not isinstance(loc, DB.LocationCurve):
            continue
        crv = loc.Curve
        try:
            proj = crv.Project(p)
            if proj is None:
                continue
            d = proj.Distance
        except:
            continue
        if best is None or d < best_d:
            best = (w, proj.XYZPoint, d)
            best_d = d
    if best is None:
        return None
    if max_dist is not None and best_d > max_dist:
        return None
    return best


def place_workplane_fixture(doc, symbol, level, point_xyz, rotation_rad=0.0):
    """Place a non-hosted fixture on a level/work-plane. Caller has open transaction.
    Returns the new FamilyInstance."""
    _activate(doc, symbol)
    inst = doc.Create.NewFamilyInstance(point_xyz, symbol, level, StructuralType.NonStructural)
    if rotation_rad and rotation_rad != 0.0:
        axis = DB.Line.CreateBound(point_xyz, DB.XYZ(point_xyz.X, point_xyz.Y, point_xyz.Z + 1))
        DB.ElementTransformUtils.RotateElement(doc, inst.Id, axis, rotation_rad)
    return inst


def place_face_hosted_fixture_on_wall(doc, symbol, point_xyz, wall=None, rotation_rad=0.0, max_search_dist=12.0):
    """Place a face-hosted fixture on a Wall. If wall is None, find nearest wall (max 12 ft default).
    Caller has open transaction. Returns the new FamilyInstance, or raises if no host found."""
    _activate(doc, symbol)
    if wall is None:
        hit = find_nearest_wall(doc, point_xyz, max_dist=max_search_dist)
        if hit is None:
            raise ValueError('No wall found within %.1f ft of (%.2f, %.2f)' % (max_search_dist, point_xyz.X, point_xyz.Y))
        wall, _proj, _d = hit
    # Use the wall's exterior face reference
    try:
        side = DB.ShellLayerType.Exterior
    except:
        side = None
    try:
        refs = DB.HostObjectUtils.GetSideFaces(wall, DB.ShellLayerType.Exterior)
    except:
        refs = None
    if not refs:
        refs = DB.HostObjectUtils.GetSideFaces(wall, DB.ShellLayerType.Interior)
    if not refs or refs.Count == 0:
        raise ValueError('Wall %d has no face references' % wall.Id.IntegerValue)
    face_ref = refs[0]
    # refDir for face placement (horizontal direction along the wall)
    ref_dir = DB.XYZ(1, 0, 0)
    inst = doc.Create.NewFamilyInstance(face_ref, point_xyz, ref_dir, symbol)
    if rotation_rad:
        axis = DB.Line.CreateBound(point_xyz, DB.XYZ(point_xyz.X, point_xyz.Y, point_xyz.Z + 1))
        DB.ElementTransformUtils.RotateElement(doc, inst.Id, axis, rotation_rad)
    return inst


def set_mounting_height_inches(elem, height_in):
    """Try the usual mounting-height parameter names. Returns True if any one set."""
    feet = float(height_in) / 12.0
    candidates = ['Elevation from Level', 'Mounting Height', 'Elevation', 'INSTANCE_ELEVATION_PARAM']
    for name in candidates:
        p = elem.LookupParameter(name)
        if p is not None and not p.IsReadOnly:
            try:
                if p.StorageType == DB.StorageType.Double:
                    p.Set(feet)
                    return True
                if p.StorageType == DB.StorageType.String:
                    p.Set('%g"' % height_in)
                    return True
            except:
                pass
    # Built-in param fallback
    try:
        bip = elem.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM)
        if bip is not None and not bip.IsReadOnly:
            bip.Set(feet)
            return True
    except:
        pass
    return False
