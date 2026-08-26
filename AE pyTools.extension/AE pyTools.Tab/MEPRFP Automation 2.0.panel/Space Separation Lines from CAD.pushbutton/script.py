# -*- coding: utf-8 -*-

"""Create Revit space separation lines from selected layers in a CAD link/import."""

from pyrevit import DB, UI, forms, revit, script


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def _layer_name(geometry_object):
    """Return the CAD layer name represented by a geometry object's graphics style."""
    try:
        style_id = geometry_object.GraphicsStyleId
        if style_id == DB.ElementId.InvalidElementId:
            return None

        style = doc.GetElement(style_id)
        if style and style.GraphicsStyleCategory:
            return style.GraphicsStyleCategory.Name
    except Exception:
        pass

    return None


def _collect_geometry(geometry, curves, layer_names=None):
    """Collect all layer names and only line/polyline segments for placement."""
    if not geometry:
        return

    for geometry_object in geometry:
        layer = _layer_name(geometry_object)
        if layer and layer_names is not None:
            # Discover the layer regardless of whether its geometry is placeable.
            # Arcs, text, points, blocks, etc. should not hide their CAD layer.
            layer_names.add(layer)

        if isinstance(geometry_object, DB.Line):
            if layer:
                curves.append((layer, geometry_object))
            continue

        if isinstance(geometry_object, DB.PolyLine):
            if layer:
                coordinates = geometry_object.GetCoordinates()
                for index in range(len(coordinates) - 1):
                    start = coordinates[index]
                    end = coordinates[index + 1]
                    if start.DistanceTo(end) > 1e-9:
                        curves.append((layer, DB.Line.CreateBound(start, end)))
            continue

        if isinstance(geometry_object, DB.GeometryInstance):
            # GetInstanceGeometry returns the nested geometry in model coordinates,
            # which also handles CAD blocks without applying the transform twice.
            try:
                _collect_geometry(geometry_object.GetInstanceGeometry(), curves, layer_names)
            except Exception:
                try:
                    _collect_geometry(geometry_object.GetSymbolGeometry(), curves, layer_names)
                except Exception:
                    pass


def _pick_cad_instance():
    reference = uidoc.Selection.PickObject(
        UI.Selection.ObjectType.Element,
        "Select the linked or imported CAD file"
    )
    element = doc.GetElement(reference.ElementId)
    if not isinstance(element, DB.ImportInstance):
        forms.alert("The selected element is not a linked or imported CAD file.", exitscript=True)
    return element


def _choose_layers(layer_names):
    selected = forms.SelectFromList.show(
        sorted(layer_names),
        title="Select CAD layers for space separation lines",
        multiselect=True,
        button_name="Create Space Separation Lines"
    )
    if not selected:
        script.exit()
    return set(selected)


def _transform_curve(curve, transform):
    """Apply the CAD instance transform to a line's endpoints."""
    start = transform.OfPoint(curve.GetEndPoint(0))
    end = transform.OfPoint(curve.GetEndPoint(1))
    return DB.Line.CreateBound(start, end)


def _collect_cad_layer_names(cad_instance, geometry, layer_names):
    """Read the CAD layer table first, then supplement it from geometry styles."""
    try:
        category = cad_instance.Category
        if category and category.SubCategories:
            for subcategory in category.SubCategories:
                if subcategory and subcategory.Name:
                    layer_names.add(subcategory.Name)
    except Exception:
        pass

    # Some linked CAD files expose additional styles only through geometry.
    _collect_geometry(geometry, [], layer_names)


def _same_line(point, line_start, line_end, tolerance):
    direction = line_end - line_start
    if direction.GetLength() <= tolerance:
        return False
    return direction.CrossProduct(point - line_start).GetLength() <= tolerance


def _merge_collinear_curves(layered_curves, tolerance=1e-5):
    """Merge touching collinear CAD segments, keeping layers separate."""
    merged = []
    by_layer = {}
    for layer, curve in layered_curves:
        by_layer.setdefault(layer, []).append(curve)

    for layer, curves in by_layer.items():
        pending = list(curves)
        changed = True
        while changed:
            changed = False
            result = []
            while pending:
                current = pending.pop(0)
                merged_current = False
                index = 0
                while index < len(pending):
                    other = pending[index]
                    a_start = current.GetEndPoint(0)
                    a_end = current.GetEndPoint(1)
                    b_start = other.GetEndPoint(0)
                    b_end = other.GetEndPoint(1)

                    touching = (
                        a_end.DistanceTo(b_start) <= tolerance or
                        a_end.DistanceTo(b_end) <= tolerance or
                        a_start.DistanceTo(b_start) <= tolerance or
                        a_start.DistanceTo(b_end) <= tolerance
                    )
                    collinear = (
                        _same_line(b_start, a_start, a_end, tolerance) and
                        _same_line(b_end, a_start, a_end, tolerance)
                    )

                    if touching and collinear:
                        points = [a_start, a_end, b_start, b_end]
                        farthest_start = points[0]
                        farthest_end = points[1]
                        farthest_distance = farthest_start.DistanceTo(farthest_end)
                        for first in points:
                            for second in points:
                                distance = first.DistanceTo(second)
                                if distance > farthest_distance:
                                    farthest_start = first
                                    farthest_end = second
                                    farthest_distance = distance
                        current = DB.Line.CreateBound(farthest_start, farthest_end)
                        pending.pop(index)
                        changed = True
                        merged_current = True
                        index = 0
                        continue
                    index += 1

                result.append(current)
            pending = result

        for curve in pending:
            merged.append((layer, curve))

    return merged


def _create_lines(cad_instance, selected_layers, view):
    options = DB.Options()
    options.ComputeReferences = False
    options.IncludeNonVisibleObjects = True

    curves = []
    _collect_geometry(cad_instance.get_Geometry(options), curves)
    curves = [(layer, curve) for layer, curve in curves if layer in selected_layers]

    # CAD geometry can be returned in the link/import's source coordinates.
    # Apply the instance transform so created lines match the CAD position in the host model.
    cad_transform = cad_instance.GetTotalTransform()
    curves = [(layer, _transform_curve(curve, cad_transform)) for layer, curve in curves]
    curves = _merge_collinear_curves(curves)

    if not curves:
        forms.alert("No line or polyline geometry was found on the selected CAD layers.", exitscript=True)

    level = view.GenLevel
    if not level:
        forms.alert("The active plan view does not have an associated level.", exitscript=True)

    plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, level.Elevation))
    transaction = DB.Transaction(doc, "Create Space Separation Lines from CAD")
    transaction.Start()

    created = 0
    skipped = 0
    try:
        sketch_plane = DB.SketchPlane.Create(doc, plane)
        for layer, curve in curves:
            try:
                # CAD linework must lie on the active plan's level plane.
                start = curve.GetEndPoint(0)
                end = curve.GetEndPoint(1)
                if abs(start.Z - level.Elevation) > 1e-6 or abs(end.Z - level.Elevation) > 1e-6:
                    skipped += 1
                    continue

                # NewSpaceBoundaryLines requires a CurveArray, even for one curve.
                curve_array = DB.CurveArray()
                curve_array.Append(curve)
                doc.Create.NewSpaceBoundaryLines(sketch_plane, curve_array, view)
                created += 1
            except Exception as error:
                skipped += 1
                output.print_md("Skipped CAD curve: {}".format(error))
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    return created, skipped


def main():
    active_view = doc.ActiveView
    if not isinstance(active_view, DB.ViewPlan):
        forms.alert("Open a plan view before creating space separation lines.", exitscript=True)

    cad_instance = _pick_cad_instance()
    options = DB.Options()
    options.ComputeReferences = False
    options.IncludeNonVisibleObjects = True

    geometry = []
    layer_names = set()
    _collect_cad_layer_names(cad_instance, cad_instance.get_Geometry(options), layer_names)
    if not layer_names:
        forms.alert("No CAD layers were found in the selected file.", exitscript=True)

    selected_layers = _choose_layers(layer_names)
    created, skipped = _create_lines(cad_instance, selected_layers, active_view)
    forms.alert(
        "Created {} space separation line(s).{}".format(
            created,
            " Skipped {} curve(s) that were not on the view level or could not be created.".format(skipped)
            if skipped else ""
        )
    )


if __name__ == "__main__":
    main()
