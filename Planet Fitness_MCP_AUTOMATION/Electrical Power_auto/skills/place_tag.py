"""
place_tag.py - reusable tag-placement utilities for the Planet Fitness MCP
automation pipeline.

IronPython 2.7 / Revit API.

Public API:
    find_tag_type(doc, family_name, type_name, category_bic) -> FamilySymbol
    list_tag_types(doc, category_bic) -> [(family_name, type_name, symbol_id), ...]
    place_fixture_tag(doc, view, fixture_id, tag_type_symbol, head_xy,
                      leader=False, tag_mode=None) -> IndependentTag

Assumes the caller is responsible for opening / committing any transactions.
"""

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB  # noqa: E402

# ---------------------------------------------------------------------------
# IronPython gotcha: DB.Element.Name on a FamilySymbol is an ambiguous /
# hidden property in IronPython and `Element.Name.GetValue(sym)` will throw
# AttributeError. Pull the PropertyInfo via reflection and invoke it directly.
# ---------------------------------------------------------------------------
_NAME_PROP = clr.GetClrType(DB.Element).GetProperty('Name')


def el_name(el):
    """Reflection-safe Element.Name accessor for IronPython."""
    if el is None:
        return None
    return _NAME_PROP.GetValue(el, None)


def _family_name(symbol):
    """Return the owning Family's name for a FamilySymbol."""
    try:
        fam = symbol.Family
    except AttributeError:
        return None
    return el_name(fam) if fam is not None else None


def _collect_tag_symbols(doc, category_bic):
    """FilteredElementCollector over FamilySymbol of the given tag category."""
    return (DB.FilteredElementCollector(doc)
              .OfClass(DB.FamilySymbol)
              .OfCategory(category_bic)
              .ToElements())


def list_tag_types(doc, category_bic):
    """List all tag types in the given category.

    Returns a list of (family_name, type_name, symbol_id_int) tuples,
    sorted by family then type.
    """
    out = []
    for sym in _collect_tag_symbols(doc, category_bic):
        fn = _family_name(sym) or ''
        tn = el_name(sym) or ''
        out.append((fn, tn, sym.Id.IntegerValue))
    out.sort(key=lambda t: (t[0].lower(), t[1].lower()))
    return out


def find_tag_type(doc, family_name, type_name, category_bic):
    """Look up a tag FamilySymbol by family + type name in the given tag
    category (e.g. OST_ElectricalFixtureTags, OST_WireTags, OST_KeynoteTags).

    Raises ValueError if no match. Activates the symbol if it is not already
    active. Caller must hold a transaction if activation will occur.
    """
    target_fam = (family_name or '').strip().lower()
    target_typ = (type_name or '').strip().lower()
    for sym in _collect_tag_symbols(doc, category_bic):
        fn = (_family_name(sym) or '').strip().lower()
        tn = (el_name(sym) or '').strip().lower()
        if fn == target_fam and tn == target_typ:
            if not sym.IsActive:
                sym.Activate()
                doc.Regenerate()
            return sym
    raise ValueError(
        "Tag type not found: family=%r type=%r category=%r"
        % (family_name, type_name, category_bic))


def place_fixture_tag(doc, view, fixture_id, tag_type_symbol, head_xy,
                      leader=False, tag_mode=None):
    """Create an IndependentTag on a given fixture.

    Parameters
    ----------
    doc              : DB.Document
    view             : DB.View              -- target view (must be a plan/etc.)
    fixture_id       : DB.ElementId or int  -- host fixture id
    tag_type_symbol  : DB.FamilySymbol      -- desired tag type
    head_xy          : (x, y) tuple of world coords for tag head
    leader           : bool                 -- attach leader line
    tag_mode         : DB.TagMode           -- defaults to TM_ADDBY_CATEGORY

    Returns the new DB.IndependentTag.  Caller owns the transaction.
    """
    # Normalize fixture id
    if isinstance(fixture_id, DB.ElementId):
        eid = fixture_id
    else:
        eid = DB.ElementId(int(fixture_id))

    fixture = doc.GetElement(eid)
    if fixture is None:
        raise ValueError("Fixture not found for id %s" % (fixture_id,))

    # Z coordinate -- prefer the view's level elevation if it has one.
    z = 0.0
    try:
        lvl = view.GenLevel
        if lvl is not None:
            z = float(lvl.Elevation)
    except AttributeError:
        z = 0.0

    head_x = float(head_xy[0])
    head_y = float(head_xy[1])
    head_pt = DB.XYZ(head_x, head_y, z)

    if tag_mode is None:
        tag_mode = DB.TagMode.TM_ADDBY_CATEGORY
    orientation = DB.TagOrientation.Horizontal

    ref = DB.Reference(fixture)

    tag = DB.IndependentTag.Create(
        doc,
        view.Id,
        ref,
        bool(leader),
        tag_mode,
        orientation,
        head_pt,
    )

    # Enforce the requested tag type after creation.
    if tag_type_symbol is not None:
        if not tag_type_symbol.IsActive:
            tag_type_symbol.Activate()
            doc.Regenerate()
        tag.ChangeTypeId(tag_type_symbol.Id)

    return tag
