# -*- coding: utf-8 -*-
"""
Helpers that enumerate loaded Family / Type combinations for the
Spaces editor dropdowns.

Two indexes:

  * ``build_model_symbol_index(doc)`` — every loadable
    ``FamilySymbol`` whose family belongs to a *model* category
    (``CategoryType.Model``). Used by the Manage Space Profiles LED
    grid's "Label (Family : Type)" combo.

  * ``build_annotation_symbol_index(doc)`` — every loadable
    ``FamilySymbol`` whose family belongs to an *annotation* category
    (``CategoryType.Annotation``: tags, generic annotations,
    keynote tags). Used by the LED Details dialog's annotation grid.

Both return a ``(labels, lookup)`` tuple where ``labels`` is a sorted
list of ``"Family : Type"`` strings and ``lookup`` is a dict from that
same string to ``{family_name, type_name, category_name, symbol_id}``.
The CLR ``FamilySymbol`` is intentionally not stored — callers re-look
it up by id when they actually need to introspect its parameters, so
the index stays cheap and survives in-Revit family loads / unloads.
"""

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    CategoryType,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
)


def _category_type(symbol):
    family = getattr(symbol, "Family", None)
    if family is None:
        return None
    cat = getattr(family, "FamilyCategory", None)
    if cat is None:
        return None
    return getattr(cat, "CategoryType", None)


def _category_name(symbol):
    family = getattr(symbol, "Family", None)
    if family is None:
        return ""
    cat = getattr(family, "FamilyCategory", None)
    if cat is None:
        return ""
    return getattr(cat, "Name", "") or ""


def _id_int(elem_id):
    if elem_id is None:
        return None
    for attr in ("Value", "IntegerValue"):
        try:
            v = getattr(elem_id, attr)
        except Exception:
            v = None
        if v is None:
            continue
        try:
            return int(v)
        except Exception:
            continue
    return None


def _build_index(doc, predicate):
    """Common worker. ``predicate(symbol)`` decides inclusion."""
    if doc is None:
        return [], {}
    lookup = {}
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            if not predicate(sym):
                continue
            family = sym.Family
            family_name = family.Name if family is not None else ""
            type_name = sym.Name or ""
        except Exception:
            continue
        if not family_name or not type_name:
            continue
        label = "{} : {}".format(family_name, type_name)
        if label in lookup:
            continue
        lookup[label] = {
            "family_name": family_name,
            "type_name": type_name,
            "category_name": _category_name(sym),
            "symbol_id": _id_int(sym.Id),
        }
    labels = sorted(lookup.keys(), key=lambda s: s.lower())
    return labels, lookup


def build_model_symbol_index(doc):
    """Sorted ``Family : Type`` labels for *model* FamilySymbols."""
    return _build_index(
        doc,
        lambda s: _category_type(s) == CategoryType.Model,
    )


def build_annotation_symbol_index(doc):
    """Sorted ``Family : Type`` labels for *annotation* FamilySymbols
    (tags, generic annotations, keynote tags)."""
    return _build_index(
        doc,
        lambda s: _category_type(s) == CategoryType.Annotation,
    )


KEYNOTE_FAMILY_NAME = "GA_Keynote Symbol_CED"


def build_tag_symbol_index(doc):
    """Annotation FamilySymbols *excluding* the GA_Keynote Symbol_CED
    family. Used by the Spaces LED editor's Family:Type dropdown when
    the row's kind is ``tag``: keynotes and text notes have their own
    dedicated indexes, so the tag picker shouldn't surface them."""
    def _is_tag_like(sym):
        if _category_type(sym) != CategoryType.Annotation:
            return False
        family = getattr(sym, "Family", None)
        family_name = getattr(family, "Name", "") if family is not None else ""
        return family_name != KEYNOTE_FAMILY_NAME
    return _build_index(doc, _is_tag_like)


def build_keynote_symbol_index(doc):
    """Only types of the keynote family ``GA_Keynote Symbol_CED``.

    Used by the Spaces LED editor's Family:Type dropdown when the row's
    kind is ``keynote`` — locks the user into picking a real keynote
    type so the placement engine's ``_place_keynote`` (which expects a
    GenericAnnotation symbol) always gets a valid family/type pair."""
    def _is_keynote_family(sym):
        family = getattr(sym, "Family", None)
        family_name = getattr(family, "Name", "") if family is not None else ""
        return family_name == KEYNOTE_FAMILY_NAME
    return _build_index(doc, _is_keynote_family)


def build_text_note_type_index(doc):
    """Sorted TextNoteType names (just the type name, no Family prefix).

    TextNoteType is a separate Revit class — it doesn't show up under
    ``FamilySymbol``, so the regular annotation index never includes
    them. Used by the Spaces LED editor's Family:Type dropdown when
    the row's kind is ``text_note``. Returned shape mirrors
    ``_build_index`` for caller consistency: ``(labels, lookup)``
    where each label is the TextNoteType's ``Name`` and the lookup
    maps that name to ``{family_name: "", type_name: <name>, ...}``.
    Empty ``family_name`` is intentional — text notes have no family
    in the FamilySymbol sense; the label is the type name alone.
    """
    if doc is None:
        return [], {}
    try:
        from Autodesk.Revit.DB import TextNoteType
    except Exception:
        return [], {}
    lookup = {}
    for tnt in FilteredElementCollector(doc).OfClass(TextNoteType):
        try:
            name = tnt.Name or ""
        except Exception:
            continue
        if not name or name in lookup:
            continue
        lookup[name] = {
            "family_name": "",
            "type_name": name,
            "category_name": "Text Notes",
            "symbol_id": _id_int(tnt.Id),
        }
    labels = sorted(lookup.keys(), key=lambda s: s.lower())
    return labels, lookup


def find_symbol_by_label(doc, label):
    """Resolve a ``"Family : Type"`` label back to a ``FamilySymbol`` element.

    Returns ``None`` if not found. Uses the doc directly rather than
    a cached lookup so the result is fresh after family loads.
    """
    if doc is None or not label or " : " not in label:
        return None
    family_name, type_name = label.split(" : ", 1)
    family_name = family_name.strip()
    type_name = type_name.strip()
    if not family_name or not type_name:
        return None
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            family = sym.Family
            if family is None:
                continue
            if family.Name == family_name and sym.Name == type_name:
                return sym
        except Exception:
            continue
    return None


def symbol_parameter_defaults(doc, symbol):
    """Return ``{name: value_string}`` for every parameter visible on
    instances of ``symbol`` — type AND instance parameters, with their
    *current value* read out for auto-fill seeding.

    Strategy:

      * Tier 1 — find any existing ``FamilyInstance`` of ``symbol`` and
        read its ``.Parameters``. This is the complete union (type +
        instance + built-in + shared bound to the category) and the
        values reflect the actual placed instance, which is the best
        possible "default" — exactly what the user would get if they
        copied that instance.
      * Tier 2 — type parameters straight off the ``FamilySymbol``,
        with values read from the symbol itself. Used to fill in
        anything the instance didn't surface (rare) and as the sole
        source when no instance exists.

    Names already covered by tier 1 aren't overwritten by tier 2.
    Values come through ``AsValueString()`` first (units-formatted,
    matches the Revit Properties palette display) and fall back to
    ``AsString()``. Empty / unset parameters are kept with an empty
    string value so the user sees the full parameter set.
    """
    out = {}

    if symbol is not None and doc is not None:
        try:
            sym_id = symbol.Id
        except Exception:
            sym_id = None
        if sym_id is not None:
            sid = _id_int(sym_id)
            try:
                for inst in (
                    FilteredElementCollector(doc)
                    .OfClass(FamilyInstance)
                    .WhereElementIsNotElementType()
                ):
                    inst_sym = getattr(inst, "Symbol", None)
                    if inst_sym is None:
                        continue
                    try:
                        if _id_int(inst_sym.Id) != sid:
                            continue
                    except Exception:
                        continue
                    _read_parameters_into(inst, out)
                    if out:
                        break  # one instance is enough
            except Exception:
                pass

    if symbol is not None:
        _read_parameters_into(symbol, out)

    return out


def _read_parameters_into(elem, out):
    """Walk ``elem.Parameters`` and record ``{name: value_string}`` for
    each parameter. Existing entries in ``out`` are preserved (caller
    calls this in priority order)."""
    try:
        params = elem.Parameters
    except Exception:
        return
    for p in params:
        if p is None:
            continue
        try:
            d = getattr(p, "Definition", None)
            name = getattr(d, "Name", None) if d is not None else None
        except Exception:
            name = None
        if not name:
            continue
        if name in out:
            continue
        value = None
        try:
            value = p.AsValueString()
        except Exception:
            value = None
        if value is None:
            try:
                value = p.AsString()
            except Exception:
                value = None
        out[name] = "" if value is None else str(value)
