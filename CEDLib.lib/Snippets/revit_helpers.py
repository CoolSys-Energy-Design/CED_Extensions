# -*- coding: utf-8 -*-
"""Small Revit API helper utilities shared by circuit tools."""

import Autodesk.Revit.DB as DB
from System import Int64


PARAMETER_DOUBLE_TOLERANCE = 1e-9


def get_elementid_value(item, default=0):
    """Return an ElementId numeric value across Revit API versions."""
    if item is None:
        return int(default or 0)
    try:
        return int(getattr(item, "Value"))
    except Exception:
        pass
    try:
        return int(getattr(item, "IntegerValue"))
    except Exception:
        return int(default or 0)


def coerce_elementid_value(value, default=0):
    """Normalize a native ElementId or boundary numeric value to an int."""
    if value is None:
        return int(default or 0)
    try:
        if isinstance(value, DB.ElementId):
            return get_elementid_value(value, default=default)
    except Exception:
        pass
    try:
        if hasattr(value, "Value") or hasattr(value, "IntegerValue"):
            return get_elementid_value(value, default=default)
    except Exception:
        pass
    try:
        return int(value)
    except Exception:
        return int(default or 0)


def elementid_from_value(value):
    """Create a Revit ElementId from an integer-like value."""
    numeric = int(value or 0)
    try:
        return DB.ElementId(Int64(numeric))
    except Exception:
        return DB.ElementId(numeric)


def get_type_element(element, doc=None):
    """Return an element type for an instance, or None when unavailable."""
    if element is None:
        return None
    if doc is None:
        try:
            doc = element.Document
        except Exception:
            doc = None
    if doc is None:
        return None
    try:
        type_id = element.GetTypeId()
    except Exception:
        return None
    if not type_id or type_id == DB.ElementId.InvalidElementId:
        return None
    try:
        return doc.GetElement(type_id)
    except Exception:
        return None


def get_parameter(element, name, include_type=False, case_insensitive=True, doc=None):
    """Return a parameter by name from instance, optionally falling back to type."""
    if element is None or not name:
        return None
    target_name = str(name)

    def _lookup(owner):
        if owner is None:
            return None
        try:
            param = owner.LookupParameter(target_name)
            if param is not None:
                return param
        except Exception:
            pass
        if not case_insensitive:
            return None
        try:
            for candidate in owner.Parameters:
                try:
                    definition = candidate.Definition
                    if definition and str(definition.Name).strip().lower() == target_name.strip().lower():
                        return candidate
                except Exception:
                    continue
        except Exception:
            pass
        return None

    param = _lookup(element)
    if param is not None or not include_type:
        return param
    return _lookup(get_type_element(element, doc=doc))


def get_parameter_value(parameter, default=None):
    """Return a best-effort native Python value for a Revit Parameter."""
    if parameter is None:
        return default
    try:
        storage_type = parameter.StorageType
    except Exception:
        return default
    try:
        if storage_type == DB.StorageType.String:
            value = parameter.AsString()
            if value is None:
                value = parameter.AsValueString()
            return value if value is not None else default
        if storage_type == DB.StorageType.Integer:
            return parameter.AsInteger()
        if storage_type == DB.StorageType.Double:
            return parameter.AsDouble()
        if storage_type == DB.StorageType.ElementId:
            return parameter.AsElementId()
    except Exception:
        return default
    return default


def parameter_matches_value(parameter, desired, double_tolerance=PARAMETER_DOUBLE_TOLERANCE):
    """Return whether a parameter already contains the desired value.

    Comparison is performed in the parameter's native storage type.  None
    represents the normal cleared value for that storage type.  This helper is
    intentionally read-only; callers can use the same parameter instance for
    a subsequent Set when the value does not match.
    """
    if parameter is None:
        return False
    try:
        storage_type = parameter.StorageType
    except Exception:
        return False

    # Revit returns the storage-type default from AsInteger/AsDouble even
    # when a newly bound parameter has never been assigned.  That is not an
    # equal value for an explicit write such as a first-calculation Yes/No
    # initialization.  None still means "leave an already-unset value clear."
    try:
        if not bool(parameter.HasValue):
            return desired is None
    except Exception:
        pass

    try:
        if storage_type == DB.StorageType.String:
            current = parameter.AsString()
            current = "" if current is None else str(current)
            expected = "" if desired is None else str(desired)
            return current == expected

        if storage_type == DB.StorageType.Integer:
            if isinstance(desired, DB.ElementId):
                return False
            expected = 0 if desired is None else int(desired)
            return parameter.AsInteger() == expected

        if storage_type == DB.StorageType.Double:
            if isinstance(desired, DB.ElementId):
                return False
            expected = 0.0 if desired is None else float(desired)
            current = parameter.AsDouble()
            if current is None:
                return expected == 0.0
            difference = abs(float(current) - expected)
            scale = max(1.0, abs(float(current)), abs(expected))
            return difference <= max(float(double_tolerance), float(double_tolerance) * scale)

        if storage_type == DB.StorageType.ElementId:
            if desired is None:
                expected = get_elementid_value(DB.ElementId.InvalidElementId, default=-1)
            elif isinstance(desired, DB.ElementId):
                expected = get_elementid_value(desired)
            elif isinstance(desired, (int, float)):
                expected = int(desired)
            else:
                # Numeric values are useful for comparison at a DTO boundary,
                # but are never passed to ElementId.Set by this helper.
                expected = get_elementid_value(desired)
            return get_elementid_value(parameter.AsElementId()) == expected
    except Exception:
        return False

    return False


def set_parameter_if_changed(parameter, desired, double_tolerance=PARAMETER_DOUBLE_TOLERANCE):
    """Set a Revit parameter only when its native value differs.

    Returns True when Set was attempted successfully and False for a no-op,
    unsupported value, or failed write. ElementId parameters always receive a
    native DB.ElementId instance.
    """
    if parameter is None or parameter_matches_value(parameter, desired, double_tolerance):
        return False

    try:
        storage_type = parameter.StorageType
        value = desired
        if storage_type == DB.StorageType.String:
            value = "" if desired is None else str(desired)
        elif storage_type == DB.StorageType.Integer:
            if isinstance(desired, DB.ElementId):
                return False
            value = 0 if desired is None else int(desired)
        elif storage_type == DB.StorageType.Double:
            if isinstance(desired, DB.ElementId):
                return False
            value = 0.0 if desired is None else float(desired)
        elif storage_type == DB.StorageType.ElementId:
            if desired is None:
                value = DB.ElementId.InvalidElementId
            elif isinstance(desired, DB.ElementId):
                value = desired
            else:
                return False
        else:
            return False
        parameter.Set(value)
        return True
    except Exception:
        return False


def get_parameter_text(element, name, include_type=False, case_insensitive=True, doc=None, default=""):
    """Return a parameter value as text from instance or type."""
    param = get_parameter(
        element,
        name,
        include_type=include_type,
        case_insensitive=case_insensitive,
        doc=doc,
    )
    value = get_parameter_value(param, default=None)
    if value is None:
        return default
    try:
        return str(value)
    except Exception:
        return default


def _is_family_symbol(element):
    """Return whether an element is a FamilySymbol without assuming the class exists."""
    family_symbol_type = getattr(DB, "FamilySymbol", None)
    if family_symbol_type is None or element is None:
        return False
    try:
        return isinstance(element, family_symbol_type)
    except Exception:
        return False


def _element_name_descriptor_value(element):
    """Read Element.Name through the CLR descriptor, bypassing IronPython Name binding."""
    if element is None:
        return ""
    try:
        name = DB.Element.Name.__get__(element)
        if name:
            return str(name)
    except Exception:
        pass
    return ""


def _builtin_parameter(element, parameter_name):
    """Return a built-in parameter when the current Revit version exposes it."""
    if element is None:
        return None
    try:
        built_in = getattr(DB.BuiltInParameter, parameter_name)
    except Exception:
        return None
    try:
        return element.get_Parameter(built_in)
    except Exception:
        return None


def _parameter_display_text(parameter):
    """Read a parameter as display text across string/value-string storage."""
    if parameter is None:
        return ""
    for method_name in ("AsString", "AsValueString"):
        try:
            value = getattr(parameter, method_name)()
            if value:
                return str(value)
        except Exception:
            pass
    return ""


def _family_symbol_type_name(family_symbol):
    """Resolve a FamilySymbol type name using Revit 2024-safe access order."""
    if family_symbol is None:
        return ""
    name = _element_name_descriptor_value(family_symbol)
    if name:
        return name
    # SYMBOL_NAME_PARAM is a FamilySymbol fallback.  It is not the instance
    # type-reference parameter and should not be queried on FamilyInstance.
    return _parameter_display_text(
        _builtin_parameter(family_symbol, "SYMBOL_NAME_PARAM")
    )


def _instance_type_parameter_name(instance, doc=None):
    """Resolve an instance type through ELEM_TYPE_PARAM when needed."""
    parameter = _builtin_parameter(instance, "ELEM_TYPE_PARAM")
    if parameter is None:
        return ""

    # Revit may expose the type reference directly as a formatted value.
    name = _parameter_display_text(parameter)
    if name:
        return name

    # Otherwise it is an ElementId.  Keep it native while resolving it in the
    # document, then use the FamilySymbol resolver above.
    try:
        type_id = parameter.AsElementId()
    except Exception:
        type_id = None
    if type_id is None or doc is None:
        return ""
    try:
        type_element = doc.GetElement(type_id)
    except Exception:
        type_element = None
    if _is_family_symbol(type_element):
        return _family_symbol_type_name(type_element)
    return _element_name_descriptor_value(type_element)


def get_family_symbol_name(element, doc=None, fallback=""):
    """Return a FamilySymbol type name using IronPython-safe Revit access."""
    if element is None:
        return fallback

    is_symbol = _is_family_symbol(element)
    family_symbol = element if is_symbol else get_type_element(element, doc=doc)
    if family_symbol is not None:
        name = _family_symbol_type_name(family_symbol)
        if name:
            return name

    # ELEM_TYPE_PARAM belongs to instances, not FamilySymbols.  Use it only
    # after the resolved FamilySymbol path has failed.
    if not is_symbol:
        name = _instance_type_parameter_name(element, doc=doc)
        if name:
            return name

    # Preserve a useful fallback for non-FamilySymbol ElementType callers
    # (for example tag types) without relying on IronPython's .Name property.
    name = _element_name_descriptor_value(element)
    return name or fallback
