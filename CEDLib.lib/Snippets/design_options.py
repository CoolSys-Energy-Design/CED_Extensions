# -*- coding: utf-8 -*-
"""Shared helpers for excluding elements owned by Revit design options."""

from pyrevit import DB

from Snippets import revit_helpers


def main_model_filter():
    """Return a Revit filter that passes only elements in the main model."""
    return DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)


def apply_main_model_filter(collector):
    """Apply the main-model design-option filter to a collector."""
    return collector.WherePasses(main_model_filter())


def is_main_model_element(element):
    """Return True only when an element can be confirmed to be in the main model."""
    if element is None:
        return False

    try:
        return bool(main_model_filter().PassesFilter(element))
    except Exception:
        pass

    try:
        return element.DesignOption is None
    except Exception:
        pass

    try:
        parameter = element.get_Parameter(DB.BuiltInParameter.DESIGN_OPTION_ID)
        if parameter and parameter.HasValue:
            option_id = revit_helpers.get_elementid_value(parameter.AsElementId())
            return option_id <= 0
    except Exception:
        pass

    # Fail closed. Electrical tools must never act on an element whose design
    # option ownership cannot be established.
    return False


def filter_main_model_elements(elements):
    """Return only confirmed main-model elements from an iterable."""
    return [element for element in list(elements or []) if is_main_model_element(element)]
