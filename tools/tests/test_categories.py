# -*- coding: utf-8 -*-
"""Focused tests for shared Revit category classification."""

from pathlib import Path
import importlib.util
import sys
import types


ROOT = Path(__file__).resolve().parents[2]


class _ElementId(object):
    def __init__(self, value):
        self.Value = int(value)
        self.IntegerValue = int(value)


_ElementId.InvalidElementId = _ElementId(-1)


class _BuiltInCategory(object):
    INVALID = -1
    OST_ElectricalEquipment = -2001040
    OST_LightingDevices = -2001050


class _Category(object):
    def __init__(self, built_in_category):
        self.BuiltInCategory = built_in_category


class _Element(object):
    def __init__(self, category):
        self.Category = category


def _load_categories():
    db = types.ModuleType("Autodesk.Revit.DB")
    db.ElementId = _ElementId
    db.BuiltInCategory = _BuiltInCategory

    pyrevit = types.ModuleType("pyrevit")
    pyrevit.DB = db

    snippets = types.ModuleType("Snippets")
    helpers = types.ModuleType("Snippets.revit_helpers")
    helpers.get_elementid_value = lambda item, default=0: int(
        getattr(item, "Value", getattr(item, "IntegerValue", default))
    )
    helpers.elementid_from_value = lambda value: _ElementId(value)
    snippets.revit_helpers = helpers

    saved = {}
    for name, module in {
        "pyrevit": pyrevit,
        "Snippets": snippets,
        "Snippets.revit_helpers": helpers,
    }.items():
        saved[name] = sys.modules.get(name)
        sys.modules[name] = module

    category_path = ROOT / "CEDLib.lib" / "Snippets" / "categories.py"
    spec = importlib.util.spec_from_file_location(
        "test_categories_under_test",
        str(category_path),
    )
    categories = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(categories)

    for name, previous in saved.items():
        if previous is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = previous
    return categories


CATEGORIES = _load_categories()


def test_get_built_in_category_from_category():
    category = _Category(_BuiltInCategory.OST_LightingDevices)
    assert CATEGORIES.get_built_in_category(category) == _BuiltInCategory.OST_LightingDevices


def test_get_built_in_category_from_element():
    category = _Category(_BuiltInCategory.OST_ElectricalEquipment)
    element = _Element(category)
    assert CATEGORIES.get_built_in_category(element) == _BuiltInCategory.OST_ElectricalEquipment


def test_get_built_in_category_rejects_invalid_or_missing_category():
    assert CATEGORIES.get_built_in_category(_Category(_BuiltInCategory.INVALID)) is None
    assert CATEGORIES.get_built_in_category(_Element(None)) is None


if __name__ == "__main__":
    tests = sorted(
        (name, value)
        for name, value in globals().items()
        if name.startswith("test_") and callable(value)
    )
    for name, test in tests:
        test()
        print("PASS {}".format(name))
