# -*- coding: utf-8 -*-
"""Focused tests for IronPython-safe FamilySymbol name resolution."""

from pathlib import Path
import importlib.util
import sys
import types


ROOT = Path(__file__).resolve().parents[2]


class _ElementId(object):
    def __init__(self, value):
        self.Value = int(value)
        self.IntegerValue = int(value)

    def __eq__(self, other):
        return isinstance(other, _ElementId) and self.Value == other.Value


_ElementId.InvalidElementId = _ElementId(-1)


class _NameDescriptor(object):
    def __get__(self, element, unused_owner=None):
        if element is None:
            return self
        return getattr(element, "descriptor_name", None)


class _Element(object):
    Name = _NameDescriptor()


class _FamilySymbol(_Element):
    def __init__(self, descriptor_name=None, symbol_name=None, element_id=10):
        self.descriptor_name = descriptor_name
        self._symbol_name = symbol_name
        self.Id = _ElementId(element_id)
        self.Document = None

    def get_Parameter(self, parameter_name):
        if parameter_name == "symbol_name":
            return _Parameter(string_value=self._symbol_name, value_string=self._symbol_name)
        return None


class _FamilyInstance(_Element):
    def __init__(self, type_id=10, value_string=None, element_id=20):
        self.descriptor_name = None
        self._type_id = _ElementId(type_id)
        self._value_string = value_string
        self.Id = _ElementId(element_id)
        self.Document = None

    def GetTypeId(self):
        return self._type_id

    def get_Parameter(self, parameter_name):
        if parameter_name == "elem_type":
            return _Parameter(value_string=self._value_string, element_id=self._type_id)
        return None


class _Parameter(object):
    def __init__(self, string_value=None, value_string=None, element_id=None):
        self.string_value = string_value
        self.value_string = value_string
        self.element_id = element_id

    def AsString(self):
        return self.string_value

    def AsValueString(self):
        return self.value_string

    def AsElementId(self):
        return self.element_id


class _Doc(object):
    def __init__(self, elements):
        self.elements = elements

    def GetElement(self, element_id):
        return self.elements.get(element_id.Value)


def _load_helpers():
    db = types.ModuleType("Autodesk.Revit.DB")
    db.Element = _Element
    db.FamilySymbol = _FamilySymbol
    db.ElementId = _ElementId
    db.BuiltInParameter = types.SimpleNamespace(
        SYMBOL_NAME_PARAM="symbol_name",
        ELEM_TYPE_PARAM="elem_type",
    )

    autodesk = types.ModuleType("Autodesk")
    revit_package = types.ModuleType("Autodesk.Revit")
    autodesk.Revit = revit_package
    revit_package.DB = db

    system = types.ModuleType("System")
    system.Int64 = int

    saved = {}
    for name, module in {
        "Autodesk": autodesk,
        "Autodesk.Revit": revit_package,
        "Autodesk.Revit.DB": db,
        "System": system,
    }.items():
        saved[name] = sys.modules.get(name)
        sys.modules[name] = module

    helper_path = ROOT / "CEDLib.lib" / "Snippets" / "revit_helpers.py"
    spec = importlib.util.spec_from_file_location("test_revit_helpers_under_test", str(helper_path))
    helpers = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helpers)

    for name, previous in saved.items():
        if previous is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = previous
    return helpers


HELPERS = _load_helpers()


def test_family_symbol_uses_clr_name_descriptor_first():
    symbol = _FamilySymbol(descriptor_name="Descriptor Type", symbol_name="Parameter Type")
    assert HELPERS.get_family_symbol_name(symbol) == "Descriptor Type"


def test_family_symbol_uses_symbol_name_parameter_fallback():
    symbol = _FamilySymbol(descriptor_name=None, symbol_name="Parameter Type")
    assert HELPERS.get_family_symbol_name(symbol) == "Parameter Type"


def test_instance_resolves_family_symbol_then_symbol_parameter():
    symbol = _FamilySymbol(descriptor_name=None, symbol_name="Resolved Type")
    doc = _Doc({10: symbol})
    instance = _FamilyInstance(type_id=10)
    instance.Document = doc
    assert HELPERS.get_family_symbol_name(instance, doc=doc) == "Resolved Type"


def test_instance_uses_element_type_value_string_when_type_unresolved():
    doc = _Doc({})
    instance = _FamilyInstance(type_id=10, value_string="Instance Type Value")
    instance.Document = doc
    assert HELPERS.get_family_symbol_name(instance, doc=doc) == "Instance Type Value"


def test_instance_uses_element_type_element_id_when_value_string_is_null():
    symbol = _FamilySymbol(descriptor_name="ID Resolved Type", symbol_name=None)
    doc = _Doc({10: symbol})
    instance = _FamilyInstance(type_id=10, value_string=None)
    instance.Document = doc
    assert HELPERS.get_family_symbol_name(instance, doc=doc) == "ID Resolved Type"
