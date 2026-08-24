# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import sys
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = os.path.dirname(HERE)
if BUNDLE_DIR not in sys.path:
    sys.path.insert(0, BUNDLE_DIR)


class _ElementId(object):
    def __init__(self, value):
        self.Value = int(value)


class _RevitHelpers(object):
    @staticmethod
    def get_elementid_value(value, default=0):
        try:
            return int(value.Value)
        except Exception:
            try:
                return int(value.IntegerValue)
            except Exception:
                return int(default or 0)

    @staticmethod
    def get_type_element(element, doc=None):
        return getattr(element, "type_element", None)


snippets_stub = types.ModuleType("Snippets")
snippets_stub.revit_helpers = _RevitHelpers
previous_snippets = sys.modules.get("Snippets")
sys.modules["Snippets"] = snippets_stub
import parameter_service
if previous_snippets is None:
    del sys.modules["Snippets"]
else:
    sys.modules["Snippets"] = previous_snippets


class _StorageType(object):
    String = "String"
    Integer = "Integer"
    Double = "Double"
    ElementId = "ElementId"


class _LabelUtils(object):
    @staticmethod
    def GetLabelForSpec(data_type):
        labels = {
            "autodesk.spec.aec:length-2.0.0": "Length",
            "spec:number": "Number",
            "spec:text": "Text",
        }
        return labels[data_type.TypeId]


class _DB(object):
    StorageType = _StorageType
    LabelUtils = _LabelUtils


class _DataType(object):
    def __init__(self, type_id):
        self.TypeId = type_id


class _Definition(object):
    def __init__(self, name, spec_type):
        self.Name = name
        self._spec_type = spec_type

    def GetDataType(self):
        return _DataType(self._spec_type)


class _Parameter(object):
    def __init__(self, name, parameter_id, storage_type="Double", spec_type="spec:number"):
        self.Definition = _Definition(name, spec_type)
        self.Id = _ElementId(parameter_id)
        self.StorageType = storage_type
        self.IsShared = False


class _Owner(object):
    def __init__(self, parameters):
        self.Parameters = list(parameters)
        self.get_parameters_calls = []
        self.get_parameter_calls = []

    def GetParameters(self, name):
        self.get_parameters_calls.append(name)
        return [
            item for item in self.Parameters
            if item.Definition.Name == name
        ]

    def get_Parameter(self, identity):
        self.get_parameter_calls.append(identity)
        return None


class ParameterIdentityTests(unittest.TestCase):
    def setUp(self):
        self.previous_db = parameter_service.DB
        parameter_service.DB = _DB
        parameter_service._SPEC_LABEL_CACHE.clear()

    def tearDown(self):
        parameter_service.DB = self.previous_db

    def test_project_parameter_uses_parameter_id_among_duplicate_names(self):
        first = _Parameter("MCA", 101)
        expected = _Parameter("MCA", 102)
        third = _Parameter("MCA", 103, storage_type="String", spec_type="spec:text")
        owner = _Owner([first, expected, third])
        descriptor = parameter_service.descriptor_from_parameter(expected, "instance")

        actual = parameter_service.resolve_parameter(owner, descriptor)

        self.assertIs(expected, actual)
        self.assertEqual(["MCA"], owner.get_parameters_calls)
        self.assertEqual([], owner.get_parameter_calls)

    def test_descriptor_uses_revit_label_for_spec_without_changing_identity(self):
        parameter = _Parameter(
            "Width",
            101,
            storage_type="Double",
            spec_type="autodesk.spec.aec:length-2.0.0",
        )

        descriptor = parameter_service.descriptor_from_parameter(parameter, "instance")

        self.assertEqual("autodesk.spec.aec:length-2.0.0", descriptor["spec_type"])
        self.assertNotIn("spec_label", descriptor)
        self.assertEqual(
            "Length",
            parameter_service.spec_type_label(descriptor["spec_type"]),
        )
        self.assertEqual("project:101:instance", descriptor["key"])

    def test_spec_type_label_humanizes_forge_type_id_as_fallback(self):
        self.assertEqual(
            "Length",
            parameter_service.spec_type_label(
                "autodesk.spec.aec:length-2.0.0"
            ),
        )

    def test_unsupported_or_object_repr_spec_type_displays_other(self):
        self.assertEqual("Other", parameter_service.spec_type_label(""))
        self.assertEqual(
            "Other",
            parameter_service.spec_type_label(
                "<Autodesk.Revit.DB.ForgeTypeId object at 0x000000000001>"
            ),
        )
        descriptor = parameter_service.descriptor_from_parameter(
            _Parameter("Phase", 201, storage_type="ElementId", spec_type=""),
            "instance",
        )
        self.assertEqual("", descriptor["spec_type"])
        self.assertEqual("Other", parameter_service.spec_type_label(descriptor["spec_type"]))

    def test_discovery_keeps_same_name_project_parameters_distinct(self):
        element = _Owner([_Parameter("MCA", 101), _Parameter("MCA", 102)])

        descriptors = parameter_service.discover_parameters([element], None)

        self.assertEqual(2, len(descriptors))
        self.assertEqual(
            set(["project:101:instance", "project:102:instance"]),
            set([item["key"] for item in descriptors]),
        )

    def test_legacy_name_resolution_requires_one_matching_value_shape(self):
        number = _Parameter("MCA", 101, storage_type="Double", spec_type="spec:number")
        text = _Parameter("MCA", 102, storage_type="String", spec_type="spec:text")
        owner = _Owner([number, text])
        descriptor = {
            "identity_kind": "name",
            "name": "MCA",
            "scope": "instance",
            "storage_type": "string",
            "spec_type": "spec:text",
        }

        self.assertIs(text, parameter_service.resolve_parameter(owner, descriptor))

        owner.Parameters.append(
            _Parameter("MCA", 103, storage_type="String", spec_type="spec:text")
        )
        self.assertIsNone(parameter_service.resolve_parameter(owner, descriptor))

    def test_import_mapping_refuses_ambiguous_portable_name(self):
        imported = {
            "identity_kind": "name",
            "name": "MCA",
            "scope": "instance",
            "storage_type": "double",
            "spec_type": "spec:number",
        }
        available = [
            dict(imported, key="project:101:instance", identity_kind="project", parameter_id=101),
            dict(imported, key="project:102:instance", identity_kind="project", parameter_id=102),
        ]

        self.assertIsNone(
            parameter_service.find_matching_descriptor(imported, available)
        )

    def test_shared_identity_never_maps_to_same_name_project_parameter(self):
        imported = {
            "identity_kind": "shared",
            "shared_guid": "11111111-1111-1111-1111-111111111111",
            "name": "MCA",
            "scope": "instance",
            "storage_type": "double",
            "spec_type": "spec:number",
        }
        project = {
            "identity_kind": "project",
            "parameter_id": 101,
            "name": "MCA",
            "scope": "instance",
            "storage_type": "double",
            "spec_type": "spec:number",
        }

        self.assertFalse(parameter_service.descriptor_matches(imported, project))


if __name__ == "__main__":
    unittest.main()
