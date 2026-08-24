# -*- coding: utf-8 -*-
from __future__ import print_function

import io
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = os.path.dirname(HERE)
WORKSPACE_ROOT = os.path.abspath(os.path.join(BUNDLE_DIR, "..", "..", "..", ".."))
REVIT_HELPERS = os.path.join(WORKSPACE_ROOT, "CEDLib.lib", "Snippets", "revit_helpers.py")


def _read(path):
    with io.open(path, "r", encoding="utf-8") as stream:
        return stream.read()


class RevitCompatibilityContractTests(unittest.TestCase):
    def test_element_id_uses_shared_64_bit_first_compatibility_helpers(self):
        helpers = _read(REVIT_HELPERS)
        monitor_sources = "\n".join([
            _read(os.path.join(BUNDLE_DIR, name))
            for name in os.listdir(BUNDLE_DIR)
            if name.endswith(".py")
        ])

        self.assertLess(
            helpers.index('getattr(item, "Value")'),
            helpers.index('getattr(item, "IntegerValue")'),
        )
        self.assertLess(
            helpers.index("DB.ElementId(Int64(numeric))"),
            helpers.index("DB.ElementId(numeric)"),
        )
        self.assertNotIn("DB.ElementId(", monitor_sources)

    def test_parameter_access_avoids_random_and_unsupported_overloads(self):
        source = _read(os.path.join(BUNDLE_DIR, "parameter_service.py"))

        self.assertNotIn(".LookupParameter(", source)
        self.assertIn("owner.GetParameters(requested)", source)
        self.assertIn('owner.get_Parameter(System.Guid(', source)
        self.assertIn("owner.get_Parameter(built_in)", source)
        self.assertNotIn("owner.get_Parameter(\n                revit_helpers.elementid_from_value", source)

    def test_storage_uses_cross_version_extensible_and_global_parameter_contracts(self):
        source = _read(os.path.join(BUNDLE_DIR, "storage_service.py"))

        self.assertIn("from Autodesk.Revit.DB.ExtensibleStorage import (", source)
        self.assertIn("return spec_type.String.MultilineText", source)
        self.assertIn("DB.ParameterType.MultilineText", source)
        self.assertNotIn(".OfClass(DB.DataStorage)", source)
        self.assertNotIn("DB.DataStorage.Create", source)

    def test_every_monitor_python_source_declares_utf8(self):
        for name in os.listdir(BUNDLE_DIR):
            if not name.endswith(".py"):
                continue
            first_line = _read(os.path.join(BUNDLE_DIR, name)).splitlines()[0]
            self.assertEqual("# -*- coding: utf-8 -*-", first_line, name)


if __name__ == "__main__":
    unittest.main()
