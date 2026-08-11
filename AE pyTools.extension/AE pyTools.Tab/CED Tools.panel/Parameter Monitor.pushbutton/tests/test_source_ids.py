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

snippets_stub = types.ModuleType("Snippets")
revit_helpers_stub = types.ModuleType("revit_helpers")
revit_helpers_stub.get_elementid_value = lambda value: value
revit_helpers_stub.elementid_from_value = lambda value: value
snippets_stub.revit_helpers = revit_helpers_stub
previous_snippets = sys.modules.get("Snippets")
sys.modules["Snippets"] = snippets_stub
import source_service
if previous_snippets is None:
    del sys.modules["Snippets"]
else:
    sys.modules["Snippets"] = previous_snippets


class PersistentIdTests(unittest.TestCase):
    def test_parse_host_id(self):
        self.assertEqual(
            source_service.parse_persistent_id("host:abc-123"),
            ("host", None, "abc-123"),
        )

    def test_parse_link_id(self):
        # Element UniqueIds contain no colons beyond our two separators,
        # but split-limit keeps any that appear intact.
        self.assertEqual(
            source_service.parse_persistent_id("link:link-uid-1:elem-uid-2"),
            ("link", "link-uid-1", "elem-uid-2"),
        )

    def test_parse_bare_id_defaults_to_host(self):
        self.assertEqual(
            source_service.parse_persistent_id("raw-uid"),
            ("host", None, "raw-uid"),
        )

    def test_unique_id_extraction(self):
        self.assertEqual(
            source_service.unique_id_from_persistent_id("link:l1:e2"), "e2"
        )
        self.assertEqual(
            source_service.unique_id_from_persistent_id("host:e3"), "e3"
        )


if __name__ == "__main__":
    unittest.main()
