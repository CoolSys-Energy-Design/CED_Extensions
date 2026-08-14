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
import models
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


class LinkAvailabilityTests(unittest.TestCase):
    def setUp(self):
        self._resolve_source = source_service.resolve_source
        self._closed_worksets = source_service._closed_user_worksets
        source_service.resolve_source = lambda _document, _source: {
            "available": True,
            "source_document": object(),
            "link_instance": object(),
            "source_state": {},
            "message": "",
        }

    def tearDown(self):
        source_service.resolve_source = self._resolve_source
        source_service._closed_user_worksets = self._closed_worksets

    @staticmethod
    def _set(records):
        return {
            "source": {"source_type": models.SOURCE_LINK},
            "elements": records,
        }

    def test_closed_relevant_link_workset_blocks_scan(self):
        source_service._closed_user_worksets = lambda _document: [
            {"id": 42, "name": "Equipment", "is_open": False},
        ]
        result = source_service.evaluate_source_for_scan(
            object(),
            self._set({
                "link:l:e": {
                    "state": models.ELEMENT_TRACKED,
                    "metadata": {"workset_id": 42},
                },
            }),
        )
        self.assertFalse(result["available"])
        self.assertEqual(
            result["set_status"], models.SET_LINK_WORKSETS_UNAVAILABLE
        )
        self.assertIn("Existing element states were retained", result["message"])

    def test_closed_unrelated_link_workset_does_not_block_complete_set(self):
        source_service._closed_user_worksets = lambda _document: [
            {"id": 42, "name": "Equipment", "is_open": False},
        ]
        result = source_service.evaluate_source_for_scan(
            object(),
            self._set({
                "link:l:e": {
                    "state": models.ELEMENT_TRACKED,
                    "metadata": {"workset_id": 7},
                },
            }),
        )
        self.assertTrue(result["available"])
        self.assertIsNone(result["set_status"])

    def test_legacy_link_set_is_conservatively_blocked_when_workset_closed(self):
        source_service._closed_user_worksets = lambda _document: [
            {"id": 42, "name": "Equipment", "is_open": False},
        ]
        result = source_service.evaluate_source_for_scan(
            object(),
            self._set({
                "link:l:e": {
                    "state": models.ELEMENT_TRACKED,
                    "metadata": {},
                },
            }),
        )
        self.assertFalse(result["available"])
        self.assertEqual(
            result["set_status"], models.SET_LINK_WORKSETS_UNAVAILABLE
        )

    def test_new_link_baseline_requires_all_worksets_open(self):
        source_service._closed_user_worksets = lambda _document: [
            {"id": 42, "name": "Equipment", "is_open": False},
        ]
        result = source_service.evaluate_source_for_scan(
            object(), self._set({}), require_complete=True
        )
        self.assertFalse(result["available"])
        self.assertIn("before creating a complete baseline", result["message"])


if __name__ == "__main__":
    unittest.main()
