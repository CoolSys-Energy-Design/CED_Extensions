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
import tracking_service
if previous_snippets is None:
    del sys.modules["Snippets"]
else:
    sys.modules["Snippets"] = previous_snippets


class LinkedSourceScanGateTests(unittest.TestCase):
    def setUp(self):
        self._evaluate = tracking_service.source_service.evaluate_source_for_scan
        self._collect = tracking_service._collect_current_map

    def tearDown(self):
        tracking_service.source_service.evaluate_source_for_scan = self._evaluate
        tracking_service._collect_current_map = self._collect

    def test_link_unavailable_preserves_existing_record_states(self):
        tracking_set = models.new_tracking_set(
            "Linked Equipment",
            {"source_type": models.SOURCE_LINK},
            {"id": 1, "name": "Mechanical Equipment"},
            [],
        )
        tracking_set["set_id"] = "set-1"
        record = models.new_tracked_element({
            "persistent_id": "link:l:e",
            "source_element_unique_id": "e",
            "metadata": {"friendly_name": "AHU-1", "workset_id": 42},
            "properties": {},
            "location": None,
        }, baseline=True)
        record["state"] = models.ELEMENT_TRACKED
        tracking_set["elements"] = {"link:l:e": record}
        store = {"tracking_sets": [tracking_set]}

        tracking_service.source_service.evaluate_source_for_scan = (
            lambda _document, _set: {
                "available": False,
                "set_status": models.SET_LINK_WORKSETS_UNAVAILABLE,
                "message": "Scan skipped: linked workset is closed.",
            }
        )
        tracking_service._collect_current_map = lambda *_args, **_kwargs: self.fail(
            "Partial source collection must not run when the link is unavailable."
        )

        updated, scanned = tracking_service.scan_tracking_set(
            object(), store, "set-1"
        )
        updated_record = updated["tracking_sets"][0]["elements"]["link:l:e"]
        self.assertEqual(models.SET_LINK_WORKSETS_UNAVAILABLE, scanned["status"])
        self.assertEqual(models.ELEMENT_TRACKED, updated_record["state"])
        self.assertEqual("AHU-1", updated_record["metadata"]["friendly_name"])


if __name__ == "__main__":
    unittest.main()
