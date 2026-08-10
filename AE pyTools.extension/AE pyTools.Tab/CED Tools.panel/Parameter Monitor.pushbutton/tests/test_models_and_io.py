# -*- coding: utf-8 -*-
from __future__ import print_function

import json
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = os.path.dirname(HERE)
if BUNDLE_DIR not in sys.path:
    sys.path.insert(0, BUNDLE_DIR)

import models
import set_io
import storage_service


class ModelsAndIoTests(unittest.TestCase):
    def _set(self):
        tracking_set = models.new_tracking_set(
            "Equipment",
            {"source_type": "link", "display_name": "MEP Link", "link_instance_unique_id": "link-uid"},
            {"id": -2001140, "name": "Mechanical Equipment"},
            [{
                "key": "shared:guid:instance",
                "name": "Asset ID",
                "scope": "instance",
                "shared_guid": "guid",
            }],
        )
        tracking_set["last_check"] = "2026-08-09T12:00:00Z"
        tracking_set["untracked_ids"] = ["link:link-uid:element-uid"]
        tracking_set["elements"] = {
            "link:link-uid:element-2": {
                "persistent_id": "link:link-uid:element-2",
                "accepted_properties": {"shared:guid:instance": {"raw": "A"}},
                "relationship": {"device_unique_id": "device-uid"},
            }
        }
        return tracking_set

    def test_definition_export_excludes_project_scan_state(self):
        document = set_io.build_export_document([self._set()])
        encoded = json.dumps(document)
        self.assertIn("tracked_properties", encoded)
        self.assertNotIn("accepted_properties", encoded)
        self.assertNotIn("untracked_ids", encoded)
        self.assertNotIn("device_unique_id", encoded)
        self.assertNotIn("last_check", encoded)
        self.assertNotIn("link_instance_unique_id", encoded)

    def test_definition_round_trip(self):
        definitions = set_io.loads(set_io.dumps([self._set()]))
        self.assertEqual(1, len(definitions))
        self.assertEqual("Equipment", definitions[0]["name"])
        self.assertEqual(1, len(definitions[0]["tracked_properties"]))

    def test_definition_version_is_enforced(self):
        with self.assertRaises(set_io.DefinitionFormatError):
            set_io.loads('{"schema_version":99,"tracking_sets":[]}')

    def test_store_payload_round_trip(self):
        store = models.new_project_store({"title": "Test"})
        store["tracking_sets"].append(self._set())
        text = storage_service.serialize_payload(store)
        restored = storage_service.deserialize_payload(text)
        self.assertEqual(models.PAYLOAD_SCHEMA_VERSION, restored["payload_schema_version"])
        self.assertEqual("Equipment", restored["tracking_sets"][0]["name"])

    def test_corrupt_payload_does_not_silently_reset(self):
        with self.assertRaises(storage_service.StorageCorruptionError):
            storage_service.deserialize_payload("{not json")

    def test_legacy_metadata_becomes_family_type_baseline(self):
        store = models.new_project_store({"title": "Legacy"})
        tracking_set = self._set()
        record = tracking_set["elements"]["link:link-uid:element-2"]
        record["metadata"] = {"family_type": "Legacy Family : Legacy Type"}
        store["tracking_sets"].append(tracking_set)

        migrated = models.migrate_project_store(store)
        migrated_record = migrated["tracking_sets"][0]["elements"]["link:link-uid:element-2"]
        self.assertEqual(
            migrated_record["metadata"], migrated_record["accepted_metadata"]
        )
        self.assertEqual(
            migrated_record["metadata"], migrated_record["current_metadata"]
        )

    def test_future_payload_is_rejected(self):
        with self.assertRaises(models.PayloadMigrationError):
            models.migrate_project_store({
                "payload_schema_version": models.PAYLOAD_SCHEMA_VERSION + 1,
                "tracking_sets": [],
            })


if __name__ == "__main__":
    unittest.main()
