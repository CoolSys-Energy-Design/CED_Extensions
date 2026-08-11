# -*- coding: utf-8 -*-
from __future__ import print_function

import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = os.path.dirname(HERE)
if BUNDLE_DIR not in sys.path:
    sys.path.insert(0, BUNDLE_DIR)

import models
import sync_logic


def _child(unique_id, parent_unique_id, led_id="LED-1", has_directives=False):
    return {
        "unique_id": unique_id,
        "parent_unique_id": parent_unique_id,
        "led_id": led_id,
        "has_directives": has_directives,
    }


def _snapshot(persistent_id, x=0.0, category="Electrical Fixtures"):
    return {
        "persistent_id": persistent_id,
        "source_element_unique_id": persistent_id.split(":", 1)[-1],
        "metadata": {
            "friendly_name": persistent_id,
            "element_id": 1,
            "category": category,
        },
        "properties": {},
        "location": {
            "state": models.VALUE_VALID,
            "x": x, "y": 0.0, "z": 0.0, "rotation": 0.0,
            "coordinate_system": "source_document_internal",
        },
        "relationship": None,
        "relationship_context": None,
    }


def _entry(persistent_id, role="child", parent=None, led_id="LED-1"):
    return {
        "persistent_id": persistent_id,
        "role": role,
        "parent_persistent_id": parent,
        "linker_meta": {"led_id": led_id} if role == "child" else None,
    }


SOURCE = {"source_type": "host", "display_name": "Host - Test", "document_title": "Test"}


class LedDirectiveIndexTests(unittest.TestCase):
    def test_walks_equipment_and_space_profiles(self):
        data = {
            "equipment_definitions": [{
                "id": "EQ-1",
                "linked_sets": [{
                    "linked_element_definitions": [
                        {"id": "LED-A", "parameters": {"Voltage": 120}},
                        {"id": "LED-B", "parameters": {
                            "Panel": {"parent_parameter": "PanelName"},
                        }},
                    ],
                }],
            }],
            "space_profiles": [{
                "id": "SP-1",
                "linked_sets": [{
                    "linked_element_definitions": [
                        {"id": "LED-C", "parameters": {
                            "Ckt": {"sibling_parameter": "S-1:Circuit"},
                        }},
                    ],
                }],
            }],
        }
        index = sync_logic.led_directive_index(data)
        self.assertEqual(index, {"LED-A": False, "LED-B": True, "LED-C": True})

    def test_handles_yamldotnet_ordereddicts(self):
        # pyRevit's YamlDotNet fallback returns OrderedDicts with
        # string-only scalars; directive detection must still work.
        from collections import OrderedDict
        data = OrderedDict([
            ("equipment_definitions", [OrderedDict([
                ("id", "EQ-1"),
                ("linked_sets", [OrderedDict([
                    ("linked_element_definitions", [OrderedDict([
                        ("id", "LED-D"),
                        ("parameters", OrderedDict([
                            ("Voltage_CED", "120"),
                            ("CKT_Panel_CEDT", OrderedDict([
                                ("parent_parameter", "PanelName"),
                            ])),
                        ])),
                    ])]),
                ])]),
            ])]),
        ])
        self.assertEqual(sync_logic.led_directive_index(data), {"LED-D": True})

    def test_tolerates_malformed_entries(self):
        data = {
            "equipment_definitions": [
                None,
                "junk",
                {"linked_sets": [None, {"linked_element_definitions": [
                    None, {"parameters": "not-a-dict"}, {"id": "", "parameters": {}},
                ]}]},
            ],
        }
        self.assertEqual(sync_logic.led_directive_index(data), {})
        self.assertEqual(sync_logic.led_directive_index(None), {})


class GroupChildrenTests(unittest.TestCase):
    def test_group_children_skips_parentless(self):
        groups = sync_logic.group_children([
            _child("a", "p1"), _child("b", "p1"), _child("c", "p2"),
            _child("d", ""),
        ])
        self.assertEqual(sorted(groups.keys()), ["p1", "p2"])
        self.assertEqual(len(groups["p1"]), 2)


class ApplySyncMembershipTests(unittest.TestCase):
    def _run_once(self, store=None):
        entries = [
            _entry("host:c1", parent="host:p1"),
            _entry("host:p1", role="parent"),
        ]
        snapshots = {
            "host:c1": _snapshot("host:c1", x=1.0),
            "host:p1": _snapshot("host:p1", x=5.0),
        }
        return sync_logic.apply_sync_membership(
            store or models.new_project_store(), entries, snapshots, SOURCE
        )

    def test_creates_explicit_set_with_location_tracking(self):
        result, set_id, report = self._run_once()
        tracking_set = models.find_set(result, set_id)
        self.assertEqual(
            report["sets_created"], ["Element Linker - Electrical Fixtures"]
        )
        self.assertEqual(tracking_set["name"], "Element Linker - Electrical Fixtures")
        self.assertEqual(tracking_set["category"]["name"], "Electrical Fixtures")
        self.assertEqual(tracking_set["membership"], models.MEMBERSHIP_EXPLICIT)
        self.assertEqual(tracking_set["origin"], models.SET_ORIGIN_ELEMENT_LINKER)
        self.assertEqual(report["added"], 2)
        child = tracking_set["elements"]["host:c1"]
        parent = tracking_set["elements"]["host:p1"]
        self.assertTrue(child["track_location"])
        self.assertTrue(parent["track_location"])
        self.assertEqual(child["parent_persistent_id"], "host:p1")
        self.assertIsNone(parent["parent_persistent_id"])
        self.assertEqual(child["linker_meta"]["role"], "child")
        self.assertEqual(parent["linker_meta"]["role"], "parent")

    def test_rerun_is_idempotent_and_preserves_accepted(self):
        result, set_id, _report = self._run_once()
        tracking_set = models.find_set(result, set_id)
        # Simulate an accepted baseline the user cares about.
        tracking_set["elements"]["host:c1"]["accepted_location"]["x"] = 99.0
        result2, set_id2, report2 = self._run_once(store=result)
        self.assertEqual(set_id, set_id2)
        self.assertEqual(report2["added"], 0)
        self.assertEqual(report2["refreshed"], 2)
        self.assertEqual(report2["sets_created"], [])
        self.assertEqual(
            report2["sets_updated"], ["Element Linker - Electrical Fixtures"]
        )
        tracking_set2 = models.find_set(result2, set_id2)
        self.assertEqual(
            tracking_set2["elements"]["host:c1"]["accepted_location"]["x"], 99.0
        )
        self.assertEqual(len(result2["tracking_sets"]), 1)

    def test_stale_clean_removed_stale_dirty_kept(self):
        result, set_id, _report = self._run_once()
        tracking_set = models.find_set(result, set_id)
        # Add two stale sync records: one clean, one with pending changes.
        clean = models.new_tracked_element(_snapshot("host:old1"), baseline=True)
        clean["linker_meta"] = {"role": "child"}
        dirty = models.new_tracked_element(_snapshot("host:old2"), baseline=True)
        dirty["linker_meta"] = {"role": "child"}
        dirty["change_count"] = 1
        dirty["changed_property_keys"] = [models.LOCATION_PROPERTY_KEY]
        tracking_set["elements"]["host:old1"] = clean
        tracking_set["elements"]["host:old2"] = dirty
        result2, set_id2, report2 = self._run_once(store=result)
        tracking_set2 = models.find_set(result2, set_id2)
        self.assertNotIn("host:old1", tracking_set2["elements"])
        self.assertIn("host:old2", tracking_set2["elements"])
        self.assertEqual(report2["stale_removed"], 1)
        self.assertEqual(len(report2["stale_kept"]), 1)

    def test_user_records_without_linker_meta_untouched(self):
        store = models.new_project_store()
        result, set_id, _report = self._run_once(store=store)
        tracking_set = models.find_set(result, set_id)
        manual = models.new_tracked_element(_snapshot("host:manual"), baseline=True)
        tracking_set["elements"]["host:manual"] = manual
        result2, set_id2, _report2 = self._run_once(store=result)
        tracking_set2 = models.find_set(result2, set_id2)
        self.assertIn("host:manual", tracking_set2["elements"])

    def test_untracked_id_is_restored_as_fresh_record(self):
        result, set_id, _report = self._run_once()
        tracking_set = models.find_set(result, set_id)
        del tracking_set["elements"]["host:c1"]
        tracking_set["untracked_ids"] = ["host:c1"]
        result2, set_id2, report2 = self._run_once(store=result)
        tracking_set2 = models.find_set(result2, set_id2)
        self.assertIn("host:c1", tracking_set2["elements"])
        self.assertNotIn("host:c1", tracking_set2["untracked_ids"])
        self.assertEqual(report2["added"], 1)

    def test_refresh_adopts_child_device_relationship(self):
        # Children synced before self-pointing relationships existed gain
        # one on the next sync, so Select Circuit lights up.
        result, set_id, _report = self._run_once()
        entries = [
            _entry("host:c1", parent="host:p1"),
            _entry("host:p1", role="parent"),
        ]
        relationship = {
            "device_unique_id": "c1", "device_id": 11, "device_name": "Recep",
        }
        context = {
            "status": "circuited", "status_text": "MDP - 12",
            "circuits": [{"circuit_id": 5, "circuit_number": "12",
                          "panel_name": "MDP", "circuit_name": "Recep"}],
        }
        child_snapshot = _snapshot("host:c1", x=1.0)
        child_snapshot["relationship"] = relationship
        child_snapshot["relationship_context"] = context
        snapshots = {
            "host:c1": child_snapshot,
            "host:p1": _snapshot("host:p1", x=5.0),
        }
        result2, set_id2, _report2 = sync_logic.apply_sync_membership(
            result, entries, snapshots, SOURCE
        )
        record = models.find_set(result2, set_id2)["elements"]["host:c1"]
        self.assertEqual(record["relationship"], relationship)
        self.assertEqual(
            record["relationship_context"]["circuits"][0]["circuit_number"], "12"
        )

    def test_linked_parent_wiring(self):
        # A child whose parent lives in a linked document carries a
        # "link:{link_uid}:{elem_uid}" parent pid; both are tracked in
        # the child's category set.
        entries = [
            _entry("host:c1", parent="link:L1:p9"),
            _entry("link:L1:p9", role="parent"),
        ]
        snapshots = {
            "host:c1": _snapshot("host:c1"),
            "link:L1:p9": _snapshot("link:L1:p9", x=9.0),
        }
        result, set_id, report = sync_logic.apply_sync_membership(
            models.new_project_store(), entries, snapshots, SOURCE
        )
        tracking_set = models.find_set(result, set_id)
        self.assertEqual(report["added"], 2)
        child = tracking_set["elements"]["host:c1"]
        parent = tracking_set["elements"]["link:L1:p9"]
        self.assertEqual(child["parent_persistent_id"], "link:L1:p9")
        self.assertEqual(parent["linker_meta"]["role"], "parent")

    def test_group_children_prefers_parent_persistent_id(self):
        groups = sync_logic.group_children([
            {"unique_id": "a", "parent_persistent_id": "link:L1:p9",
             "parent_unique_id": "p9"},
            {"unique_id": "b", "parent_persistent_id": "host:p9",
             "parent_unique_id": "p9"},
            {"unique_id": "c", "parent_persistent_id": None,
             "parent_unique_id": ""},
        ])
        self.assertEqual(sorted(groups.keys()), ["host:p9", "link:L1:p9"])

    def test_missing_snapshot_warns_and_skips(self):
        entries = [_entry("host:c1", parent="host:p1")]
        result, set_id, report = sync_logic.apply_sync_membership(
            models.new_project_store(), entries, {}, SOURCE
        )
        self.assertIsNone(set_id)
        self.assertEqual(result["tracking_sets"], [])
        self.assertEqual(len(report["warnings"]), 1)

    def test_sets_split_by_parent_category(self):
        # The set category comes from the PARENT equipment, not the
        # children: two parents of different categories make two sets,
        # each holding that parent and ALL its children regardless of
        # the children's own categories.
        entries = [
            _entry("host:c1", parent="host:p1"),
            _entry("host:c2", parent="host:p1"),
            _entry("host:c3", parent="host:p2"),
            _entry("host:p1", role="parent"),
            _entry("host:p2", role="parent"),
        ]
        snapshots = {
            "host:c1": _snapshot("host:c1", category="Electrical Fixtures"),
            "host:c2": _snapshot("host:c2", category="Lighting Fixtures"),
            "host:c3": _snapshot("host:c3", category="Electrical Fixtures"),
            "host:p1": _snapshot("host:p1", category="Electrical Equipment"),
            "host:p2": _snapshot("host:p2", category="Mechanical Equipment"),
        }
        result, set_id, report = sync_logic.apply_sync_membership(
            models.new_project_store(), entries, snapshots, SOURCE
        )
        self.assertEqual(
            sorted(report["sets_created"]),
            ["Element Linker - Electrical Equipment",
             "Element Linker - Mechanical Equipment"],
        )
        self.assertEqual(len(result["tracking_sets"]), 2)
        by_name = dict([
            (tracking_set["name"], tracking_set)
            for tracking_set in result["tracking_sets"]
        ])
        electrical = by_name["Element Linker - Electrical Equipment"]
        mechanical = by_name["Element Linker - Mechanical Equipment"]
        self.assertEqual(
            sorted(electrical["elements"].keys()),
            ["host:c1", "host:c2", "host:p1"],
        )
        self.assertEqual(
            sorted(mechanical["elements"].keys()), ["host:c3", "host:p2"]
        )
        # Primary set is the first category alphabetically.
        self.assertEqual(set_id, electrical["set_id"])

    def test_legacy_mixed_set_is_migrated_and_removed(self):
        legacy = models.new_tracking_set(
            "Element Linker Sync",
            SOURCE,
            {"id": None, "name": "Mixed (Element Linker)"},
            [],
            membership=models.MEMBERSHIP_EXPLICIT,
            origin=models.SET_ORIGIN_ELEMENT_LINKER,
        )
        old_record = models.new_tracked_element(
            _snapshot("host:c1"), baseline=True, track_location=True
        )
        old_record["linker_meta"] = {"role": "child"}
        legacy["elements"]["host:c1"] = old_record
        store = models.new_project_store()
        store["tracking_sets"].append(legacy)

        result, set_id, report = self._run_once(store=store)
        names = [
            tracking_set["name"] for tracking_set in result["tracking_sets"]
        ]
        self.assertEqual(names, ["Element Linker - Electrical Fixtures"])
        self.assertEqual(report["sets_removed"], ["Element Linker Sync"])
        self.assertEqual(report["stale_removed"], 1)
        tracking_set = models.find_set(result, set_id)
        self.assertIn("host:c1", tracking_set["elements"])


if __name__ == "__main__":
    unittest.main()
