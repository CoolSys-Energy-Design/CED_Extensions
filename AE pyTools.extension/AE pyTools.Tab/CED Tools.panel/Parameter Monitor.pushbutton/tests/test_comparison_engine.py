# -*- coding: utf-8 -*-
from __future__ import print_function

import copy
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = os.path.dirname(HERE)
if BUNDLE_DIR not in sys.path:
    sys.path.insert(0, BUNDLE_DIR)

import comparison_engine
import models


PROPERTY = {
    "key": "builtin:-1001203:instance",
    "name": "Mark",
    "scope": "instance",
    "storage_type": "string",
}
SECOND_PROPERTY = {
    "key": "name:comments:instance",
    "name": "Comments",
    "scope": "instance",
    "storage_type": "string",
}


def value(raw):
    return models.normalized_value(models.VALUE_VALID, "string", raw, raw)


def location(x=1.0, y=2.0, z=3.0, rotation=0.0):
    return {
        "state": models.VALUE_VALID,
        "x": x,
        "y": y,
        "z": z,
        "rotation": rotation,
    }


def snapshot(key="host:uid-1", mark="A", loc=None, family="Family", type_name="Type"):
    return {
        "persistent_id": key,
        "source_element_unique_id": key.split(":")[-1],
        "metadata": {
            "friendly_name": mark,
            "element_id": 1,
            "level": "Level 1",
            "family_type": "{} : {}".format(family, type_name),
            "family_name": family,
            "type_name": type_name,
        },
        "properties": {PROPERTY["key"]: value(mark)},
        "location": copy.deepcopy(loc),
        "relationship": None,
        "relationship_context": None,
    }


def tracking_set_with_baseline(track_location=False):
    tracking_set = models.new_tracking_set(
        "Test Set",
        {"source_type": models.SOURCE_HOST},
        {"id": -2001140, "name": "Mechanical Equipment"},
        [PROPERTY],
    )
    baseline = snapshot(loc=location())
    tracking_set["elements"][baseline["persistent_id"]] = models.new_tracked_element(
        baseline, baseline=True, track_location=track_location
    )
    tracking_set["status"] = models.SET_CLEAN
    return tracking_set


class ComparisonEngineTests(unittest.TestCase):
    def test_family_and_type_changes_are_independent_and_resolvable(self):
        tracking_set = tracking_set_with_baseline()
        type_changed = comparison_engine.apply_scan(
            tracking_set,
            {"host:uid-1": snapshot(type_name="Type B")},
            "t1",
        )
        record = type_changed["elements"]["host:uid-1"]
        self.assertIn(models.TYPE_PROPERTY_KEY, record["changed_property_keys"])
        self.assertNotIn(models.FAMILY_PROPERTY_KEY, record["changed_property_keys"])
        self.assertEqual("Type", record["accepted_metadata"]["type_name"])
        self.assertEqual("Type B", record["current_metadata"]["type_name"])

        type_resolved = comparison_engine.resolve_property(
            type_changed, "host:uid-1", models.TYPE_PROPERTY_KEY
        )
        self.assertEqual(models.SET_CLEAN, type_resolved["status"])

        family_changed = comparison_engine.apply_scan(
            type_resolved,
            {"host:uid-1": snapshot(family="Family B", type_name="Type B")},
            "t2",
        )
        family_record = family_changed["elements"]["host:uid-1"]
        self.assertIn(models.FAMILY_PROPERTY_KEY, family_record["changed_property_keys"])
        self.assertNotIn(models.TYPE_PROPERTY_KEY, family_record["changed_property_keys"])

    def test_change_does_not_advance_accepted(self):
        tracking_set = tracking_set_with_baseline()
        scanned = comparison_engine.apply_scan(
            tracking_set,
            {"host:uid-1": snapshot(mark="B")},
            "2026-08-09T12:00:00Z",
        )
        record = scanned["elements"]["host:uid-1"]
        self.assertEqual("A", record["accepted_properties"][PROPERTY["key"]]["raw"])
        self.assertEqual("B", record["current_properties"][PROPERTY["key"]]["raw"])
        self.assertEqual(models.SET_DIRTY, scanned["status"])

        scanned_again = comparison_engine.apply_scan(
            scanned,
            {"host:uid-1": snapshot(mark="C")},
            "2026-08-09T12:01:00Z",
        )
        record = scanned_again["elements"]["host:uid-1"]
        self.assertEqual("A", record["accepted_properties"][PROPERTY["key"]]["raw"])
        self.assertEqual("C", record["current_properties"][PROPERTY["key"]]["raw"])

    def test_revert_to_accepted_becomes_clean(self):
        tracking_set = tracking_set_with_baseline()
        dirty = comparison_engine.apply_scan(
            tracking_set, {"host:uid-1": snapshot(mark="B")}, "t1"
        )
        clean = comparison_engine.apply_scan(
            dirty, {"host:uid-1": snapshot(mark="A")}, "t2"
        )
        self.assertEqual(models.SET_CLEAN, clean["status"])
        self.assertEqual(0, clean["elements"]["host:uid-1"]["change_count"])

    def test_added_element_requires_resolve(self):
        tracking_set = tracking_set_with_baseline()
        current = {
            "host:uid-1": snapshot(),
            "host:uid-2": snapshot(key="host:uid-2", mark="NEW"),
        }
        scanned = comparison_engine.apply_scan(tracking_set, current, "t1")
        self.assertEqual(models.ELEMENT_ADDED, scanned["elements"]["host:uid-2"]["state"])
        self.assertEqual(models.SET_DIRTY, scanned["status"])
        resolved = comparison_engine.resolve_element(scanned, "host:uid-2")
        self.assertEqual(models.ELEMENT_TRACKED, resolved["elements"]["host:uid-2"]["state"])
        self.assertEqual(models.SET_CLEAN, resolved["status"])

    def test_unaccepted_added_element_that_disappears_is_dropped(self):
        tracking_set = tracking_set_with_baseline()
        current = {
            "host:uid-1": snapshot(),
            "host:uid-2": snapshot(key="host:uid-2", mark="NEW"),
        }
        added = comparison_engine.apply_scan(tracking_set, current, "t1")
        disappeared = comparison_engine.apply_scan(
            added, {"host:uid-1": snapshot()}, "t2"
        )
        self.assertNotIn("host:uid-2", disappeared["elements"])

    def test_removed_record_is_retained_and_reappears(self):
        tracking_set = tracking_set_with_baseline()
        removed = comparison_engine.apply_scan(tracking_set, {}, "t1")
        record = removed["elements"]["host:uid-1"]
        self.assertEqual(models.ELEMENT_REMOVED, record["state"])
        self.assertEqual("A", record["accepted_properties"][PROPERTY["key"]]["raw"])
        reappeared = comparison_engine.apply_scan(
            removed, {"host:uid-1": snapshot(mark="B")}, "t2"
        )
        record = reappeared["elements"]["host:uid-1"]
        self.assertEqual(models.ELEMENT_TRACKED, record["state"])
        self.assertEqual(1, record["change_count"])

    def test_missing_is_not_blank(self):
        missing = models.missing_value()
        blank = models.normalized_value(models.VALUE_BLANK, "string", None, "")
        self.assertFalse(comparison_engine.normalized_values_equal(missing, blank))
        self.assertTrue(comparison_engine.normalized_values_equal(missing, models.missing_value()))

    def test_untrack_preserves_last_known_record_and_restore_baselines(self):
        tracking_set = tracking_set_with_baseline()
        untracked = comparison_engine.untrack_element(tracking_set, "host:uid-1")
        self.assertIn("host:uid-1", untracked["elements"])
        self.assertEqual(["host:uid-1"], untracked["untracked_ids"])
        preserved = untracked["elements"]["host:uid-1"]["metadata"]
        self.assertEqual(1, preserved["element_id"])
        self.assertEqual("Family", preserved["family_name"])
        self.assertEqual("Type", preserved["type_name"])
        self.assertEqual("Level 1", preserved["level"])
        rescanned = comparison_engine.apply_scan(untracked, {}, "t1")
        self.assertIn("host:uid-1", rescanned["elements"])
        self.assertEqual(0, comparison_engine.summarize_set(rescanned)["tracked"])
        restored = comparison_engine.restore_element(
            rescanned, "host:uid-1", snapshot(mark="RESTORED")
        )
        self.assertNotIn("host:uid-1", restored["untracked_ids"])
        record = restored["elements"]["host:uid-1"]
        self.assertEqual("RESTORED", record["accepted_properties"][PROPERTY["key"]]["raw"])

    def test_bulk_untrack_updates_all_selected_records_in_one_result(self):
        tracking_set = tracking_set_with_baseline()
        second = snapshot(key="host:uid-2", mark="B")
        tracking_set["elements"]["host:uid-2"] = models.new_tracked_element(
            second,
            baseline=True,
        )

        updated = comparison_engine.untrack_elements(
            tracking_set,
            ["host:uid-1", "host:uid-2"],
        )

        self.assertEqual(
            set(["host:uid-1", "host:uid-2"]),
            set(updated["elements"].keys()),
        )
        self.assertEqual(["host:uid-1", "host:uid-2"], updated["untracked_ids"])
        self.assertEqual(2, len(tracking_set["elements"]))

    def test_location_tolerance_and_change(self):
        tracking_set = tracking_set_with_baseline(track_location=True)
        within = comparison_engine.apply_scan(
            tracking_set,
            {"host:uid-1": snapshot(loc=location(x=1.0005))},
            "t1",
        )
        self.assertEqual(0, within["elements"]["host:uid-1"]["change_count"])
        moved = comparison_engine.apply_scan(
            tracking_set,
            {"host:uid-1": snapshot(loc=location(x=1.01))},
            "t2",
        )
        self.assertIn(
            models.LOCATION_PROPERTY_KEY,
            moved["elements"]["host:uid-1"]["changed_property_keys"],
        )

    def test_link_transform_is_one_source_condition(self):
        tracking_set = tracking_set_with_baseline(track_location=True)
        tracking_set["accepted_source_state"] = {"matrix": [1.0, 0.0, 0.0]}
        scanned = comparison_engine.apply_scan(
            tracking_set,
            {"host:uid-1": snapshot(loc=location())},
            "t1",
            source_state={"matrix": [1.0, 0.0, 2.0]},
        )
        self.assertEqual(1, len(scanned["source_conditions"]))
        self.assertEqual(0, scanned["elements"]["host:uid-1"]["change_count"])

    def test_link_transform_is_ignored_without_location_tracking(self):
        tracking_set = tracking_set_with_baseline(track_location=False)
        tracking_set["accepted_source_state"] = {"matrix": [1.0, 0.0, 0.0]}
        scanned = comparison_engine.apply_scan(
            tracking_set,
            {"host:uid-1": snapshot()},
            "t1",
            source_state={"matrix": [1.0, 0.0, 2.0]},
        )
        self.assertEqual([], scanned["source_conditions"])
        self.assertEqual(
            {"matrix": [1.0, 0.0, 2.0]},
            scanned["accepted_source_state"],
        )

    def test_resolve_set_keeps_removed_records(self):
        tracking_set = tracking_set_with_baseline()
        removed = comparison_engine.apply_scan(tracking_set, {}, "t1")
        resolved = comparison_engine.resolve_set(removed)
        self.assertIn("host:uid-1", resolved["elements"])
        self.assertEqual(models.ELEMENT_REMOVED, resolved["elements"]["host:uid-1"]["state"])
        self.assertEqual(models.SET_DIRTY, resolved["status"])

    def test_adding_property_accepts_current_immediately(self):
        tracking_set = tracking_set_with_baseline()
        current_snapshot = snapshot()
        current_snapshot["properties"][SECOND_PROPERTY["key"]] = value("new baseline")
        updated = comparison_engine.update_tracked_properties(
            tracking_set,
            [PROPERTY, SECOND_PROPERTY],
            current_map={"host:uid-1": current_snapshot},
        )
        record = updated["elements"]["host:uid-1"]
        self.assertEqual(
            "new baseline",
            record["accepted_properties"][SECOND_PROPERTY["key"]]["raw"],
        )
        self.assertEqual(0, record["change_count"])


if __name__ == "__main__":
    unittest.main()
