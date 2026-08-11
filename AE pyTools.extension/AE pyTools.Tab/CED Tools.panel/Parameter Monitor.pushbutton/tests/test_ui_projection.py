# -*- coding: utf-8 -*-
from __future__ import print_function

import sys
import types
import unittest

import models

tracking_service_stub = types.ModuleType("tracking_service")
tracking_service_stub.set_summary = lambda tracking_set: {}
previous_tracking_service = sys.modules.get("tracking_service")
sys.modules["tracking_service"] = tracking_service_stub
import viewmodel
if previous_tracking_service is None:
    del sys.modules["tracking_service"]
else:
    sys.modules["tracking_service"] = previous_tracking_service


def _record(name, accepted, current, changed=False):
    key = "name:sample:instance"
    return {
        "metadata": {
            "friendly_name": name,
            "family_type": "Family : Type",
            "element_id": name[-1],
            "level": "Level 1",
        },
        "state": models.ELEMENT_TRACKED,
        "accepted_properties": {
            key: {"state": models.VALUE_VALID, "display": accepted},
        },
        "current_properties": {
            key: {"state": models.VALUE_VALID, "display": current},
        },
        "changed_property_keys": [key] if changed else [],
        "change_count": 1 if changed else 0,
        "missing_count": 0,
        "track_location": False,
    }


class UiProjectionTests(unittest.TestCase):
    def test_property_grid_values_follow_selected_element_record(self):
        key = "name:sample:instance"
        tracking_set = {
            "tracked_properties": [
                {"key": key, "name": "Sample", "scope": "instance"},
            ],
        }
        first = viewmodel.ElementRow("first", _record("Element 1", "A1", "C1"))
        second = viewmodel.ElementRow(
            "second",
            _record("Element 2", "A2", "C2", changed=True),
        )

        first_rows = viewmodel.property_rows(tracking_set, first)
        second_rows = viewmodel.property_rows(tracking_set, second)

        first_property = [row for row in first_rows if row.key == key][0]
        second_property = [row for row in second_rows if row.key == key][0]
        self.assertEqual((first_property.accepted, first_property.current), ("A1", "C1"))
        self.assertEqual((second_property.accepted, second_property.current), ("A2", "C2"))
        self.assertFalse(first_property.changed)
        self.assertTrue(second_property.changed)

    def test_linker_children_hidden_from_grid_and_listed_under_parent(self):
        location = {
            "state": models.VALUE_VALID,
            "x": 0.0, "y": 0.0, "z": 0.0, "rotation": 0.0,
        }
        parent = _record("Parent 1", "A", "A")
        parent["persistent_id"] = "host:p1"
        parent["track_location"] = True
        parent["accepted_location"] = dict(location)
        parent["current_location"] = dict(location, x=5.0)
        child_profile = _record("Child 1", "A", "A")
        child_profile["persistent_id"] = "host:c1"
        child_profile["parent_persistent_id"] = "host:p1"
        child_profile["linker_meta"] = {"role": "child", "led_id": "LED-1"}
        child_profile["metadata"]["family_type"] = "Recep Family : Quad"
        child_manual = _record("Child 2", "A", "A")
        child_manual["persistent_id"] = "host:c2"
        child_manual["parent_persistent_id"] = "host:p1"
        child_removed = _record("Child 3", "A", "A")
        child_removed["persistent_id"] = "host:c3"
        child_removed["parent_persistent_id"] = "host:p1"
        child_removed["state"] = models.ELEMENT_REMOVED
        tracking_set = {
            "elements": {
                "host:p1": parent,
                "host:c1": child_profile,
                "host:c2": child_manual,
                "host:c3": child_removed,
            },
            "location_defaults": {
                "translation_tolerance": 0.001,
                "angular_tolerance": 0.0017453292519943296,
            },
        }

        grid_rows = viewmodel.element_rows(tracking_set)
        self.assertEqual([row.persistent_id for row in grid_rows], ["host:p1"])

        info = viewmodel.linked_children_info(tracking_set, parent)
        self.assertEqual(info["count"], 3)
        by_id = dict([(row.persistent_id, row) for row in info["children"]])
        self.assertEqual(by_id["host:c1"].origin, "Profile")
        self.assertEqual(by_id["host:c1"].family, "Recep Family")
        self.assertEqual(by_id["host:c1"].type, "Quad")
        self.assertEqual(by_id["host:c2"].origin, "Manual")
        self.assertTrue(info["parent_moved"])
        self.assertEqual(
            sorted(info["movable_child_ids"]), ["host:c1", "host:c2"]
        )

    def test_linked_children_info_in_sync_parent(self):
        parent = _record("Parent 1", "A", "A")
        parent["persistent_id"] = "host:p1"
        child = _record("Child 1", "A", "A")
        child["persistent_id"] = "host:c1"
        child["parent_persistent_id"] = "host:p1"
        tracking_set = {"elements": {"host:p1": parent, "host:c1": child}}
        info = viewmodel.linked_children_info(tracking_set, parent)
        self.assertEqual(info["count"], 1)
        self.assertFalse(info["parent_moved"])
        self.assertEqual(info["movable_child_ids"], [])
        empty = viewmodel.linked_children_info(tracking_set, child)
        self.assertEqual(empty["count"], 0)

    def test_element_rows_keep_their_own_record_objects(self):
        tracking_set = {
            "elements": {
                "first": _record("Element 1", "A1", "C1"),
                "second": _record("Element 2", "A2", "C2", changed=True),
            },
        }

        rows = viewmodel.element_rows(tracking_set)
        by_id = dict((row.persistent_id, row) for row in rows)

        self.assertIsNot(by_id["first"].record, by_id["second"].record)
        self.assertEqual(by_id["first"].record["current_properties"]["name:sample:instance"]["display"], "C1")
        self.assertEqual(by_id["second"].record["current_properties"]["name:sample:instance"]["display"], "C2")

    def test_projects_two_thousand_element_rows(self):
        tracking_set = {"elements": {}}
        for index in range(2000):
            persistent_id = "host:uid-{}".format(index)
            tracking_set["elements"][persistent_id] = _record(
                "Element {}".format(index),
                "A{}".format(index),
                "C{}".format(index),
                changed=(index % 10 == 0),
            )

        rows = viewmodel.element_rows(tracking_set)

        self.assertEqual(2000, len(rows))
        self.assertEqual(200, len([row for row in rows if row.status == "Changed"]))


if __name__ == "__main__":
    unittest.main()
