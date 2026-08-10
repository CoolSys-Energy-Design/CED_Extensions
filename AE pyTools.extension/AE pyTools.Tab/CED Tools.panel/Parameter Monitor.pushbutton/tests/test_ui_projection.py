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
