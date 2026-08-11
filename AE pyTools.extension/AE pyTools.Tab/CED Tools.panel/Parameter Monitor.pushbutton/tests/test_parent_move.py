# -*- coding: utf-8 -*-
from __future__ import print_function

import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
BUNDLE_DIR = os.path.dirname(HERE)
if BUNDLE_DIR not in sys.path:
    sys.path.insert(0, BUNDLE_DIR)

import models
import parent_move_math


def _location(x=0.0, y=0.0, z=0.0, rotation=0.0, state=models.VALUE_VALID):
    return {
        "state": state, "x": x, "y": y, "z": z, "rotation": rotation,
        "coordinate_system": "source_document_internal",
    }


class ComputeFollowDeltaTests(unittest.TestCase):
    def test_translation_only(self):
        delta = parent_move_math.compute_follow_delta(
            _location(1.0, 2.0, 3.0), _location(4.0, 6.0, 3.0)
        )
        self.assertEqual(delta["translation"], (3.0, 4.0, 0.0))
        self.assertEqual(delta["rotation_delta"], 0.0)
        self.assertEqual(delta["pivot"], (4.0, 6.0, 3.0))

    def test_rotation_wraps_signed(self):
        # 350deg -> 10deg should be +20deg, not -340deg.
        delta = parent_move_math.compute_follow_delta(
            _location(rotation=math.radians(350.0)),
            _location(rotation=math.radians(10.0)),
        )
        self.assertAlmostEqual(delta["rotation_delta"], math.radians(20.0), places=9)
        delta = parent_move_math.compute_follow_delta(
            _location(rotation=math.radians(10.0)),
            _location(rotation=math.radians(350.0)),
        )
        self.assertAlmostEqual(delta["rotation_delta"], math.radians(-20.0), places=9)

    def test_invalid_locations_return_none(self):
        self.assertIsNone(parent_move_math.compute_follow_delta(None, _location()))
        self.assertIsNone(parent_move_math.compute_follow_delta(
            _location(state=models.VALUE_UNSUPPORTED), _location()
        ))


class TransformPointTests(unittest.TestCase):
    def test_translation_only(self):
        delta = {"translation": (1.0, 2.0, 3.0), "rotation_delta": 0.0,
                 "pivot": (0.0, 0.0, 0.0)}
        self.assertEqual(
            parent_move_math.transform_point((1.0, 1.0, 1.0), delta),
            (2.0, 3.0, 4.0),
        )

    def test_rotation_about_pivot(self):
        # Parent at origin rotates +90deg; child 1 ft east ends up 1 ft north.
        delta = {"translation": (0.0, 0.0, 0.0),
                 "rotation_delta": math.pi / 2.0,
                 "pivot": (0.0, 0.0, 0.0)}
        x, y, z = parent_move_math.transform_point((1.0, 0.0, 0.0), delta)
        self.assertAlmostEqual(x, 0.0, places=9)
        self.assertAlmostEqual(y, 1.0, places=9)
        self.assertAlmostEqual(z, 0.0, places=9)

    def test_combined_translate_then_rotate(self):
        # Parent moved from (0,0) to (10,0) and rotated +90deg. Child was at
        # (1,0): translate -> (11,0), rotate about pivot (10,0) -> (10,1).
        delta = {"translation": (10.0, 0.0, 0.0),
                 "rotation_delta": math.pi / 2.0,
                 "pivot": (10.0, 0.0, 0.0)}
        x, y, _z = parent_move_math.transform_point((1.0, 0.0, 0.0), delta)
        self.assertAlmostEqual(x, 10.0, places=9)
        self.assertAlmostEqual(y, 1.0, places=9)


class SignificantDeltaTests(unittest.TestCase):
    def test_below_tolerances_is_insignificant(self):
        delta = {"translation": (0.0005, 0.0, 0.0), "rotation_delta": 0.0005,
                 "pivot": (0.0, 0.0, 0.0)}
        self.assertFalse(
            parent_move_math.significant_delta(delta, 0.001, 0.0017453292519943296)
        )

    def test_translation_or_rotation_triggers(self):
        base = {"translation": (0.0, 0.0, 0.0), "rotation_delta": 0.0,
                "pivot": (0.0, 0.0, 0.0)}
        moved = dict(base, translation=(0.5, 0.0, 0.0))
        rotated = dict(base, rotation_delta=0.1)
        self.assertTrue(parent_move_math.significant_delta(moved, 0.001, 0.001))
        self.assertTrue(parent_move_math.significant_delta(rotated, 0.001, 0.001))
        self.assertFalse(parent_move_math.significant_delta(None, 0.001, 0.001))


class LinkerPoseUpdatesTests(unittest.TestCase):
    def test_degree_conversion_and_fields(self):
        child = _location(1.0, 2.0, 3.0, rotation=math.pi / 2.0)
        parent = _location(10.0, 20.0, 0.0, rotation=math.pi)
        updates = parent_move_math.linker_pose_updates(child, parent)
        self.assertEqual(updates["location_ft"], [1.0, 2.0, 3.0])
        self.assertAlmostEqual(updates["rotation_deg"], 90.0, places=9)
        self.assertEqual(updates["parent_location_ft"], [10.0, 20.0, 0.0])
        self.assertAlmostEqual(updates["parent_rotation_deg"], 180.0, places=9)

    def test_invalid_child_location_omits_child_fields(self):
        updates = parent_move_math.linker_pose_updates(
            _location(state=models.VALUE_UNSUPPORTED), _location(1.0)
        )
        self.assertNotIn("location_ft", updates)
        self.assertIn("parent_location_ft", updates)


class AcceptParentTests(unittest.TestCase):
    def _set_with_parent_and_children(self, dirty_child=False):
        tracking_set = models.new_tracking_set(
            "T", {}, {"id": None, "name": "Mixed"}, [],
            membership=models.MEMBERSHIP_EXPLICIT,
        )
        parent = models.new_tracked_element({
            "persistent_id": "host:p", "metadata": {}, "properties": {},
            "location": _location(0.0),
        }, baseline=True, track_location=True)
        parent["current_location"] = _location(5.0)
        parent["changed_property_keys"] = [models.LOCATION_PROPERTY_KEY]
        parent["change_count"] = 1
        child = models.new_tracked_element({
            "persistent_id": "host:c", "metadata": {}, "properties": {},
            "location": _location(1.0),
        }, baseline=True, track_location=True)
        child["parent_persistent_id"] = "host:p"
        if dirty_child:
            child["changed_property_keys"] = [models.LOCATION_PROPERTY_KEY]
            child["change_count"] = 1
        tracking_set["elements"] = {"host:p": parent, "host:c": child}
        return tracking_set

    def test_accepts_when_all_children_clean(self):
        tracking_set = self._set_with_parent_and_children(dirty_child=False)
        accepted = parent_move_math.accept_parent_location_if_children_clean(
            tracking_set, "host:p"
        )
        self.assertTrue(accepted)
        parent = tracking_set["elements"]["host:p"]
        self.assertEqual(parent["accepted_location"]["x"], 5.0)
        self.assertNotIn(
            models.LOCATION_PROPERTY_KEY,
            list(parent.get("changed_property_keys") or []),
        )

    def test_does_not_accept_with_dirty_child(self):
        tracking_set = self._set_with_parent_and_children(dirty_child=True)
        accepted = parent_move_math.accept_parent_location_if_children_clean(
            tracking_set, "host:p"
        )
        self.assertFalse(accepted)
        parent = tracking_set["elements"]["host:p"]
        self.assertEqual(parent["accepted_location"]["x"], 0.0)

    def test_missing_parent_returns_false(self):
        tracking_set = self._set_with_parent_and_children()
        self.assertFalse(
            parent_move_math.accept_parent_location_if_children_clean(
                tracking_set, "host:nope"
            )
        )


if __name__ == "__main__":
    unittest.main()
