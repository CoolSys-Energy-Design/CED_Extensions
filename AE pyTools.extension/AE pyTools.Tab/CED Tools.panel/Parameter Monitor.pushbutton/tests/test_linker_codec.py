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

import element_linker_codec as codec


class _FakeParam(object):
    def __init__(self, text, read_only=False):
        self._text = text
        self.IsReadOnly = read_only

    def AsString(self):
        return self._text

    def Set(self, text):
        self._text = text


class _FakeElement(object):
    def __init__(self, param):
        self._param = param

    def LookupParameter(self, name):
        return self._param if name == codec.PARAMETER_NAME else None


class ParsePayloadTests(unittest.TestCase):
    def test_json_round_trip_matches_meprfp_shape(self):
        linker = codec.empty_linker()
        linker["led_id"] = "SET-1-LED-2"
        linker["parent_element_id"] = 12345
        linker["parent_location_ft"] = [1.0, 2.0, 3.0]
        linker["host_name"] = "RTU Family"
        text = codec.serialize_payload(linker)
        data = json.loads(text)
        self.assertEqual(data["v"], codec.CODEC_VERSION)
        for name in codec.FIELDS:
            self.assertIn(name, data)
        parsed = codec.parse_payload(text)
        self.assertEqual(parsed, linker)

    def test_unsupported_version_and_blank_return_none(self):
        self.assertIsNone(codec.parse_payload('{"v": 99, "led_id": "X"}'))
        self.assertIsNone(codec.parse_payload(""))
        self.assertIsNone(codec.parse_payload(None))
        self.assertIsNone(codec.parse_payload("not json, not legacy"))

    def test_legacy_multiline_text(self):
        text = (
            "Linked Element Definition ID: SET-1-LED-9\n"
            "Parent ElementId: 4242\n"
            "Location XYZ (ft): 1.5, 2.5, 0.0\n"
            "Rotation (deg): 90\n"
            "Host Name: RTU Family\n"
        )
        parsed = codec.parse_payload(text)
        self.assertEqual(parsed["led_id"], "SET-1-LED-9")
        self.assertEqual(parsed["parent_element_id"], 4242)
        self.assertEqual(parsed["location_ft"], [1.5, 2.5, 0.0])
        self.assertEqual(parsed["rotation_deg"], 90.0)
        self.assertEqual(parsed["host_name"], "RTU Family")

    def test_legacy_inline_text(self):
        text = (
            "Linked Element Definition ID: SET-1-LED-9, "
            "Parent ElementId: 4242, Host Name: RTU Family"
        )
        parsed = codec.parse_payload(text)
        self.assertEqual(parsed["led_id"], "SET-1-LED-9")
        self.assertEqual(parsed["parent_element_id"], 4242)
        self.assertEqual(parsed["host_name"], "RTU Family")

    def test_legacy_not_found_becomes_none(self):
        parsed = codec.parse_payload("Parent ElementId: Not found\nHost Name: X\n")
        self.assertIsNone(parsed["parent_element_id"])


class ElementIoTests(unittest.TestCase):
    def test_read_and_update_round_trip(self):
        original = codec.empty_linker()
        original["led_id"] = "LED-1"
        original["parent_element_id"] = 7
        original["parent_location_ft"] = [0.0, 0.0, 0.0]
        param = _FakeParam(codec.serialize_payload(original))
        element = _FakeElement(param)

        self.assertEqual(codec.read_linker(element)["led_id"], "LED-1")
        codec.update_linker(element, {
            "location_ft": [9.0, 9.0, 0.0],
            "parent_location_ft": [5.0, 5.0, 0.0],
            "parent_rotation_deg": 90.0,
        })
        reread = codec.read_linker(element)
        self.assertEqual(reread["location_ft"], [9.0, 9.0, 0.0])
        self.assertEqual(reread["parent_location_ft"], [5.0, 5.0, 0.0])
        self.assertEqual(reread["parent_rotation_deg"], 90.0)
        self.assertEqual(reread["led_id"], "LED-1")
        self.assertEqual(reread["parent_element_id"], 7)

    def test_update_unknown_field_raises(self):
        element = _FakeElement(_FakeParam(codec.serialize_payload(codec.empty_linker())))
        self.assertRaises(
            codec.LinkerCodecError,
            codec.update_linker, element, {"nope": 1},
        )

    def test_update_missing_or_readonly_param_raises(self):
        self.assertRaises(
            codec.LinkerCodecError,
            codec.update_linker, _FakeElement(None), {"led_id": "X"},
        )
        element = _FakeElement(_FakeParam("{}", read_only=True))
        self.assertRaises(
            codec.LinkerCodecError,
            codec.update_linker, element, {"led_id": "X"},
        )

    def test_read_missing_param_returns_none(self):
        self.assertIsNone(codec.read_linker(_FakeElement(None)))
        self.assertIsNone(codec.read_linker(None))


if __name__ == "__main__":
    unittest.main()
