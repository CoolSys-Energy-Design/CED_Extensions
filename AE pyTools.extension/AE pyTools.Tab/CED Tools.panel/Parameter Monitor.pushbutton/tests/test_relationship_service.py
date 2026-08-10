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
snippets_stub.revit_helpers = revit_helpers_stub
previous_snippets = sys.modules.get("Snippets")
sys.modules["Snippets"] = snippets_stub
import relationship_service
if previous_snippets is None:
    del sys.modules["Snippets"]
else:
    sys.modules["Snippets"] = previous_snippets


class _ObjectType(object):
    Element = object()


class _ExceptionType(object):
    FullName = "Autodesk.Revit.Exceptions.OperationCanceledException"


class _CancelledPick(Exception):
    def GetType(self):
        return _ExceptionType()


class _Selection(object):
    def PickObject(self, *args):
        raise _CancelledPick("The user aborted the pick operation.")


class _UiDocument(object):
    Selection = _Selection()


class RelationshipServiceTests(unittest.TestCase):
    def test_escape_during_pick_returns_cancelled_result(self):
        previous = relationship_service.ObjectType
        relationship_service.ObjectType = _ObjectType
        try:
            self.assertIsNone(relationship_service.pick_device(_UiDocument()))
        finally:
            relationship_service.ObjectType = previous


if __name__ == "__main__":
    unittest.main()
