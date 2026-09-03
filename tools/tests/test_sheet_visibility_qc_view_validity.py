# -*- coding: utf-8 -*-
"""Representative API-signature check for Sheet Visibility QC view scans."""

import ast
from pathlib import Path
import types


ROOT = Path(__file__).resolve().parents[2]
SOURCE_PATH = (
    ROOT
    / "AE pyTools.extension"
    / "AE pyTools.Tab"
    / "Miscellaneous.panel"
    / "MiscTools1.stack"
    / "MiscTools.pulldown"
    / "SheetVisibilityQC.pushbutton"
    / "script.py"
)
SOURCE = SOURCE_PATH.read_text(encoding="utf-8-sig")
TREE = ast.parse(SOURCE)


def _load_collect_function(validity_callback):
    function_node = None
    for node in ast.walk(TREE):
        if isinstance(node, ast.FunctionDef) and node.name == "_collect_ids_from_view":
            function_node = node
            break
    if function_node is None:
        raise AssertionError("Missing _collect_ids_from_view")

    class _Collector(object):
        IsViewValidForElementIteration = staticmethod(validity_callback)

    module = ast.Module(body=[function_node], type_ignores=[])
    ast.fix_missing_locations(module)
    namespace = {
        "DB": types.SimpleNamespace(FilteredElementCollector=_Collector),
        "_view_name": lambda view: getattr(view, "Name", "View"),
    }
    exec(compile(module, str(SOURCE_PATH), "exec"), namespace)
    return namespace["_collect_ids_from_view"]


def test_view_validity_check_receives_document_and_native_view_id():
    calls = []

    def validity_check(document, view_id):
        calls.append((document, view_id))
        return False

    collect = _load_collect_function(validity_check)
    document = object()
    view_id = object()
    view = types.SimpleNamespace(Id=view_id, Name="Schedule")
    warnings = []

    assert collect(document, view, None, warnings=warnings) == set()
    assert calls == [(document, view_id)]
    assert len(warnings) == 1


if __name__ == "__main__":
    test_view_validity_check_receives_document_and_native_view_id()
    print("PASS test_view_validity_check_receives_document_and_native_view_id")
