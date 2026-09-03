# -*- coding: utf-8 -*-
"""Representative document-ownership checks for CircuitBranch feeder logic."""

import ast
from pathlib import Path
import types


ROOT = Path(__file__).resolve().parents[2]
SOURCE_PATH = ROOT / "CEDLib.lib" / "CEDElectrical" / "Model" / "CircuitBranch.py"
SOURCE = SOURCE_PATH.read_text(encoding="utf-8-sig")
TREE = ast.parse(SOURCE)


def _method_node(name):
    for node in ast.walk(TREE):
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError("Missing method: {}".format(name))


class _FamilyInstance(object):
    def __init__(self, parameter):
        self._parameter = parameter

    def get_Parameter(self, unused_parameter_id):
        return self._parameter


class _DistributionSystem(object):
    VoltageLineToGround = 120.0


class _Parameter(object):
    HasValue = True

    def __init__(self, element_id):
        self._element_id = element_id

    def AsElementId(self):
        return self._element_id


class _Document(object):
    def __init__(self, result):
        self.result = result
        self.requested_ids = []

    def GetElement(self, element_id):
        self.requested_ids.append(element_id)
        return self.result


def _load_method():
    node = _method_node("_has_feeder_ln_voltage")
    module = ast.Module(body=[node], type_ignores=[])
    ast.fix_missing_locations(module)
    namespace = {
        "DB": types.SimpleNamespace(
            FamilyInstance=_FamilyInstance,
            BuiltInParameter=types.SimpleNamespace(
                RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM="distribution_system",
            ),
        ),
        "DBE": types.SimpleNamespace(DistributionSysType=_DistributionSystem),
        "design_options": types.SimpleNamespace(
            is_main_model_element=lambda unused_element: True,
        ),
        "logger": types.SimpleNamespace(debug=lambda unused_message: None),
    }
    exec(compile(module, str(SOURCE_PATH), "exec"), namespace)
    return namespace["_has_feeder_ln_voltage"]


def test_feeder_lookup_uses_the_circuits_document():
    method = _load_method()
    element_id = object()
    owning_document = _Document(_DistributionSystem())
    circuit = types.SimpleNamespace(Document=owning_document)
    branch = types.SimpleNamespace(
        circuit=circuit,
        connected_elements=[_FamilyInstance(_Parameter(element_id))],
        name="Feeder",
    )

    assert method(branch) == 1
    assert owning_document.requested_ids == [element_id]


def test_feeder_lookup_fails_safe_without_an_owning_document():
    method = _load_method()
    branch = types.SimpleNamespace(
        circuit=types.SimpleNamespace(Document=None),
        connected_elements=[],
        name="Feeder",
    )
    assert method(branch) == 0


def test_feeder_lookup_has_no_active_document_dependency():
    method_source = ast.get_source_segment(SOURCE, _method_node("_has_feeder_ln_voltage"))
    assert "revit.doc" not in method_source
    assert 'getattr(self.circuit, "Document", None)' in method_source


if __name__ == "__main__":
    tests = sorted(
        (name, value)
        for name, value in globals().items()
        if name.startswith("test_") and callable(value)
    )
    for name, test in tests:
        test()
        print("PASS {}".format(name))
