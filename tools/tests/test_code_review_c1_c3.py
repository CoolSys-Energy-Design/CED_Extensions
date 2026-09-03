# -*- coding: utf-8 -*-
"""Static regression contracts for code-review findings C1 through C3."""

import ast
import importlib.util
from pathlib import Path
import sys
import types


ROOT = Path(__file__).resolve().parents[2]
OPERATIONS = ROOT / "CEDLib.lib" / "CEDElectrical" / "Application" / "operations"
SETTINGS = ROOT / "CEDLib.lib" / "CEDElectrical" / "Domain" / "settings_manager.py"
ALERTS_BUNDLE = (
    ROOT
    / "CED ElecTools.extension"
    / "AE pyTools.tab"
    / "Electrical.panel"
    / "Alerts Manager.pushbutton"
)


def _read(path):
    return path.read_text(encoding="utf-8-sig")


def _function_source(path, function_name):
    source = _read(path)
    tree = ast.parse(source)
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == function_name:
            return ast.get_source_segment(source, node)
    raise AssertionError("Missing function: {}".format(function_name))


class _RepresentativeElementId(object):
    def __init__(self, value):
        self.Value = int(value)
        self.IntegerValue = int(value)


def _load_revit_helpers():
    db = types.ModuleType("Autodesk.Revit.DB")
    db.ElementId = _RepresentativeElementId
    db.Element = type("Element", (object,), {})
    db.FamilySymbol = type("FamilySymbol", (object,), {})
    db.BuiltInParameter = types.SimpleNamespace(
        SYMBOL_NAME_PARAM="symbol_name",
        ELEM_TYPE_PARAM="element_type",
    )
    autodesk = types.ModuleType("Autodesk")
    revit_package = types.ModuleType("Autodesk.Revit")
    autodesk.Revit = revit_package
    revit_package.DB = db
    system = types.ModuleType("System")
    system.Int64 = int

    saved = {}
    for name, module in {
        "Autodesk": autodesk,
        "Autodesk.Revit": revit_package,
        "Autodesk.Revit.DB": db,
        "System": system,
    }.items():
        saved[name] = sys.modules.get(name)
        sys.modules[name] = module
    try:
        path = ROOT / "CEDLib.lib" / "Snippets" / "revit_helpers.py"
        spec = importlib.util.spec_from_file_location("c1_c3_revit_helpers", str(path))
        helpers = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(helpers)
        return helpers
    finally:
        for name, previous in saved.items():
            if previous is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = previous


def test_c1_compound_operations_only_assimilate_success():
    names = (
        "edit_circuit_properties_and_recalculate_operation.py",
        "mark_existing_and_recalculate_operation.py",
        "set_include_and_recalculate_operation.py",
    )
    for name in names:
        source = _read(OPERATIONS / name)
        assert "calc_status = str(calc_result.get(" in source
        assert "if calc_status != " in source
        assert source.count("tg.Assimilate()") == 1


def test_c2_settings_transactions_roll_back_before_reraising():
    source = _read(SETTINGS)
    assert "transaction.HasStarted()" in source
    assert "transaction.RollBack()" in source
    for function_name in ("_create_global_param", "save_circuit_settings"):
        function_source = _function_source(SETTINGS, function_name)
        assert "except Exception:" in function_source
        assert "_rollback_started_transaction(t)" in function_source
        assert "raise" in function_source


def test_c3_alerts_payloads_are_bound_to_document_and_type_checked():
    script = _read(ALERTS_BUNDLE / "script.py")
    services = _read(ALERTS_BUNDLE / "alerts_browser_services.py")
    assert 'data["doc_key"] = self._document_key' in script
    assert 'active_key != expected_key' in script
    assert 'self._gateway._document_key != expected_key' in script
    assert "revit_helpers.coerce_elementid_value" in script
    assert "isinstance(circuit, DBE.ElectricalSystem)" in script
    assert 'status = "invalid_context"' in script
    assert '"doc_key": get_document_key(doc)' in services
    assert "doc.GetHashCode()" in services


def test_c3_boundary_id_coercion_accepts_numeric_and_native_representative_ids():
    helpers = _load_revit_helpers()
    assert helpers.coerce_elementid_value(314) == 314
    assert helpers.coerce_elementid_value("314") == 314
    assert helpers.coerce_elementid_value(_RepresentativeElementId(314)) == 314


if __name__ == "__main__":
    tests = sorted(
        (name, value)
        for name, value in globals().items()
        if name.startswith("test_") and callable(value)
    )
    for name, test in tests:
        test()
        print("PASS {}".format(name))
