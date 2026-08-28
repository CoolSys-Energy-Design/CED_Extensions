# -*- coding: utf-8 -*-
"""Static contract checks for Calculate Circuits operation callers."""

from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
OPERATIONS = ROOT / "CEDLib.lib" / "CEDElectrical" / "Application" / "operations"


def _read(path):
    return path.read_text(encoding="utf-8-sig")


def test_compound_operations_rerun_instead_of_forwarding_a_stage():
    names = (
        "edit_circuit_properties_and_recalculate_operation.py",
        "mark_existing_and_recalculate_operation.py",
        "set_include_and_recalculate_operation.py",
    )
    for name in names:
        source = _read(OPERATIONS / name)
        assert "use_existing_transaction_group" in source
        assert "staged_calculation" not in source
        assert "staged_result" not in source


def test_autosize_can_forward_its_explicit_builtin_stage():
    source = _read(OPERATIONS / "autosize_breaker_and_recalculate_operation.py")
    assert "staged_builtin_values_by_id" in source
    assert "staged_calculation" in source


def test_standalone_ribbon_uses_the_explicit_apply_operation():
    path = (
        ROOT
        / "CED ElecTools.extension"
        / "AE pyTools.tab"
        / "Electrical.panel"
        / "Circuits2.stack"
        / "Calculate Circuits.pushbutton"
        / "script.py"
    )
    source = _read(path)
    assert 'operation_key = "apply_calculated_circuits"' in source
    assert 'result.get("staged_calculation")' in source


def test_preview_operation_skips_pre_2026_legacy_special_elements():
    source = _read(OPERATIONS / "calculate_circuits_preview_operation.py")
    assert "get_special_processing_mode(doc, circuit) == SPECIAL_MODE_LEGACY" in source


if __name__ == "__main__":
    tests = sorted(
        (name, value)
        for name, value in globals().items()
        if name.startswith("test_") and callable(value)
    )
    for name, test in tests:
        test()
        print("PASS {}".format(name))
