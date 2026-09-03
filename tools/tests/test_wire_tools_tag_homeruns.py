# -*- coding: utf-8 -*-
"""Static contracts for Wire Tools homerun tag batch placement."""

import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SOURCE_PATH = (
    ROOT
    / "CED ElecTools.extension"
    / "AE pyTools.tab"
    / "Electrical.panel"
    / "Circuits1.stack"
    / "Wire Tools.pushbutton"
    / "lib"
    / "wire_tools_logic.py"
)
SOURCE = SOURCE_PATH.read_text(encoding="utf-8-sig")
TREE = ast.parse(SOURCE)


def _function_source(name):
    for node in ast.walk(TREE):
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return ast.get_source_segment(SOURCE, node)
    raise AssertionError("Missing function: {}".format(name))


def test_tag_creation_helpers_do_not_regenerate_per_wire():
    for name in (
        "_place_no_leader_tag",
        "_refine_no_leader_tag",
        "_place_leader_tag",
        "_create_wire_tag",
    ):
        assert "Regenerate(" not in _function_source(name)


def test_tag_homeruns_regenerates_once_between_creation_and_refinement():
    source = _function_source("tag_homeruns")
    assert source.count("document.Regenerate()") == 1
    assert "batch regeneration was unavailable" not in source
    assert "refinement_jobs.append" in source
    assert "for wire, open_connector, created_tag in refinement_jobs:" in source
    assert source.index("subtransaction.Commit()") < source.index("document.Regenerate()")
    assert source.index("document.Regenerate()") < source.index("_refine_no_leader_tag(")


def test_tag_homeruns_retains_per_wire_failure_isolation():
    source = _function_source("tag_homeruns")
    assert "DB.SubTransaction(document)" in source
    assert "subtransaction.RollBack()" in source
    assert '"Homerun tag creation failed: {}"' in source


if __name__ == "__main__":
    tests = sorted(
        (name, value)
        for name, value in globals().items()
        if name.startswith("test_") and callable(value)
    )
    for name, test in tests:
        test()
        print("PASS {}".format(name))
