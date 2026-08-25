# -*- coding: utf-8 -*-
"""Automated checks for documentation build and runtime-independent services."""

from __future__ import print_function

import hashlib
import os
import sys
import tempfile
import unittest

REPOSITORY_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
DOCUMENTATION_ROOT = os.path.join(REPOSITORY_ROOT, "docs", "user-guide")
TOOLS_ROOT = os.path.join(REPOSITORY_ROOT, "tools", "docs")
LIB_ROOT = os.path.join(REPOSITORY_ROOT, "CEDLib.lib")
for path in (TOOLS_ROOT, LIB_ROOT):
    if path not in sys.path:
        sys.path.insert(0, path)

from documentation_build import build_catalog, validate
from Documentation.catalog import Catalog, build_ribbon_tree
from Documentation.highlighting import highlight_segments, query_terms
from Documentation.history import NavigationHistory
from Documentation.markdown_parser import parse, parse_inlines
from Documentation.markdown_parser import runtime_frontmatter
from Documentation.pathing import DocumentationPathError, resolve_local_path


class DocumentationBuildTests(unittest.TestCase):
    def test_repository_documentation_validates(self):
        documents, issues = validate(DOCUMENTATION_ROOT, repository_root=REPOSITORY_ROOT)
        errors = [item.format() for item in issues if item.severity == "error"]
        self.assertGreaterEqual(len(documents), 20)
        self.assertEqual([], errors)

    def test_catalog_contains_full_text_and_headings(self):
        catalog, _issues = build_catalog(DOCUMENTATION_ROOT, repository_root=REPOSITORY_ROOT)
        zoom = next(item for item in catalog["documents"] if item["id"] == "ae-pytools-zoom-to-selection")
        self.assertIn("bounding", zoom["text"].lower())
        self.assertIn("what-it-does", [item["anchor"] for item in zoom["headings"]])
        self.assertEqual(64, len(zoom["source_sha256"]))
        with open(os.path.join(DOCUMENTATION_ROOT, zoom["path"]), "r", encoding="utf-8-sig") as stream:
            source = stream.read()
        self.assertEqual(hashlib.sha256(source.encode("utf-8")).hexdigest(), zoom["source_sha256"])
        self.assertNotEqual(hashlib.sha256((source + "stale").encode("utf-8")).hexdigest(), zoom["source_sha256"])


class CatalogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.catalog = Catalog.load(DOCUMENTATION_ROOT)

    def test_searches_page_content_not_only_keywords(self):
        results = self.catalog.search("bounding area")
        self.assertTrue(results)
        self.assertEqual("ae-pytools-zoom-to-selection", results[0]["id"])

    def test_filters_by_extension_and_ribbon(self):
        results = self.catalog.search(
            "selection",
            extension="AE pyTools",
            ribbon_path="AE pyTools > Selection > Selection",
        )
        self.assertTrue(results)
        self.assertTrue(all(item["extension"] == "AE pyTools" for item in results))
        self.assertTrue(all(item["ribbon_path"] == "AE pyTools > Selection > Selection" for item in results))

    def test_search_with_no_results_is_empty(self):
        self.assertEqual([], self.catalog.search("term-that-does-not-exist-8f71c0"))

    def test_full_page_title_ranks_before_cross_references(self):
        results = self.catalog.search("Zoom to Selection")
        self.assertTrue(results)
        self.assertEqual("ae-pytools-zoom-to-selection", results[0]["id"])

    def test_missing_catalog_has_useful_error(self):
        with tempfile.TemporaryDirectory() as temporary_root:
            with self.assertRaisesRegex(Exception, "catalog is missing"):
                Catalog.load(temporary_root)

    def test_ribbon_locations_stop_before_matching_command_name(self):
        docs_page = self.catalog.by_id["ae-pytools-docs"]
        self.assertEqual("AE pyTools > CED Tools", docs_page["ribbon_location"])

    def test_tree_is_derived_from_navigation_metadata(self):
        tree = build_ribbon_tree(self.catalog.documents)
        extension = next(group for group in tree["groups"] if group["label"] == "AE pyTools")
        labels = {group["label"] for group in extension["groups"]}
        self.assertTrue({"CED Tools", "Orientation", "Selection", "Revisions"}.issubset(labels))
        self.assertEqual("ae-pytools-index", extension["index"]["id"])

        electrical = next(group for group in tree["groups"] if group["label"] == "CED ElecTools")
        electrical_labels = {group["label"] for group in electrical["groups"]}
        self.assertTrue({"Circuit Tools", "QC Check"}.issubset(electrical_labels))
        self.assertEqual("ced-electools-index", electrical["index"]["id"])


class PathAndHistoryTests(unittest.TestCase):
    def test_paths_cannot_escape_documentation_root(self):
        with self.assertRaises(DocumentationPathError):
            resolve_local_path(DOCUMENTATION_ROOT, "../../README.md", must_exist=False)

    def test_relative_page_path_resolves_inside_root(self):
        current = os.path.join(DOCUMENTATION_ROOT, "ae-pytools", "selection.md")
        resolved = resolve_local_path(
            DOCUMENTATION_ROOT,
            "orientation.md",
            current_document=current,
        )
        self.assertTrue(resolved.endswith(os.path.join("ae-pytools", "orientation.md")))

    def test_history_discards_forward_branch(self):
        history = NavigationHistory()
        history.push("a.md")
        history.push("b.md", "heading")
        self.assertEqual(("a.md", ""), history.back())
        history.push("c.md")
        self.assertFalse(history.can_forward)
        self.assertEqual(("c.md", ""), history.current)


class MarkdownParserTests(unittest.TestCase):
    def test_fixture_exercises_guaranteed_blocks(self):
        fixture = os.path.join(DOCUMENTATION_ROOT, "_fixtures", "renderer-compatibility.md")
        with open(fixture, "r", encoding="utf-8-sig") as stream:
            blocks = parse(stream.read())
        block_types = {item["type"] for item in blocks}
        self.assertTrue(
            {"heading", "paragraph", "list", "table", "code", "blockquote", "alert"}.issubset(block_types)
        )
        self.assertEqual(
            {"NOTE", "TIP", "IMPORTANT", "WARNING", "CAUTION"},
            {item["alert_type"] for item in blocks if item["type"] == "alert"},
        )

    def test_inline_parser_supports_promised_inline_features(self):
        tokens = parse_inlines("**bold** *italic* `code` [link](other.md) ![image](a.png)")
        self.assertEqual(
            {"bold", "italic", "code", "link", "image"},
            {item["type"] for item in tokens if item["type"] != "text"},
        )

    def test_runtime_frontmatter_rejects_invalid_page(self):
        with self.assertRaises(ValueError):
            runtime_frontmatter("# Missing frontmatter")


class SearchHighlightTests(unittest.TestCase):
    def test_terms_are_unique_and_case_insensitive(self):
        self.assertEqual(["Zoom", "selection"], query_terms(" Zoom selection ZOOM "))

    def test_segments_preserve_text_and_mark_each_match(self):
        segments = highlight_segments("Zoom to Selection tools", "zoom selection")
        self.assertEqual("Zoom to Selection tools", "".join(value for value, _matched in segments))
        self.assertEqual(["Zoom", "Selection"], [value for value, matched in segments if matched])


if __name__ == "__main__":
    unittest.main()
