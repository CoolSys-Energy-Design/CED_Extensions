import ast
from pathlib import Path
import unittest
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
PANEL_DIR = ROOT / (
    "CED ElecTools.extension/AE pyTools.tab/Electrical.panel/"
    "Circuit Manager.pushbutton"
)
XAML_PATH = PANEL_DIR / "CircuitBrowserPanel.xaml"
PYTHON_PATH = PANEL_DIR / "CircuitBrowserPanel.py"
WPF = "{http://schemas.microsoft.com/winfx/2006/xaml/presentation}"


def _method(source, name):
    start = source.index("    def {}(".format(name))
    end = source.find("\n    def ", start + 5)
    return source[start:] if end < 0 else source[start:end]


def _load_matchers(source):
    tree = ast.parse(source.lstrip("\ufeff"), filename=str(PYTHON_PATH))
    names = {
        "_parse_boolean_filter",
        "matches_panel",
        "matches_poles",
        "matches_rating",
        "matches_neutral",
        "matches_ig",
        "matches_free_text",
    }
    functions = [
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in names
    ]
    namespace = {}
    exec(compile(ast.Module(body=functions, type_ignores=[]), str(PYTHON_PATH), "exec"), namespace)
    return namespace


class CircuitBrowserStructuredSearchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = ET.parse(str(XAML_PATH)).getroot()
        cls.source = PYTHON_PATH.read_text(encoding="utf-8")

    def test_search_host_replaces_legacy_xaml_search_controls(self):
        host = next(
            element
            for element in self.root.iter(WPF + "Grid")
            if element.attrib.get("Name") == "SearchHost"
        )
        self.assertEqual("2", host.attrib.get("Grid.Row"))
        self.assertEqual("24", host.attrib.get("MinHeight"))
        self.assertEqual("0,0,0,4", host.attrib.get("Margin"))
        self.assertNotIn('Name="SearchBox"', XAML_PATH.read_text(encoding="utf-8"))
        self.assertNotIn('Name="SearchPlaceholderText"', XAML_PATH.read_text(encoding="utf-8"))
        self.assertNotIn('Name="ClearSearchButton"', XAML_PATH.read_text(encoding="utf-8"))
        for handler in (
            "search_changed",
            "search_got_focus",
            "search_lost_focus",
            "clear_search_clicked",
        ):
            self.assertNotIn(handler, XAML_PATH.read_text(encoding="utf-8"))

    def test_filters_and_query_callback_are_host_owned(self):
        self.assertIn(
            "from UIClasses import SearchFilterDefinition, StructuredSearchBox",
            self.source,
        )
        self.assertIn('SearchFilterDefinition(\n                    "panel",', self.source)
        self.assertIn('SearchFilterDefinition(\n                    "poles",', self.source)
        for key in ("rating", "neutral", "ig"):
            self.assertIn('SearchFilterDefinition(\n                    "{}",'.format(key), self.source)
        self.assertIn("matcher=matches_panel", self.source)
        self.assertIn("matcher=matches_poles", self.source)
        self.assertIn("matcher=matches_rating", self.source)
        self.assertIn("matcher=matches_neutral", self.source)
        self.assertIn("matcher=matches_ig", self.source)
        self.assertEqual(3, self.source.count("allow_multiple=True"))
        self.assertIn('active_placeholder="Type / to add a search field"', self.source)
        for style_key in (
            "CED.SearchBox.Token.Panel",
            "CED.SearchBox.Token.Neutral",
            "CED.SearchBox.Token.IG",
            "CED.SearchBox.Token.Rating",
            "CED.SearchBox.Token.Poles",
        ):
            self.assertIn(style_key, self.source)
        self.assertIn("self._search.refresh_resources(self)", self.source)
        self.assertIn("self._search.add_query_changed_handler(self._structured_search_changed)", self.source)
        self.assertIn('value_hint="yes or no"', self.source)

        callback = _method(self.source, "_structured_search_changed")
        self.assertIn("self._search_query = args.query", callback)
        self.assertIn("self._refresh_list()", callback)
        self.assertNotIn("_safe_load_items", callback)
        self.assertNotIn("Items.Refresh", callback)
        self.assertNotIn("_rebuild_full_content_widths", callback)

    def test_refresh_keeps_query_and_reloads_items(self):
        refresh = _method(self.source, "refresh_clicked")
        self.assertIn("self._safe_load_items()", refresh)
        self.assertNotIn("self._search_query =", refresh)

        refresh_list = _method(self.source, "_refresh_list")
        self.assertIn("query = self._search_query", refresh_list)
        self.assertIn("if query is not None and not query.is_empty", refresh_list)
        self.assertIn("query.matches(item, matches_free_text)", refresh_list)

    def test_ctrl_f_uses_public_search_control_methods(self):
        handler = _method(self.source, "panel_preview_key_down")
        self.assertIn("self._search.focus_input()", handler)
        self.assertIn("self._search.select_all()", handler)
        self.assertNotIn("self._search._input", handler)

        control_source = (ROOT / "CEDLib.lib" / "UIClasses" / "structured_search_box.py").read_text(
            encoding="utf-8"
        )
        self.assertIn("CED.Brush.ButtonStateOverlayHover", control_source)
        self.assertIn("MouseDoubleClick", control_source)
        self.assertIn("CED.SearchBox.TokenOperator", control_source)
        self.assertIn("TextAlignment.Center", control_source)
        self.assertIn('footer_text.Text = "Press Enter to add"', control_source)
        self.assertIn("ColumnDefinition", control_source)
        self.assertIn("SetResourceReference", control_source)
        self.assertIn("resource_owner=None", control_source)
        self.assertIn("self.PreviewMouseLeftButtonDown += self._search_surface_mouse_down", control_source)
        self.assertIn("command_input_text", control_source)
        self.assertIn("self._popup.Closed += self._popup_closed", control_source)

    def test_host_matchers_are_exact_case_insensitive_and_safe_while_typing(self):
        matchers = _load_matchers(self.source)

        class Item(object):
            panel = " Panel-A "
            sort_poles = 3
            rating_amps = 20
            has_neutral = True
            has_ig = False
            search_name = "panel-a receptacle"

        item = Item()
        self.assertTrue(matchers["matches_panel"](item, "panel-a"))
        self.assertFalse(matchers["matches_panel"](item, "panel"))
        self.assertTrue(matchers["matches_poles"](item, " 3 "))
        for value in ("", " ", "3x", None):
            self.assertFalse(matchers["matches_poles"](item, value))
        self.assertTrue(matchers["matches_rating"](item, "20A"))
        self.assertTrue(matchers["matches_rating"](item, "20 a"))
        self.assertFalse(matchers["matches_rating"](item, "30A"))
        self.assertTrue(matchers["matches_neutral"](item, "true"))
        self.assertTrue(matchers["matches_ig"](item, "no"))
        self.assertFalse(matchers["matches_ig"](item, "maybe"))
        self.assertTrue(matchers["matches_free_text"](item, "RECEPT"))
        self.assertTrue(matchers["matches_free_text"](item, "  "))


if __name__ == "__main__":
    unittest.main()
