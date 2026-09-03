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
XAML = "{http://schemas.microsoft.com/winfx/2006/xaml}"


class CircuitBrowserLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = ET.parse(str(XAML_PATH)).getroot()

    def test_virtualization_and_recycling_remain_enabled(self):
        circuit_list = next(
            element
            for element in self.root.iter(WPF + "ListView")
            if element.attrib.get("Name") == "CircuitList"
        )
        self.assertEqual("True", circuit_list.attrib.get("VirtualizingPanel.IsVirtualizing"))
        self.assertEqual("Recycling", circuit_list.attrib.get("VirtualizingPanel.VirtualizationMode"))
        self.assertEqual("True", circuit_list.attrib.get("ScrollViewer.CanContentScroll"))

    def test_item_containers_use_one_shared_width(self):
        shared_style = next(
            element
            for element in self.root.iter(WPF + "Style")
            if element.attrib.get(XAML + "Key") == "CircuitManager.ListViewItem.SharedWidth"
        )
        setters = {
            setter.attrib.get("Property"): setter.attrib.get("Value", "")
            for setter in shared_style.findall(WPF + "Setter")
        }
        self.assertEqual("Left", setters.get("HorizontalAlignment"))
        self.assertEqual("Stretch", setters.get("HorizontalContentAlignment"))
        self.assertIn("Path=DataContext", setters.get("Width", ""))
        self.assertIn("AncestorType={x:Type ListView}", setters.get("Width", ""))

        circuit_list = next(
            element
            for element in self.root.iter(WPF + "ListView")
            if element.attrib.get("Name") == "CircuitList"
        )
        self.assertEqual(
            "{StaticResource CircuitManager.ListViewItem.SharedWidth}",
            circuit_list.attrib.get("ItemContainerStyle"),
        )
        self.assertFalse(any(self.root.iter(WPF + "ItemsPanelTemplate")))

        for template_key in ("CompactTemplate", "CardTemplate"):
            template = next(
                element
                for element in self.root.iter(WPF + "DataTemplate")
                if element.attrib.get(XAML + "Key") == template_key
            )
            row_border = template.find(WPF + "Border")
            self.assertNotIn("Width", row_border.attrib)
            self.assertNotIn("MaxWidth", row_border.attrib)

    def test_runtime_theme_path_reuses_shared_width_style(self):
        source = PYTHON_PATH.read_text(encoding="utf-8")
        self.assertEqual(
            2,
            source.count(
                '_try_find_resource(self, "CircuitManager.ListViewItem.SharedWidth")'
            ),
        )
        self.assertNotIn(
            '_try_find_resource(self, "CED.ListViewItem.SurfaceBehavior")',
            source,
        )
        self.assertNotIn("self._list.ItemContainerStyle = None", source)

    def test_full_width_uses_all_items_and_conditional_accessories(self):
        source = PYTHON_PATH.read_text(encoding="utf-8")
        rebuild_start = source.index("    def _rebuild_full_content_widths(self):")
        rebuild_end = source.index("\n    def ", rebuild_start + 5)
        rebuild = source[rebuild_start:rebuild_end]
        self.assertIn("self._all_items", rebuild)

        compact_start = source.index("    def _compute_compact_intrinsic_width(")
        compact_end = source.index("\n    def ", compact_start + 5)
        compact = source[compact_start:compact_end]
        for property_name in (
            "neutral_badge_visibility",
            "ig_badge_visibility",
            "override_badge_visibility",
            "sync_lock_badge_visibility",
            "alert_visibility",
        ):
            self.assertIn(property_name, compact)

    def test_width_path_does_not_refresh_or_walk_realized_rows(self):
        source = PYTHON_PATH.read_text(encoding="utf-8")
        apply_start = source.index("    def _apply_shared_row_width(self):")
        apply_end = source.index("\n    def ", apply_start + 5)
        apply_width = source[apply_start:apply_end]
        self.assertNotIn("Items.Refresh", apply_width)
        self.assertNotIn("UpdateLayout", apply_width)
        self.assertNotIn("ListViewItem", apply_width)
        self.assertNotIn("_visible_items", apply_width)

    def test_load_name_uses_wpf_word_ellipsis(self):
        load_name_blocks = [
            element
            for element in self.root.iter(WPF + "TextBlock")
            if element.attrib.get("Text") == "{Binding load_name}"
        ]
        self.assertEqual(2, len(load_name_blocks))
        for block in load_name_blocks:
            self.assertEqual("WordEllipsis", block.attrib.get("TextTrimming"))

    def test_toolbar_refresh_still_performs_a_full_model_reload(self):
        refresh_button = next(
            element
            for element in self.root.iter(WPF + "Button")
            if element.attrib.get("Name") == "RefreshButton"
        )
        self.assertEqual("refresh_clicked", refresh_button.attrib.get("Click"))

        source = PYTHON_PATH.read_text(encoding="utf-8")
        handler_start = source.index("    def refresh_clicked(self, sender, args):")
        handler_end = source.index("\n    def ", handler_start + 5)
        handler = source[handler_start:handler_end]
        self.assertIn("self._safe_load_items()", handler)


if __name__ == "__main__":
    unittest.main()
