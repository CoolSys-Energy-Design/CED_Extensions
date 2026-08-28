import importlib.util
from pathlib import Path
import unittest
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[2]
MODULE_PATH = ROOT / "CEDLib.lib" / "UIClasses" / "filterable_combo_box.py"
THEME_XAML_PATH = ROOT / (
    "AE pyTools.extension/AE pyTools.Tab/CED Tools.panel/"
    "Theme.pushbutton/ThemeWindow.xaml"
)
THEME_SCRIPT_PATH = ROOT / (
    "AE pyTools.extension/AE pyTools.Tab/CED Tools.panel/"
    "Theme.pushbutton/script.py"
)
RESOURCE_XAML_PATH = ROOT / (
    "CEDLib.lib/UIClasses/Resources/Controls/FilterableComboBox.xaml"
)
LOADER_PATH = ROOT / "CEDLib.lib" / "UIClasses" / "resource_loader.py"
CURSED_THEME_PATH = ROOT / "CEDLib.lib/UIClasses/Resources/Themes/CEDTheme.Cursed.xaml"
SETTINGS_XAML_PATH = ROOT / (
    "CED ElecTools.extension/AE pyTools.tab/Electrical.panel/"
    "Circuits2.stack/Calculate Circuits.pushbutton/settings.xaml"
)
SETTINGS_SCRIPT_PATH = ROOT / (
    "CED ElecTools.extension/AE pyTools.tab/Electrical.panel/"
    "Circuits2.stack/Calculate Circuits.pushbutton/config.py"
)
CIRCUIT_BROWSER_PATH = ROOT / (
    "CED ElecTools.extension/AE pyTools.tab/Electrical.panel/"
    "Circuit Manager.pushbutton/CircuitBrowserPanel.py"
)
SPEC = importlib.util.spec_from_file_location("ced_filterable_combo_box", str(MODULE_PATH))
FILTERABLE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(FILTERABLE)
WPF = "{http://schemas.microsoft.com/winfx/2006/xaml/presentation}"
XAML = "{http://schemas.microsoft.com/winfx/2006/xaml}"


class EventHook(object):
    def __init__(self):
        self.handlers = []

    def __iadd__(self, handler):
        self.handlers.append(handler)
        return self

    def __isub__(self, handler):
        if handler in self.handlers:
            self.handlers.remove(handler)
        return self

    def fire(self, sender, args):
        for handler in list(self.handlers):
            handler(sender, args)


class FakeItems(object):
    def __init__(self, source):
        self.source = list(source)
        self.visible = list(source)
        self.Filter = None

    @property
    def Count(self):
        return len(self.visible)

    def Refresh(self):
        if self.Filter is None:
            self.visible = list(self.source)
        else:
            self.visible = [item for item in self.source if self.Filter(item)]

    def __iter__(self):
        return iter(self.visible)

    def __len__(self):
        return len(self.visible)

    def __getitem__(self, index):
        return self.visible[index]


class FakeTemplate(object):
    def __init__(self, textbox, popup):
        self.textbox = textbox
        self.popup = popup

    def FindName(self, name, owner):
        if name == "PART_EditableTextBox":
            return self.textbox
        if name == "PART_Popup":
            return self.popup
        return None


class FakeContainer(object):
    def __init__(self, index):
        self.index = index
        self.PreviewKeyDown = EventHook()
        self.PreviewMouseLeftButtonDown = EventHook()
        self.DataContext = None
        self.IsHighlighted = False


class FakePopup(object):
    def __init__(self):
        self.IsKeyboardFocusWithin = False
        self.PreviewKeyDown = EventHook()
        self.PreviewMouseLeftButtonDown = EventHook()


class FakeGenerator(object):
    def __init__(self, items):
        self.containers = [FakeContainer(index) for index in range(len(items))]

    def ContainerFromIndex(self, index):
        try:
            return self.containers[index]
        except IndexError:
            return None

    def IndexFromContainer(self, container):
        return container.index


class FakeTextBox(object):
    def __init__(self, combo):
        self.combo = combo
        self._text = ""
        self.CaretIndex = 0
        self.SelectionStart = 0
        self.SelectionLength = 0
        self.TextChanged = EventHook()
        self.GotKeyboardFocus = EventHook()
        self.LostKeyboardFocus = EventHook()
        self.PreviewKeyDown = EventHook()
        self.PreviewMouseDoubleClick = EventHook()

    @property
    def Text(self):
        return self._text

    @Text.setter
    def Text(self, value):
        value = "" if value is None else str(value)
        if value == self._text:
            return
        self._text = value
        self.TextChanged.fire(self, object())

    def Focus(self):
        self.combo.IsKeyboardFocusWithin = True
        return True

    def Select(self, start, length):
        self.SelectionStart = start
        self.SelectionLength = length
        self.CaretIndex = start

    def SelectAll(self):
        self.Select(0, len(self.Text))

    def user_type(self, text):
        self.Text = text
        self.CaretIndex = len(text)


class FakeVisualTreeHelper(object):
    @staticmethod
    def GetParent(node):
        return getattr(node, "Parent", None)


class FakeCombo(object):
    def __init__(self, source):
        self.ItemsSource = list(source)
        self.Items = FakeItems(source)
        self._selected_item = None
        self._selected_index = -1
        self.Text = ""
        self.IsDropDownOpen = False
        self.IsKeyboardFocusWithin = True
        self.ItemTemplate = None
        self.ItemContainerGenerator = FakeGenerator(source)
        self.Loaded = EventHook()
        self.DropDownOpened = EventHook()
        self.DropDownClosed = EventHook()
        self.SelectionChanged = EventHook()
        self.GotKeyboardFocus = EventHook()
        self.LostKeyboardFocus = EventHook()
        self.PreviewKeyDown = EventHook()
        self.PreviewMouseLeftButtonDown = EventHook()
        self.popup = FakePopup()
        self.textbox = FakeTextBox(self)
        self.Template = FakeTemplate(self.textbox, self.popup)

    @property
    def SelectedItem(self):
        return self._selected_item

    @SelectedItem.setter
    def SelectedItem(self, value):
        changed = value != self._selected_item
        self._selected_item = value
        self._selected_index = (
            self.ItemsSource.index(value) if value in self.ItemsSource else -1
        )
        if changed:
            self.SelectionChanged.fire(self, object())

    @property
    def SelectedIndex(self):
        return self._selected_index

    @SelectedIndex.setter
    def SelectedIndex(self, value):
        previous = self._selected_item
        self._selected_index = int(value)
        self._selected_item = (
            self.ItemsSource[self._selected_index]
            if self._selected_index >= 0
            else None
        )
        if self._selected_item != previous:
            self.SelectionChanged.fire(self, object())

    def ApplyTemplate(self):
        return True

    def TryFindResource(self, key):
        return None

    def UpdateLayout(self):
        return None

    def open_dropdown(self):
        self.IsDropDownOpen = True
        self.DropDownOpened.fire(self, object())

    def click_outside(self):
        self.IsKeyboardFocusWithin = False
        self.popup.IsKeyboardFocusWithin = False
        was_open = self.IsDropDownOpen
        self.IsDropDownOpen = False
        if was_open:
            self.DropDownClosed.fire(self, object())
        self.LostKeyboardFocus.fire(self, object())


class FakeKeyboard(object):
    active_textbox = None

    @classmethod
    def Focus(cls, textbox):
        cls.active_textbox = textbox
        return textbox.Focus()

    @classmethod
    def ClearFocus(cls):
        if cls.active_textbox is not None:
            cls.active_textbox.combo.IsKeyboardFocusWithin = False


class KeyArgs(object):
    def __init__(self, key):
        self.Key = key
        self.Handled = False


class MouseArgs(object):
    def __init__(self, original_source):
        self.OriginalSource = original_source
        self.Handled = False


class FilterableComboBoxTests(unittest.TestCase):
    def setUp(self):
        self.original_keyboard = FILTERABLE.Keyboard
        self.original_combo_box_item = FILTERABLE.ComboBoxItem
        self.original_visual_tree_helper = FILTERABLE.VisualTreeHelper
        FILTERABLE.Keyboard = FakeKeyboard
        FILTERABLE.ComboBoxItem = FakeContainer
        FILTERABLE.VisualTreeHelper = FakeVisualTreeHelper

    def tearDown(self):
        FILTERABLE.Keyboard = self.original_keyboard
        FILTERABLE.ComboBoxItem = self.original_combo_box_item
        FILTERABLE.VisualTreeHelper = self.original_visual_tree_helper

    def _control(
        self,
        allow_custom_values,
        callback=None,
        fallback=FILTERABLE.FALLBACK_LAST_VALID,
        fallback_item=None,
    ):
        values = [
            "Lighting - Pendant",
            "Lighting - Recessed",
            "Power - Duplex Receptacle",
            "Power - Floor Box",
            "Mechanical - Return Grille",
        ]
        combo = FakeCombo(values)
        behavior = FILTERABLE.FilterableComboBox(
            combo,
            allow_custom_values=allow_custom_values,
            fallback=fallback,
            fallback_item=fallback_item,
            on_value_committed=callback,
        )
        return behavior, combo

    def _press(self, surface, key):
        args = KeyArgs(key)
        surface.PreviewKeyDown.fire(surface, args)
        return args

    def _click_row(self, combo, visible_index):
        row = combo.ItemContainerGenerator.ContainerFromIndex(visible_index)
        args = MouseArgs(row)
        combo.popup.PreviewMouseLeftButtonDown.fire(combo.popup, args)
        return args

    def test_filter_items_is_case_insensitive_contains_match(self):
        values = ["Lighting - Pendant", "Power - Floor Box", "Mechanical - Return Grille"]
        self.assertEqual(
            ["Lighting - Pendant", "Mechanical - Return Grille"],
            FILTERABLE.filter_items(values, "i"),
        )

    def test_filter_items_empty_query_restores_all_items(self):
        self.assertEqual(["Alpha", "Beta"], FILTERABLE.filter_items(["Alpha", "Beta"], "  "))

    def test_filter_items_supports_display_member_path(self):
        class Item(object):
            def __init__(self, label):
                self.label = label

        values = [Item("North"), Item("South")]
        self.assertEqual([values[1]], FILTERABLE.filter_items(values, "sou", display_member_path="label"))

    def test_highlight_segments_marks_case_insensitive_matches(self):
        self.assertEqual(
            [("Light", False), ("ing", True), (" - Pendant", False)],
            FILTERABLE.highlight_segments("Lighting - Pendant", "ING"),
        )

    def test_navigation_index_clamps_and_starts_at_visible_edges(self):
        self.assertEqual(0, FILTERABLE.navigation_index(-1, 3, 1))
        self.assertEqual(2, FILTERABLE.navigation_index(-1, 3, -1))
        self.assertEqual(2, FILTERABLE.navigation_index(2, 3, 1))
        self.assertEqual(0, FILTERABLE.navigation_index(0, 3, -1))
        self.assertEqual(-1, FILTERABLE.navigation_index(-1, 0, 1))

    def test_wpf_return_key_alias_is_treated_as_enter(self):
        behavior, _combo = self._control(False)
        self.assertTrue(behavior._key_is(KeyArgs("Return"), "Enter"))

    def test_find_exact_item_is_case_insensitive(self):
        values = ["Lighting - Pendant", "Power - Floor Box"]
        self.assertEqual("Power - Floor Box", FILTERABLE.find_exact_item(values, "power - floor box"))
        self.assertIsNone(FILTERABLE.find_exact_item(values, "Power"))

    def test_fallback_can_use_last_valid_or_configured_default(self):
        values = ["Alpha", "Beta"]
        self.assertEqual(
            "Beta",
            FILTERABLE.resolve_fallback_item(
                values,
                fallback=FILTERABLE.FALLBACK_LAST_VALID,
                last_valid_item="Beta",
            ),
        )
        self.assertEqual(
            "Alpha",
            FILTERABLE.resolve_fallback_item(
                values,
                fallback=FILTERABLE.FALLBACK_DEFAULT_ITEM,
                default_item="Alpha",
            ),
        )
        self.assertIsNone(
            FILTERABLE.resolve_fallback_item(
                values,
                fallback=FILTERABLE.FALLBACK_BLANK,
                last_valid_item="Beta",
            )
        )

    def test_enforced_mode_filters_replaces_text_and_commits_enter(self):
        commits = []
        behavior, combo = self._control(False, lambda control, value, item: commits.append((value, item)))
        combo.open_dropdown()
        combo.textbox.user_type("Power")
        self.assertEqual("Power", behavior.query)
        self.assertEqual("Power", combo.textbox.Text)
        self.assertEqual("Power - Duplex Receptacle", behavior._candidate_item)

        down = self._press(combo.textbox, "Down")
        self.assertTrue(down.Handled)
        self.assertEqual("Power - Duplex Receptacle", combo.Text)
        self.assertEqual("Power - Duplex Receptacle", combo.textbox.Text)
        self.assertEqual("Power", behavior.query)

        second_down = self._press(combo.textbox, "Down")
        self.assertTrue(second_down.Handled)
        self.assertEqual("Power - Floor Box", combo.Text)
        self.assertEqual("Power", behavior.query)

        enter = self._press(combo.textbox, "Enter")
        self.assertTrue(enter.Handled)
        self.assertEqual("Power - Floor Box", combo.SelectedItem)
        self.assertEqual("Power - Floor Box", behavior.selected_item)
        self.assertEqual(len("Power - Floor Box"), combo.textbox.CaretIndex)
        self.assertFalse(combo.IsDropDownOpen)
        self.assertFalse(combo.IsKeyboardFocusWithin)
        self.assertEqual([("Power - Floor Box", "Power - Floor Box")], commits)

    def test_arrow_opened_popup_routes_navigation_from_combo_and_rows(self):
        behavior, combo = self._control(False)
        combo.open_dropdown()
        self.assertTrue(combo.IsKeyboardFocusWithin)
        combo_args = self._press(combo, "Down")
        self.assertTrue(combo_args.Handled)
        self.assertEqual("Lighting - Pendant", combo.Text)

        row_args = self._press(combo.popup, "Down")
        self.assertTrue(row_args.Handled)
        self.assertEqual("Lighting - Recessed", combo.Text)
        up_args = self._press(combo.popup, "Up")
        self.assertTrue(up_args.Handled)
        self.assertEqual("Lighting - Pendant", combo.Text)

    def test_enforced_row_click_commits_caret_and_focus(self):
        commits = []
        behavior, combo = self._control(False, lambda control, value, item: commits.append(value))
        combo.open_dropdown()
        args = self._click_row(combo, 3)
        self.assertTrue(args.Handled)
        self.assertEqual("Power - Floor Box", combo.Text)
        self.assertEqual("Power - Floor Box", combo.SelectedItem)
        self.assertEqual(len(combo.Text), combo.textbox.CaretIndex)
        self.assertFalse(combo.IsKeyboardFocusWithin)
        self.assertEqual(["Power - Floor Box"], commits)

    def test_custom_mode_enter_keeps_typed_text_until_row_is_chosen(self):
        commits = []
        behavior, combo = self._control(True, lambda control, value, item: commits.append((value, item)))
        combo.open_dropdown()
        combo.textbox.user_type("brand new value")
        self.assertIsNone(behavior._candidate_item)
        enter = self._press(combo.textbox, "Enter")
        self.assertTrue(enter.Handled)
        self.assertEqual("brand new value", combo.Text)
        self.assertIsNone(combo.SelectedItem)
        self.assertFalse(combo.IsDropDownOpen)
        self.assertFalse(combo.IsKeyboardFocusWithin)
        self.assertEqual([("brand new value", None)], commits)

        behavior, combo = self._control(True)
        combo.open_dropdown()
        combo.textbox.user_type("Power")
        self.assertIsNone(behavior._candidate_item)
        down = self._press(combo.textbox, "Down")
        self.assertEqual("Power", combo.Text)
        self.assertIsNotNone(behavior._candidate_item)
        self._press(combo.textbox, "Enter")
        self.assertEqual("Power - Duplex Receptacle", combo.Text)

    def test_clicking_back_into_committed_text_selects_all_and_double_click_selects_all(self):
        behavior, combo = self._control(False)
        behavior._commit_item("Power - Floor Box")
        combo.textbox.SelectionStart = 2
        combo.textbox.SelectionLength = 0
        behavior._on_textbox_got_focus(combo.textbox, object())
        self.assertEqual((0, len(combo.Text)), (combo.textbox.SelectionStart, combo.textbox.SelectionLength))
        combo.textbox.SelectionStart = 2
        combo.textbox.SelectionLength = 0
        args = KeyArgs("DoubleClick")
        behavior._on_textbox_double_click(combo.textbox, args)
        self.assertEqual((0, len(combo.Text)), (combo.textbox.SelectionStart, combo.textbox.SelectionLength))
        self.assertTrue(args.Handled)

    def test_enforced_invalid_blur_uses_last_valid_blank_and_default_fallbacks(self):
        behavior, combo = self._control(False)
        behavior._commit_item("Lighting - Recessed")
        combo.IsKeyboardFocusWithin = True
        combo.textbox.user_type("not a list value")
        combo.click_outside()
        self.assertEqual("Lighting - Recessed", combo.SelectedItem)
        self.assertEqual("Lighting - Recessed", combo.Text)

        behavior, combo = self._control(False, fallback=FILTERABLE.FALLBACK_BLANK)
        combo.textbox.user_type("not a list value")
        combo.click_outside()
        self.assertIsNone(combo.SelectedItem)
        self.assertEqual("", combo.Text)

        behavior, combo = self._control(
            False,
            fallback=FILTERABLE.FALLBACK_DEFAULT_ITEM,
            fallback_item="Mechanical - Return Grille",
        )
        combo.textbox.user_type("not a list value")
        combo.click_outside()
        self.assertEqual("Mechanical - Return Grille", combo.SelectedItem)

    def test_delete_committed_value_restores_full_list_and_navigation(self):
        behavior, combo = self._control(False)
        behavior._commit_item("Power - Floor Box")
        combo.IsKeyboardFocusWithin = True
        combo.textbox.GotKeyboardFocus.fire(combo.textbox, object())
        combo.textbox.user_type("")
        self.assertEqual("", behavior.query)
        self.assertEqual("", combo.Text)
        self.assertEqual("", combo.textbox.Text)
        self.assertIsNone(combo.SelectedItem)
        self.assertEqual(combo.ItemsSource, list(combo.Items))
        self._press(combo.textbox, "Down")
        self.assertEqual("Lighting - Pendant", combo.Text)
        self.assertEqual("", behavior.query)

    def test_custom_text_blur_preserves_value_and_custom_row_click_replaces_it(self):
        commits = []
        behavior, combo = self._control(
            True,
            lambda control, value, item: commits.append((value, item)),
        )
        combo.open_dropdown()
        combo.textbox.user_type("free form")
        combo.click_outside()
        self.assertEqual("free form", combo.Text)
        self.assertEqual("free form", combo.textbox.Text)
        self.assertIsNone(combo.SelectedItem)
        self.assertEqual(("free form", None), commits[-1])

        combo.IsKeyboardFocusWithin = True
        combo.open_dropdown()
        combo.textbox.user_type("Power")
        args = self._click_row(combo, 1)
        self.assertTrue(args.Handled)
        self.assertEqual("Power - Floor Box", combo.SelectedItem)
        self.assertEqual("Power - Floor Box", combo.Text)
        self.assertEqual(("Power - Floor Box", "Power - Floor Box"), commits[-1])

    def test_clicking_already_highlighted_row_commits_it(self):
        behavior, combo = self._control(False)
        combo.open_dropdown()
        combo.textbox.user_type("Power")
        self._press(combo.textbox, "Down")
        self.assertEqual("Power - Duplex Receptacle", behavior._candidate_item)
        args = self._click_row(combo, 0)
        self.assertTrue(args.Handled)
        self.assertEqual("Power - Duplex Receptacle", combo.SelectedItem)
        self.assertFalse(combo.IsKeyboardFocusWithin)

    def test_custom_open_has_no_candidate_and_popup_down_enter_chooses_first_row(self):
        behavior, combo = self._control(True)
        combo.open_dropdown()
        self.assertIsNone(behavior._candidate_item)
        combo.textbox.user_type("Power")
        self._press(combo.popup, "Down")
        self.assertEqual("Power", combo.Text)
        self.assertEqual("Power - Duplex Receptacle", behavior._candidate_item)
        self._press(combo.popup, "Enter")
        self.assertEqual("Power - Duplex Receptacle", combo.SelectedItem)
        self.assertEqual("Power - Duplex Receptacle", combo.Text)

    def test_focus_and_double_click_events_select_the_whole_committed_value(self):
        behavior, combo = self._control(False)
        behavior._commit_item("Mechanical - Return Grille")
        combo.IsKeyboardFocusWithin = True
        combo.textbox.SelectionStart = 5
        combo.textbox.SelectionLength = 0
        combo.textbox.GotKeyboardFocus.fire(combo.textbox, object())
        self.assertEqual(0, combo.textbox.SelectionStart)
        self.assertEqual(len(combo.Text), combo.textbox.SelectionLength)
        combo.textbox.SelectionStart = 5
        combo.textbox.SelectionLength = 0
        args = MouseArgs(combo.textbox)
        combo.textbox.PreviewMouseDoubleClick.fire(combo.textbox, args)
        self.assertTrue(args.Handled)
        self.assertEqual(0, combo.textbox.SelectionStart)
        self.assertEqual(len(combo.Text), combo.textbox.SelectionLength)

    def test_single_click_refocus_selects_all_before_native_caret_placement(self):
        behavior, combo = self._control(False)
        behavior._commit_item("Power - Floor Box")
        combo.textbox.SelectionStart = 4
        combo.textbox.SelectionLength = 0
        args = MouseArgs(combo.textbox)
        combo.PreviewMouseLeftButtonDown.fire(combo, args)
        self.assertTrue(args.Handled)
        self.assertEqual(0, combo.textbox.SelectionStart)
        self.assertEqual(len(combo.Text), combo.textbox.SelectionLength)

    def test_implicit_enforced_candidate_does_not_replace_public_or_visible_query(self):
        behavior, combo = self._control(False)
        combo.open_dropdown()
        combo.textbox.user_type("Power")
        self.assertEqual("Power - Duplex Receptacle", behavior._candidate_item)
        self.assertFalse(behavior._candidate_engaged)
        self.assertEqual("Power", behavior.value)
        self.assertEqual("Power", combo.Text)

    def test_combo_and_editable_textbox_stay_synchronized_across_edit_navigation_commit(self):
        behavior, combo = self._control(False)
        combo.open_dropdown()
        combo.textbox.user_type("power")
        self.assertEqual(combo.Text, combo.textbox.Text)
        self._press(combo.textbox, "Down")
        self.assertEqual(combo.Text, combo.textbox.Text)
        self._press(combo.textbox, "Enter")
        self.assertEqual(combo.Text, combo.textbox.Text)
        self.assertEqual(combo.Text, combo.SelectedItem)

    def test_filterable_combo_resource_is_well_formed_xaml(self):
        root = ET.parse(str(RESOURCE_XAML_PATH)).getroot()
        self.assertEqual(WPF + "ResourceDictionary", root.tag)

    def test_theme_preview_uses_editable_filterable_combo(self):
        root = ET.parse(str(THEME_XAML_PATH)).getroot()
        combo = next(
            element
            for element in root.iter(WPF + "ComboBox")
            if element.attrib.get(XAML + "Name") == "PreviewFilterCombo"
        )
        self.assertEqual(
            "{DynamicResource CED.Input.ComboBox.Flat.Compact.Editable}",
            combo.attrib.get("Style"),
        )
        source = THEME_SCRIPT_PATH.read_text(encoding="utf-8")
        self.assertIn("FilterableComboBox(", source)
        self.assertIn("allow_custom_values=False", source)

    def test_shared_resource_uses_docs_search_highlight_brushes(self):
        source = RESOURCE_XAML_PATH.read_text(encoding="utf-8")
        self.assertIn("CED.Input.ComboBox.FilterableItemTemplate", source)
        helper_source = MODULE_PATH.read_text(encoding="utf-8")
        self.assertIn("CED.Brush.SearchHighlightBackground", helper_source)
        self.assertIn("CED.Brush.SearchHighlightForeground", helper_source)

    def test_theme_loader_exposes_descriptor_driven_modes(self):
        source = LOADER_PATH.read_text(encoding="utf-8")
        self.assertIn("THEME_DESCRIPTORS", source)
        self.assertIn("CEDTheme.Cursed.xaml", source)
        self.assertIn("def theme_descriptors", source)

    def test_all_theme_selectors_are_loader_driven(self):
        theme_window = THEME_XAML_PATH.read_text(encoding="utf-8")
        theme_script = THEME_SCRIPT_PATH.read_text(encoding="utf-8")
        settings_xaml = SETTINGS_XAML_PATH.read_text(encoding="utf-8")
        settings_script = SETTINGS_SCRIPT_PATH.read_text(encoding="utf-8")
        browser_script = CIRCUIT_BROWSER_PATH.read_text(encoding="utf-8")

        self.assertNotIn('<ListBoxItem Tag="light"', theme_window)
        self.assertIn("theme_picker.ItemsSource = resource_loader.theme_descriptors()", theme_script)
        self.assertNotIn('<ComboBoxItem Content="Light" Tag="light"', settings_xaml)
        self.assertIn("resource_loader.theme_descriptors()", settings_script)
        self.assertIn("for descriptor in resource_loader.theme_descriptors()", browser_script)

    def test_cursed_theme_overrides_every_base_color_token(self):
        key_namespace = "{http://schemas.microsoft.com/winfx/2006/xaml}"
        base_root = ET.parse(str(ROOT / "CEDLib.lib/UIClasses/Resources/Themes/CED.Colors.xaml")).getroot()
        cursed_root = ET.parse(str(CURSED_THEME_PATH)).getroot()
        base_keys = {
            element.attrib.get(key_namespace + "Key")
            for element in base_root
            if element.tag == WPF + "Color"
        }
        cursed_keys = {
            element.attrib.get(key_namespace + "Key")
            for element in cursed_root
            if element.tag == WPF + "Color"
        }
        self.assertTrue(base_keys.issubset(cursed_keys))


if __name__ == "__main__":
    unittest.main()
