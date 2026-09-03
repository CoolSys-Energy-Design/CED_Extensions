# -*- coding: utf-8 -*-
"""Static and pure-Python checks for shared theme unlocking."""

import importlib.util
import os
import xml.etree.ElementTree as ET


REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
MANAGER_PATH = os.path.join(REPO_ROOT, "CEDLib.lib", "UIClasses", "theme_manager.py")
THEMES_ROOT = os.path.join(REPO_ROOT, "CEDLib.lib", "UIClasses", "Resources", "Themes")
THEME_SCRIPT = os.path.join(
    REPO_ROOT,
    "AE pyTools.extension",
    "AE pyTools.Tab",
    "CED Tools.panel",
    "Theme.pushbutton",
    "script.py",
)
THEME_XAML = os.path.join(
    REPO_ROOT,
    "AE pyTools.extension",
    "AE pyTools.Tab",
    "CED Tools.panel",
    "Theme.pushbutton",
    "ThemeWindow.xaml",
)
THEME_BRUSHES_XAML = os.path.join(
    REPO_ROOT, "CEDLib.lib", "UIClasses", "Resources", "Themes", "CED.Brushes.xaml"
)
INPUT_STYLES_XAML = os.path.join(
    REPO_ROOT, "CEDLib.lib", "UIClasses", "Resources", "Styles", "InputStyles.xaml"
)


def _load_manager():
    spec = importlib.util.spec_from_file_location("ced_theme_manager_test", MANAGER_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FakeConfig(object):
    def __init__(self, values=None):
        self.values = dict(values or {})

    def get_option(self, name, default=None):
        return self.values.get(name, default)

    def set_option(self, name, value):
        self.values[name] = value


def _modes(descriptors):
    return [item.Mode for item in descriptors]


def _color_keys(path):
    namespace = "{http://schemas.microsoft.com/winfx/2006/xaml/presentation}"
    key_attr = "{http://schemas.microsoft.com/winfx/2006/xaml}Key"
    root = ET.parse(path).getroot()
    return set(item.attrib[key_attr] for item in root.findall(namespace + "Color"))


def test_locked_themes_are_absent_without_config_unlocks():
    manager = _load_manager()
    assert manager.DEFAULT_ACCENT_MODE == "neutral"
    cfg = FakeConfig({manager.THEME_UNLOCK_SCHEMA_KEY: manager.THEME_UNLOCK_SCHEMA_VERSION})
    modes = _modes(manager.theme_descriptors(config=cfg, save_callback=lambda: None))
    assert modes == ["light", "dark", "dark_alt"]
    assert "cursed" not in modes
    assert "chatgpt" not in modes


def test_aislop_unlocks_chatgpt_and_survives_config_reads():
    manager = _load_manager()
    cfg = FakeConfig({manager.THEME_UNLOCK_SCHEMA_KEY: manager.THEME_UNLOCK_SCHEMA_VERSION})
    code = manager.match_unlock_code(list("xxaislop"))
    assert code is not None
    assert code.ThemeMode == "chatgpt"
    descriptor = manager.unlock_theme(code, config=cfg, save_callback=lambda: None)
    assert descriptor.Label == "chatGPT"
    assert "chatgpt" in _modes(manager.theme_descriptors(config=cfg, save_callback=lambda: None))


def test_code_matching_uses_suffix_and_normalizes_konami_enter():
    manager = _load_manager()
    keys = ["Q", "Up", "Up", "Down", "Down", "Left", "Right", "Left", "Right", "B", "A", "Return"]
    code = manager.match_unlock_code(keys)
    assert code is not None
    assert code.ThemeMode == "cursed"


def test_key_debouncer_emits_once_until_release_and_ignores_repeat():
    manager = _load_manager()
    debouncer = manager.UnlockKeyDebouncer()
    assert debouncer.press("A") == "A"
    assert debouncer.press("A") is None
    assert debouncer.press("A", is_repeat=True) is None
    debouncer.release("A")
    assert debouncer.press("A") == "A"
    debouncer.reset()
    assert debouncer.press("A") == "A"


def test_cleared_config_relocks_every_hidden_theme():
    manager = _load_manager()
    unlocked = FakeConfig({manager.THEME_UNLOCK_SCHEMA_KEY: manager.THEME_UNLOCK_SCHEMA_VERSION})
    manager.unlock_theme("chatgpt_aislop", config=unlocked, save_callback=lambda: None)
    assert manager.is_theme_unlocked("chatgpt", config=unlocked)

    cleared = FakeConfig({manager.THEME_UNLOCK_SCHEMA_KEY: manager.THEME_UNLOCK_SCHEMA_VERSION})
    assert not manager.is_theme_unlocked("chatgpt", config=cleared)
    assert "chatgpt" not in _modes(manager.theme_descriptors(config=cleared, save_callback=lambda: None))


def test_existing_cursed_selection_is_grandfathered_once_but_chatgpt_is_not():
    manager = _load_manager()
    cursed_cfg = FakeConfig({manager.THEME_CONFIG_THEME_KEY: "cursed"})
    cursed_modes = _modes(manager.theme_descriptors(config=cursed_cfg, save_callback=lambda: None))
    assert "cursed" in cursed_modes

    chatgpt_cfg = FakeConfig({manager.THEME_CONFIG_THEME_KEY: "chatgpt"})
    chatgpt_modes = _modes(manager.theme_descriptors(config=chatgpt_cfg, save_callback=lambda: None))
    assert "chatgpt" not in chatgpt_modes


def test_locked_config_value_cannot_apply_theme():
    manager = _load_manager()
    cfg = FakeConfig({manager.THEME_UNLOCK_SCHEMA_KEY: manager.THEME_UNLOCK_SCHEMA_VERSION})
    assert manager.available_theme_mode("chatgpt", "light", config=cfg) == "light"


def test_chatgpt_theme_is_complete_and_valid_xaml():
    light = os.path.join(THEMES_ROOT, "CED.Colors.xaml")
    chatgpt = os.path.join(THEMES_ROOT, "CEDTheme.ChatGPT.xaml")
    assert _color_keys(light).issubset(_color_keys(chatgpt))


def test_theme_window_uses_global_input_capture_and_shared_manager():
    with open(THEME_SCRIPT, "r") as stream:
        source = stream.read()
    assert "InputManager.Current.PreProcessInput" in source
    assert "theme_manager.match_unlock_code" in source
    assert "theme_manager.unlock_theme" in source
    assert "Dispatcher.BeginInvoke" in source
    assert "_show_unlock_toast" in source
    assert "DoubleAnimation" in source
    assert "ThemeUnlockAlertWindow" not in source
    assert "forms.alert(code.SuccessMessage" not in source


def test_unlock_success_toast_uses_ced_style_inside_scrollable_picker():
    root = ET.parse(THEME_XAML).getroot()
    source = open(THEME_XAML, "r").read()
    assert root.tag.endswith("Window")
    assert 'x:Name="UnlockToast"' in source
    assert "CED.Alert.Success" in source
    assert "CED.Alert.Text.Success" in source
    assert "CED.Alert.Icon.Success" in source
    assert source.count('VerticalScrollBarVisibility="Auto"') == 1
    assert 'Text="Read Only"' in source
    assert 'Content="Start"' not in source
    assert 'Content="Clear"' not in source
    assert 'x:Name="ClearUnlockButton"' in source
    assert "CED.Button.IconSmall" in source
    assert "CED.Icon.Close" in source
    assert "CED.Input.TextBox.ReadOnly" in source
    assert 'x:Name="PreviewFilterStatus"\n                       Grid.Row="3"' in source
    assert 'Text="Read Only"\n                       Grid.Row="4"' in source
    assert 'Height="154"' in source
    assert 'Text="Other controls"' not in source


def test_readonly_textbox_uses_shared_theme_brushes_and_style():
    ET.parse(THEME_BRUSHES_XAML)
    ET.parse(INPUT_STYLES_XAML)
    brushes = open(THEME_BRUSHES_XAML, "r").read()
    styles = open(INPUT_STYLES_XAML, "r").read()
    assert "CED.Brush.InputReadOnlyBackground" in brushes
    assert "CED.Brush.InputReadOnlyBorder" in brushes
    assert "CED.Brush.InputReadOnlyForeground" in brushes
    assert 'x:Key="CED.Input.TextBox.ReadOnly"' in styles
    assert "CED.Brush.InputReadOnlyBackground" in styles
    assert "CED.Brush.InputReadOnlyBorder" in styles
    assert "CED.Brush.InputReadOnlyForeground" in styles
    assert '<Setter Property="MinHeight" Value="28"/>' in styles
