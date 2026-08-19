# -*- coding: utf-8 -*-
"""Global CED theme picker available from the non-electrical toolbar."""

import os
import sys

import clr

# Load every assembly this command relies on before importing WPF/.NET types.
# UIClasses does this too, but keeping the command self-contained avoids
# import-order differences between pyRevit engines and toolbar load states.
for _assembly in (
    "WindowsBase",
    "PresentationCore",
    "PresentationFramework",
    "System.Data",
):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass

from System import Action
from System.Data import DataTable

from pyrevit import forms, script

from UIClasses import CEDWindowBase, resource_loader
from UIClasses import (
    THEME_CONFIG_ACCENT_KEY,
    THEME_CONFIG_SECTION,
    THEME_CONFIG_THEME_KEY,
)


THIS_DIR = os.path.abspath(os.path.dirname(__file__))
XAML_PATH = os.path.join(THIS_DIR, "ThemeWindow.xaml")


THEME_LABELS = {
    "light": "Light",
    "dark": "Dark",
    "dark_alt": "Dark Alt",
}

_CIRCUIT_MANAGER_MODULE_NAME = "ced_electools_circuit_manager_panel"


def _save_theme_state(theme_mode, accent_mode):
    """Persist the shared theme state used by all CED tool windows."""
    cfg = script.get_config(THEME_CONFIG_SECTION)
    if cfg is None:
        return False
    cfg.set_option(
        THEME_CONFIG_THEME_KEY,
        resource_loader.normalize_theme_mode(theme_mode, "light"),
    )
    cfg.set_option(
        THEME_CONFIG_ACCENT_KEY,
        resource_loader.normalize_accent_mode(accent_mode, "blue"),
    )
    script.save_config()
    return True


def _refresh_circuit_manager_theme():
    """Apply the newly saved config to the live Circuit Manager pane, if open."""
    module = sys.modules.get(_CIRCUIT_MANAGER_MODULE_NAME)
    if module is None:
        return False
    panel_class = getattr(module, "CircuitBrowserPanel", None)
    if panel_class is None or not hasattr(panel_class, "get_instance"):
        return False
    try:
        panel = panel_class.get_instance()
    except Exception:
        return False
    if panel is None or not hasattr(panel, "_sync_theme_from_config"):
        return False

    def apply_saved_theme():
        panel._sync_theme_from_config(apply_if_changed=True)

    try:
        dispatcher = getattr(panel, "Dispatcher", None)
        if dispatcher is not None:
            dispatcher.BeginInvoke(Action(apply_saved_theme))
            return True
    except Exception:
        pass
    try:
        apply_saved_theme()
        return True
    except Exception:
        return False


class ThemePickerWindow(CEDWindowBase):
    """Displays one live swatch and applies a newly selected theme immediately."""

    theme_aware = True
    default_theme_mode = "light"
    default_accent_mode = "blue"

    def __init__(self):
        CEDWindowBase.__init__(self, xaml_source=XAML_PATH, theme_aware=True)
        self._is_loading = True
        self.theme_picker = self.FindName("ThemePicker")
        self.preview_title = self.FindName("PreviewTitle")
        self.preview_slider = self.FindName("PreviewSlider")
        self.preview_data_grid = self.FindName("PreviewDataGrid")
        self.current_theme_text = self.FindName("CurrentThemeText")
        self.close_button = self.FindName("CloseButton")

        self._load_preview_grid()
        self._select_current_theme()
        self._refresh_labels()

        self.theme_picker.SelectionChanged += self._on_theme_changed
        self.preview_slider.PreviewMouseLeftButtonDown += self._on_preview_slider_mouse_down
        self.close_button.Click += self._close_window
        self._is_loading = False

    def _select_current_theme(self):
        selected = None
        for item in self.theme_picker.Items:
            mode = resource_loader.normalize_theme_mode(
                getattr(item, "Tag", None),
                "light",
            )
            if mode == self._theme_mode:
                selected = item
                break
        if selected is not None:
            self.theme_picker.SelectedItem = selected

    def _selected_theme_mode(self):
        selected = self.theme_picker.SelectedItem
        return resource_loader.normalize_theme_mode(
            getattr(selected, "Tag", None),
            self._theme_mode,
        )

    def _load_preview_grid(self):
        """Supply a compact, read-only data-grid sample for the live swatch."""
        table = DataTable()
        table.Columns.Add("Name")
        table.Columns.Add("Status")
        table.Columns.Add("Value")
        for values in (
            ("Panel A", "Ready", "120 A"),
            ("Circuit 1", "Connected", "20 A"),
            ("Device 1", "Checked", "120 V"),
            ("Circuit 2", "Pending", "15 A"),
        ):
            row = table.NewRow()
            row["Name"] = values[0]
            row["Status"] = values[1]
            row["Value"] = values[2]
            table.Rows.Add(row)
        self._preview_grid_source = table
        self.preview_data_grid.ItemsSource = table.DefaultView

    def _on_preview_slider_mouse_down(self, sender, args):
        """Move the slider directly to a clicked track position."""
        if str(getattr(args, "ChangedButton", "")) != "Left":
            return
        try:
            width = float(sender.ActualWidth)
            if width <= 0:
                return
            point = args.GetPosition(sender)
            ratio = max(0.0, min(1.0, float(point.X) / width))
            sender.Value = sender.Minimum + ((sender.Maximum - sender.Minimum) * ratio)
            args.Handled = True
        except Exception:
            pass

    def _refresh_labels(self):
        label = THEME_LABELS.get(self._theme_mode, self._theme_mode)
        self.preview_title.Text = "{} theme preview".format(label)
        self.current_theme_text.Text = "{} is active and saved for CED tools.".format(label)

    def _on_theme_changed(self, sender, args):
        if self._is_loading:
            return
        selected_mode = self._selected_theme_mode()
        if selected_mode == self._theme_mode:
            return

        self._theme_mode = selected_mode
        saved = _save_theme_state(self._theme_mode, self._accent_mode)
        self.apply_ced_theme(
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )
        if saved:
            _refresh_circuit_manager_theme()
        self._refresh_labels()

    def _close_window(self, sender, args):
        self.Close()


def main():
    try:
        window = ThemePickerWindow()
        window.ShowDialog()
    except Exception as exc:
        script.get_logger().error("Theme picker failed to open: %s", exc)
        forms.alert(
            "The Theme picker could not open.\n\n{}".format(exc),
            title="CED Tool Theme",
            warn_icon=True,
        )


if __name__ == "__main__":
    main()
