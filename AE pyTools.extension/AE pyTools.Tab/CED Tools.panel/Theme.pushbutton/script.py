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

from System import Action, Boolean, Int32
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
    """Preview themes locally and persist the selected theme only on Apply."""

    theme_aware = True
    default_theme_mode = "light"
    default_accent_mode = "blue"

    def __init__(self):
        CEDWindowBase.__init__(self, xaml_source=XAML_PATH, theme_aware=True)
        self._saved_theme_mode = self._theme_mode
        self._saved_accent_mode = self._accent_mode
        self._is_loading = True
        self.theme_picker = self.FindName("ThemePicker")
        self.preview_title = self.FindName("PreviewTitle")
        self.preview_slider = self.FindName("PreviewSlider")
        self.preview_data_grid = self.FindName("PreviewDataGrid")
        self.preview_alternating_toggle = self.FindName("PreviewAlternatingToggle")
        self.current_theme_text = self.FindName("CurrentThemeText")
        self.apply_button = self.FindName("ApplyButton")

        self._load_preview_grid()
        self._apply_preview_grid_mode()
        self._select_current_theme()
        self._refresh_labels()

        self.theme_picker.SelectionChanged += self._on_theme_changed
        self.preview_alternating_toggle.Checked += self._on_alternating_rows_changed
        self.preview_alternating_toggle.Unchecked += self._on_alternating_rows_changed
        self.preview_slider.PreviewMouseLeftButtonDown += self._on_preview_slider_mouse_down
        self.apply_button.Click += self._apply_theme_clicked
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
        """Supply representative shared cell and row states for the live swatch."""
        table = DataTable()
        table.Columns.Add("sample")
        table.Columns.Add("value")
        table.Columns.Add("new_qty")
        table.Columns.Add("state")
        table.Columns.Add("is_enabled", Boolean)
        table.Columns.Add("new_qty_changed", Boolean)
        table.Columns.Add("new_rating_changed", Boolean)
        table.Columns.Add("rating_warning_level", Int32)
        for values in (
            ("Normal", "Ready", "120 A", "Normal", True, False, False, 0),
            ("Hidden / disabled", "Suppressed", "20 A", "Hidden", False, False, False, 0),
            ("Changed", "Pending", "15 A", "Changed", True, True, False, 0),
            ("Warning", "Review", "80 A", "Warning", True, False, True, 1),
            ("Error", "Blocked", "0 A", "Error", True, False, True, 2),
        ):
            row = table.NewRow()
            row["sample"] = values[0]
            row["value"] = values[1]
            row["new_qty"] = values[2]
            row["state"] = values[3]
            row["is_enabled"] = values[4]
            row["new_qty_changed"] = values[5]
            row["new_rating_changed"] = values[6]
            row["rating_warning_level"] = values[7]
            table.Rows.Add(row)
        self._preview_grid_source = table
        self.preview_data_grid.ItemsSource = table.DefaultView

    def _apply_preview_grid_mode(self):
        """Switch the preview between shared flat and alternating resource variants."""
        alternating = bool(self.preview_alternating_toggle.IsChecked)
        grid_key = "CED.DataGrid.Display.Alternating" if alternating else "CED.DataGrid.Display.Flat"
        row_key = "CED.DataGrid.RowDisabledAware.Alternating" if alternating else "CED.DataGrid.RowDisabledAware"
        self.preview_data_grid.Style = resource_loader.try_find_resource(self, grid_key)
        self.preview_data_grid.RowStyle = resource_loader.try_find_resource(self, row_key)

        cell_style_pairs = (
            ("CED.DataGrid.Cell.DisplayReadonly", "CED.DataGrid.Cell.DisplayReadonly.Alternating"),
            ("CED.DataGrid.Cell.DisplayReadonly", "CED.DataGrid.Cell.DisplayReadonly.Alternating"),
            ("ReadonlyCell.NeutralSelection", "ReadonlyCell.NeutralSelection.Alternating"),
            ("CED.DataGrid.Cell.DisplayEditable", "CED.DataGrid.Cell.DisplayEditable.Alternating"),
            ("NewQtyCell", "NewQtyCell.Alternating"),
            ("NewValueCell", "NewValueCell.Alternating"),
        )
        for index, (flat_key, alternating_key) in enumerate(cell_style_pairs):
            key = alternating_key if alternating else flat_key
            self.preview_data_grid.Columns[index].CellStyle = resource_loader.try_find_resource(self, key)

    def _on_alternating_rows_changed(self, sender, args):
        if self._is_loading:
            return
        self._apply_preview_grid_mode()

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
        if (
            self._theme_mode == self._saved_theme_mode
            and self._accent_mode == self._saved_accent_mode
        ):
            self.current_theme_text.Text = "{} is active and saved for CED tools.".format(label)
        else:
            self.current_theme_text.Text = "{} is previewing. Click Apply to save for CED tools.".format(label)

    def _on_theme_changed(self, sender, args):
        if self._is_loading:
            return
        selected_mode = self._selected_theme_mode()
        if selected_mode == self._theme_mode:
            return

        self._theme_mode = selected_mode
        self.apply_ced_theme(
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )
        self._apply_preview_grid_mode()
        self._refresh_labels()

    def _apply_theme_clicked(self, sender, args):
        """Persist the previewed theme and refresh any live theme-aware pane."""
        if (
            self._theme_mode == self._saved_theme_mode
            and self._accent_mode == self._saved_accent_mode
        ):
            self._refresh_labels()
            return

        saved = _save_theme_state(self._theme_mode, self._accent_mode)
        if not saved:
            forms.alert(
                "The theme could not be saved to the pyRevit configuration.",
                title="CED Tool Theme",
                warn_icon=True,
            )
            return

        self._saved_theme_mode = self._theme_mode
        self._saved_accent_mode = self._accent_mode
        _refresh_circuit_manager_theme()
        self._refresh_labels()


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
