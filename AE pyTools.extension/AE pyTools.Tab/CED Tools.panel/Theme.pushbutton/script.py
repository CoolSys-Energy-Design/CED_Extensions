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

from System import Action, Boolean, Int32, TimeSpan
from System.Data import DataTable
from System.Windows import Duration, UIElement, Visibility
from System.Windows.Input import InputManager, Keyboard, KeyEventArgs
from System.Windows.Media.Animation import DoubleAnimation

from pyrevit import forms, script

from UIClasses import (
    CEDWindowBase,
    FALLBACK_LAST_VALID,
    FilterableComboBox,
    resource_loader,
    theme_manager,
)
from UIClasses import (
    THEME_CONFIG_ACCENT_KEY,
    THEME_CONFIG_SECTION,
    THEME_CONFIG_THEME_KEY,
)


THIS_DIR = os.path.abspath(os.path.dirname(__file__))
XAML_PATH = os.path.join(THIS_DIR, "ThemeWindow.xaml")


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
        resource_loader.normalize_accent_mode(accent_mode, "neutral"),
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
    default_accent_mode = "neutral"

    def __init__(self):
        CEDWindowBase.__init__(self, xaml_source=XAML_PATH, theme_aware=True)
        self._saved_theme_mode = self._theme_mode
        self._saved_accent_mode = self._accent_mode
        self._is_loading = True
        self.theme_picker = self.FindName("ThemePicker")
        self.preview_title = self.FindName("PreviewTitle")
        self.preview_slider = self.FindName("PreviewSlider")
        self.preview_filter_combo = self.FindName("PreviewFilterCombo")
        self.preview_filter_status = self.FindName("PreviewFilterStatus")
        self.preview_data_grid = self.FindName("PreviewDataGrid")
        self.preview_alternating_toggle = self.FindName("PreviewAlternatingToggle")
        self.current_theme_text = self.FindName("CurrentThemeText")
        self.apply_button = self.FindName("ApplyButton")
        self.unlock_sequence_text = self.FindName("UnlockSequenceTextBox")
        self.start_unlock_button = self.FindName("StartUnlockButton")
        self.clear_unlock_button = self.FindName("ClearUnlockButton")
        self.unlock_toast = self.FindName("UnlockToast")
        self.unlock_toast_message = self.FindName("UnlockToastMessage")
        self._unlock_toast_animation = None
        self._unlock_keys = []
        self._unlock_key_debouncer = theme_manager.UnlockKeyDebouncer()
        self._unlock_listening = False
        self._unlock_input_attached = False

        self._load_theme_options()
        self._load_filterable_combo()
        self._load_preview_grid()
        self._apply_preview_grid_mode()
        self._select_current_theme()
        self._refresh_labels()

        self.theme_picker.SelectionChanged += self._on_theme_changed
        self.preview_alternating_toggle.Checked += self._on_alternating_rows_changed
        self.preview_alternating_toggle.Unchecked += self._on_alternating_rows_changed
        self.preview_slider.PreviewMouseLeftButtonDown += self._on_preview_slider_mouse_down
        self.apply_button.Click += self._apply_theme_clicked
        self.start_unlock_button.Click += self._start_unlock_capture
        self.clear_unlock_button.Click += self._clear_unlock_capture
        self.Closed += self._on_window_closed
        self._is_loading = False

    def _load_theme_options(self):
        """Populate the selector from the shared loader-owned descriptors."""
        selected_mode = getattr(self, "_theme_mode", "light")
        self.theme_picker.ItemsSource = resource_loader.theme_descriptors()
        selected = None
        for item in self.theme_picker.Items:
            if resource_loader.theme_descriptor_mode(item, "light") == selected_mode:
                selected = item
                break
        if selected is not None:
            self.theme_picker.SelectedItem = selected

    def _attach_unlock_input(self):
        if self._unlock_input_attached:
            return
        InputManager.Current.PreProcessInput += self._on_unlock_input
        self._unlock_input_attached = True

    def _detach_unlock_input(self):
        if not self._unlock_input_attached:
            return
        try:
            InputManager.Current.PreProcessInput -= self._on_unlock_input
        except Exception:
            pass
        self._unlock_input_attached = False

    def _start_unlock_capture(self, sender, args):
        """Begin window-scoped key capture without exposing known codes."""
        self._unlock_keys = []
        self._unlock_key_debouncer.reset()
        self.unlock_sequence_text.Text = ""
        self._unlock_listening = True
        self.start_unlock_button.Content = "◆"
        self._attach_unlock_input()
        try:
            self.unlock_sequence_text.Focus()
        except Exception:
            pass

    def _clear_unlock_capture(self, sender, args):
        """Clear the displayed attempt while leaving active capture running."""
        self._unlock_keys = []
        self._unlock_key_debouncer.reset()
        self.unlock_sequence_text.Text = ""

    def _on_unlock_input(self, sender, args):
        if not self._unlock_listening or not bool(getattr(self, "IsActive", False)):
            return
        try:
            input_event = args.StagingItem.Input
        except Exception:
            return
        if not isinstance(input_event, KeyEventArgs):
            return
        key_value = getattr(input_event, "Key", None)
        if str(key_value) == "System":
            key_value = getattr(input_event, "SystemKey", key_value)
        routed_event = getattr(input_event, "RoutedEvent", None)
        if routed_event == Keyboard.KeyUpEvent:
            self._unlock_key_debouncer.release(key_value)
            return
        if routed_event != Keyboard.KeyDownEvent:
            return
        key = self._unlock_key_debouncer.press(
            key_value,
            is_repeat=bool(getattr(input_event, "IsRepeat", False)),
        )
        if not key:
            return
        self._unlock_keys.append(key)
        self.unlock_sequence_text.Text = " ".join(
            theme_manager.display_unlock_key(item)
            for item in self._unlock_keys
        )
        code = theme_manager.match_unlock_code(self._unlock_keys)
        if code is None:
            return
        try:
            input_event.Handled = True
        except Exception:
            pass
        self._unlock_listening = False
        self._unlock_key_debouncer.reset()
        self.start_unlock_button.Content = "◇"
        self._detach_unlock_input()
        self.Dispatcher.BeginInvoke(Action(lambda: self._complete_theme_unlock(code)))

    def _complete_theme_unlock(self, code):
        """Persist and display an unlock after WPF finishes routing the key event."""
        descriptor = theme_manager.unlock_theme(code)
        if descriptor is None:
            forms.alert(
                "The theme code was recognized, but the unlock could not be saved.",
                title="CED Tool Theme",
                warn_icon=True,
            )
            return
        self._is_loading = True
        self._load_theme_options()
        self._is_loading = False
        self._show_unlock_toast(code.SuccessMessage)

    def _show_unlock_toast(self, message):
        """Overlay a themed success message, then fade it away in place."""
        try:
            self.unlock_toast.BeginAnimation(UIElement.OpacityProperty, None)
        except Exception:
            pass
        self.unlock_toast_message.Text = str(message or "Theme unlocked.")
        self.unlock_toast.Visibility = Visibility.Visible
        self.unlock_toast.Opacity = 1.0

        fade = DoubleAnimation()
        fade.From = 1.0
        fade.To = 0.0
        fade.BeginTime = TimeSpan.FromSeconds(2.0)
        fade.Duration = Duration(TimeSpan.FromMilliseconds(500.0))
        fade.Completed += self._hide_unlock_toast
        self._unlock_toast_animation = fade
        self.unlock_toast.BeginAnimation(UIElement.OpacityProperty, fade)

    def _hide_unlock_toast(self, sender, args):
        self.unlock_toast.Visibility = Visibility.Collapsed
        self._unlock_toast_animation = None

    def _on_window_closed(self, sender, args):
        self._unlock_listening = False
        self._unlock_key_debouncer.reset()
        self._detach_unlock_input()

    def _load_filterable_combo(self):
        """Load a realistic option set into the editable filter demo."""
        self.preview_filter_combo.ItemsSource = [
            "Lighting - Pendant",
            "Lighting - Recessed",
            "Lighting - Wall Sconce",
            "Power - Duplex Receptacle",
            "Power - Floor Box",
            "Mechanical - Return Grille",
            "Mechanical - Supply Diffuser",
            "Security - Card Reader",
        ]
        self.preview_filter_combo.SelectedIndex = -1
        self._filterable_combo = FilterableComboBox(
            self.preview_filter_combo,
            on_filter_changed=self._on_filterable_combo_changed,
            allow_custom_values=False,
            fallback=FALLBACK_LAST_VALID,
        )
        self.preview_filter_status.Text = "8 choices available. Type to filter."

    def _on_filterable_combo_changed(self, behavior, query, count):
        if query:
            self.preview_filter_status.Text = (
                "Showing {0} match{1} for \"{2}\"."
                .format(count, "" if count == 1 else "es", query)
            )
        else:
            self.preview_filter_status.Text = (
                "{0} choices available. Type to filter.".format(count)
            )

    def _select_current_theme(self):
        selected = None
        for item in self.theme_picker.Items:
            mode = resource_loader.theme_descriptor_mode(item, "light")
            if mode == self._theme_mode:
                selected = item
                break
        if selected is not None:
            self.theme_picker.SelectedItem = selected

    def _selected_theme_mode(self):
        selected = self.theme_picker.SelectedItem
        return resource_loader.theme_descriptor_mode(selected, self._theme_mode)

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
        label = resource_loader.theme_descriptor_label(
            self._theme_mode,
            self._theme_mode,
        )
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
