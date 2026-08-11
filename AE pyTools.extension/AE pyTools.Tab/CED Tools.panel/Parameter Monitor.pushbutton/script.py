# -*- coding: utf-8 -*-
"""Parameter Monitor modeless Phase 1 entry point."""

from __future__ import print_function

import os
import traceback

import clr

for _assembly in ("PresentationFramework", "PresentationCore", "WindowsBase", "System.Data"):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass

from System import Boolean, DateTime, Object, String
from System.Collections.Generic import List
from System.Data import DataTable
from System.Windows import Application, GridLength, GridUnitType, Thickness, VerticalAlignment, Visibility, WindowState
from System.Windows.Controls import Dock
from System.Windows.Controls import Button, CheckBox, ContextMenu, DockPanel, MenuItem, Separator, TextBlock
from System.Windows.Input import Key, Keyboard, ModifierKeys, MouseButtonState
from System.Windows.Media import VisualTreeHelper
from pyrevit import forms, revit, script

from UIClasses import pathing as ui_pathing


TITLE = "Parameter Monitor"
WINDOW_MARKER = "_ced_parameter_monitor_modeless_v1"
THIS_DIR = os.path.abspath(os.path.dirname(__file__))
LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate CEDLib.lib for Parameter Monitor.", title=TITLE)
    raise SystemExit

from UIClasses import load_theme_state_from_config
from UIClasses.ui_bases import CEDWindowBase

import external_events
import models
import storage_service
import viewmodel


LOGGER = script.get_logger()
UI_RESOURCES_ROOT = (
    ui_pathing.resolve_ui_resources_root(LIB_ROOT)
    or os.path.abspath(os.path.join(LIB_ROOT, "UIClasses", "Resources"))
)


def _find_existing_window():
    try:
        application = Application.Current
        if application is None:
            return None
        for window in application.Windows:
            try:
                if str(getattr(window, "Tag", "") or "") == WINDOW_MARKER:
                    return window
            except Exception:
                continue
    except Exception:
        pass
    return None


def _activate_existing(window):
    try:
        if window.WindowState == WindowState.Minimized:
            window.WindowState = WindowState.Normal
        window.Show()
        window.Activate()
        window.Topmost = True
        window.Topmost = False
        return True
    except Exception:
        return False


class ParameterMonitorWindow(CEDWindowBase):
    theme_aware = True
    use_config_theme = True

    def __init__(self, store, gateway, theme_mode, accent_mode):
        # Keep strong references for the lifetime of this modeless window. pyRevit
        # can release command-module globals after the launch command returns.
        self._viewmodel = viewmodel
        self._models = models
        self._forms = forms
        self._logger = LOGGER
        self._traceback = traceback
        self._title = TITLE
        self._object_list_type = List[Object]
        self._data_table_type = DataTable
        self._string_type = String
        self._boolean_type = Boolean
        self._date_time_type = DateTime
        self._visibility = Visibility
        self._grid_length_type = GridLength
        self._grid_unit_type = GridUnitType
        self._system_thickness = Thickness
        self._dock_right = Dock.Right
        self._vertical_alignment = VerticalAlignment
        self._key = Key
        self._keyboard = Keyboard
        self._modifier_keys = ModifierKeys
        self._mouse_button_state = MouseButtonState
        self._visual_tree = VisualTreeHelper
        self._button_type = Button
        self._check_box_type = CheckBox
        self._context_menu_type = ContextMenu
        self._dock_panel_type = DockPanel
        self._menu_item_type = MenuItem
        self._separator_type = Separator
        self._text_block_type = TextBlock
        self._store = store
        self._gateway = gateway
        self._refreshing = False
        self._selected_set_id = None
        self._selected_persistent_id = None
        self._selected_persistent_ids = []
        self._all_element_rows = []
        self._base_element_rows = []
        self._column_filters = {}
        self._column_filter_buttons = {}
        self._element_column_definitions = [
            (1, "Status", "status"),
            (2, "Element", "element"),
            (3, "Family", "family"),
            (4, "Type", "type"),
            (5, "ID", "element_id"),
            (6, "Level", "level"),
            (7, "Parameter Changes", "parameter_change_text"),
            (8, "Missing", "missing_count"),
            (9, "Location Tracking", "location_text"),
            (10, "Location Change", "location_change_text"),
            (11, "Circuit / Device", "circuit"),
        ]
        self._tracking_sets_expanded_width = 290.0
        self._console_expanded_height = 130.0
        xaml_path = os.path.join(THIS_DIR, "ParameterMonitorWindow.xaml")
        CEDWindowBase.__init__(
            self,
            xaml_source=xaml_path,
            resources_root=UI_RESOURCES_ROOT,
            theme_mode=theme_mode,
            accent_mode=accent_mode,
            theme_aware=True,
            use_config_theme=True,
        )
        try:
            self.Tag = WINDOW_MARKER
        except Exception:
            pass
        self._tracking_rows = self._object_list_type()
        self._element_rows = self._object_list_type()
        self._property_rows = self._object_list_type()
        self._property_table = None
        self._console_lines = []
        self.TrackingSetList.ItemsSource = self._tracking_rows
        self.ElementGrid.ItemsSource = self._element_rows
        self.PropertyGrid.ItemsSource = self._property_rows
        self.FilterCombo.ItemsSource = list(self._viewmodel.FILTER_OPTIONS)
        self.FilterCombo.SelectedIndex = 0
        self._configure_element_grid_headers()
        self._log_console("INFO", "Parameter Monitor window opened.")
        self._apply_store(store)

    def _selected_set_row(self):
        try:
            return self.TrackingSetList.SelectedItem
        except Exception:
            return None

    def _selected_element_row(self):
        rows = self._selected_element_rows()
        return rows[0] if len(rows) == 1 else None

    def _selected_element_rows(self):
        try:
            rows = list(self.ElementGrid.SelectedItems)
            if rows:
                return rows
            selected = self.ElementGrid.SelectedItem
            return [selected] if selected is not None else []
        except Exception:
            return []

    def _sync_selection_from_grid(self):
        rows = []
        try:
            rows = list(self.ElementGrid.SelectedItems)
        except Exception:
            pass
        self._selected_persistent_ids = [row.persistent_id for row in rows]
        self._selected_persistent_id = rows[0].persistent_id if len(rows) == 1 else None
        return rows

    def _log_ui_exception(self, message):
        details = self._traceback.format_exc()
        self._log_console("ERROR", message, details)
        try:
            self._logger.exception(message)
        except Exception:
            pass

    def _configure_element_grid_headers(self):
        """Give every data column a sort surface and an Excel-style filter menu."""
        if len(self.ElementGrid.Columns):
            select_all = self._check_box_type()
            select_all.ToolTip = "Select or clear all displayed rows"
            select_all.Click += self.select_all_elements_clicked
            self.ElementGrid.Columns[0].Header = select_all
        for index, label, field in self._element_column_definitions:
            if index >= len(self.ElementGrid.Columns):
                continue
            panel = self._dock_panel_type()
            panel.LastChildFill = True
            filter_button = self._button_type()
            filter_button.Content = "▼"
            filter_button.Tag = field
            filter_button.Padding = self._system_thickness(3.0, 0.0, 3.0, 0.0)
            filter_button.Margin = self._system_thickness(5.0, 0.0, 0.0, 0.0)
            filter_button.ToolTip = "Filter {}".format(label)
            filter_button.Click += self.column_filter_clicked
            self._dock_panel_type.SetDock(filter_button, self._dock_right)
            header_text = self._text_block_type()
            header_text.Text = label
            header_text.VerticalAlignment = self._vertical_alignment.Center
            panel.Children.Add(filter_button)
            panel.Children.Add(header_text)
            column = self.ElementGrid.Columns[index]
            column.Header = panel
            column.SortMemberPath = field
            self._column_filter_buttons[field] = filter_button

    @staticmethod
    def _filter_value(row, field):
        value = getattr(row, str(field or ""), "")
        return str(value if value is not None else "")

    def _update_filter_button(self, field):
        button = self._column_filter_buttons.get(field)
        if button is None:
            return
        filter_spec = self._column_filters.get(field)
        button.Content = "●" if filter_spec else "▼"
        button.ToolTip = (
            "{} filter: {}".format(filter_spec.get("mode"), filter_spec.get("value"))
            if filter_spec else "Filter {}".format(field.replace("_", " ").title())
        )

    def column_filter_clicked(self, sender, args):
        try:
            field = str(getattr(sender, "Tag", "") or "")
            values = sorted(set([
                self._filter_value(row, field) for row in self._base_element_rows
            ]), key=lambda item: item.lower())
            menu = self._context_menu_type()
            clear_item = self._menu_item_type()
            clear_item.Header = "All values"
            clear_item.IsCheckable = True
            clear_item.IsChecked = field not in self._column_filters
            clear_item.Tag = "{}\x1fclear\x1f".format(field)
            clear_item.Click += self.column_filter_item_clicked
            menu.Items.Add(clear_item)
            contains_item = self._menu_item_type()
            contains_item.Header = "Contains..."
            contains_item.Tag = "{}\x1fcontains\x1f".format(field)
            contains_item.Click += self.column_filter_item_clicked
            menu.Items.Add(contains_item)
            menu.Items.Add(self._separator_type())
            for value in values[:100]:
                item = self._menu_item_type()
                item.Header = value or "(blank)"
                item.IsCheckable = True
                current = self._column_filters.get(field) or {}
                item.IsChecked = current.get("mode") == "equals" and current.get("value") == value
                item.Tag = "{}\x1fequals\x1f{}".format(field, value)
                item.Click += self.column_filter_item_clicked
                menu.Items.Add(item)
            if len(values) > 100:
                overflow = self._menu_item_type()
                overflow.Header = "{} values; use Contains...".format(len(values))
                overflow.IsEnabled = False
                menu.Items.Add(overflow)
            menu.PlacementTarget = sender
            menu.IsOpen = True
            args.Handled = True
        except Exception:
            self._log_ui_exception("Parameter Monitor column filter menu failed")

    def column_filter_item_clicked(self, sender, args):
        try:
            parts = str(getattr(sender, "Tag", "") or "").split("\x1f", 2)
            field, mode = parts[0], parts[1]
            value = parts[2] if len(parts) > 2 else ""
            if mode == "clear":
                self._column_filters.pop(field, None)
            elif mode == "contains":
                current = (self._column_filters.get(field) or {}).get("value", "")
                value = self._forms.ask_for_string(
                    default=current,
                    prompt="Show rows where {} contains:".format(field.replace("_", " ")),
                    title="Filter {}".format(field.replace("_", " ").title()),
                )
                if value is None:
                    return
                if str(value) == "":
                    self._column_filters.pop(field, None)
                else:
                    self._column_filters[field] = {"mode": "contains", "value": str(value)}
            else:
                self._column_filters[field] = {"mode": "equals", "value": value}
            self._update_filter_button(field)
            self._selected_persistent_id = None
            self._selected_persistent_ids = []
            self._refresh_elements()
            self._log_console("FILTER", "Column filters updated: {}".format(self._column_filters))
        except Exception:
            self._log_ui_exception("Parameter Monitor could not apply a column filter")

    def columns_menu_clicked(self, sender, args):
        try:
            menu = self._context_menu_type()
            for index, label, _field in self._element_column_definitions:
                if index >= len(self.ElementGrid.Columns):
                    continue
                item = self._menu_item_type()
                item.Header = label
                item.IsCheckable = True
                item.IsChecked = self.ElementGrid.Columns[index].Visibility == self._visibility.Visible
                item.Tag = str(index)
                item.Click += self.column_visibility_clicked
                menu.Items.Add(item)
            menu.PlacementTarget = sender
            menu.IsOpen = True
        except Exception:
            self._log_ui_exception("Parameter Monitor column menu failed")

    def column_visibility_clicked(self, sender, args):
        try:
            index = int(str(sender.Tag))
            column = self.ElementGrid.Columns[index]
            column.Visibility = self._visibility.Visible if bool(sender.IsChecked) else self._visibility.Collapsed
            self._log_console("UI", "Column {} visibility={}.".format(index, bool(sender.IsChecked)))
        except Exception:
            self._log_ui_exception("Parameter Monitor could not change column visibility")

    def _selected_property_row(self):
        try:
            return self.PropertyGrid.SelectedItem
        except Exception:
            return None

    @staticmethod
    def _property_field(row, name, default=None):
        if row is None:
            return default
        try:
            value = row.Row[str(name)]
            return default if value is None else value
        except Exception:
            pass
        try:
            value = row[str(name)]
            return default if value is None else value
        except Exception:
            pass
        return getattr(row, name, default)

    def _log_console(self, level, message, details=None):
        """Append a timestamped, persistent diagnostic entry to the UI console."""
        try:
            stamp = self._date_time_type.Now.ToString("HH:mm:ss.fff")
            lines = ["[{}] {:<5} {}".format(stamp, str(level or "INFO").upper(), str(message or ""))]
            if details:
                lines.extend(["    {}".format(line) for line in str(details).splitlines()])
            self._console_lines.extend(lines)
            if len(self._console_lines) > 1500:
                self._console_lines = self._console_lines[-1200:]
                self.OutputConsoleTextBox.Text = "\n".join(self._console_lines) + "\n"
            else:
                self.OutputConsoleTextBox.AppendText("\n".join(lines) + "\n")
            self.OutputConsoleTextBox.ScrollToEnd()
        except Exception:
            pass

    def _set_status(self, message):
        try:
            self.StatusBarText.Text = str(message or "")
        except Exception:
            pass

    def _set_busy(self, busy, message=None):
        try:
            self.MainContent.IsEnabled = not bool(busy)
        except Exception:
            pass
        if message:
            self._set_status(message)

    def _run(self, operation, payload=None):
        if self._gateway.is_busy():
            self._set_status("Another Parameter Monitor operation is already running.")
            return
        action_payload = payload or {}
        self._log_console("QUEUE", "{} | payload={}".format(operation, action_payload))
        self._set_busy(True, "Working: {}...".format(str(operation).replace("_", " ")))
        try:
            raised = self._gateway.raise_action(
                operation,
                payload=action_payload,
                callback=self._external_complete,
            )
        except Exception as ex:
            self._set_busy(False, "Operation could not be queued.")
            self._log_console("ERROR", "Could not queue {}: {}".format(operation, ex), self._traceback.format_exc())
            self._forms.alert("Parameter Monitor could not queue the operation:\n\n{}".format(ex), title=self._title)
            return
        if not raised:
            self._set_busy(False, "Another operation is already pending.")

    def _external_complete(self, status, operation, result, error):
        """Apply an ExternalEvent result immediately on Revit's UI thread."""
        error_details = (
            getattr(error, "_parameter_monitor_traceback", None) or repr(error)
            if error is not None else ""
        )
        self._log_console(str(status).upper(), "{} completed with status {}.".format(operation, status), error_details)
        try:
            self._apply_external_complete(status, operation, result, error)
        except Exception as ex:
            details = self._traceback.format_exc()
            self._set_busy(False, "UI refresh failed: {}".format(ex))
            self._log_console("ERROR", "UI refresh failed after {}: {}".format(operation, ex), details)
            try:
                self._logger.exception("Parameter Monitor could not apply UI result")
            except Exception:
                pass
            self._forms.alert(
                "Parameter Monitor completed the model operation, but could not refresh "
                "the open window. Full details are in the Output Console.\n\n{}".format(ex),
                title=self._title,
            )

    def _apply_external_complete(self, status, operation, result, error):
        self._set_busy(False)
        if status == "error":
            self._set_status("{} failed.".format(str(operation).replace("_", " ").title()))
            self._forms.alert(
                "Parameter Monitor operation failed:\n\n{}".format(error),
                title=self._title,
            )
            return
        if status == "cancelled":
            self._set_status("Operation cancelled. No project monitor data changed.")
            return
        result = result or {}
        store = result.get("store")
        if store is not None:
            if operation == "add_set":
                added_sets = list(store.get("tracking_sets") or [])
                if added_sets:
                    self._selected_set_id = str(added_sets[-1].get("set_id") or "")
                    self._selected_persistent_id = None
                    self._selected_persistent_ids = []
            if operation == "sync_element_linker" and result.get("sync_set_id"):
                self._selected_set_id = str(result.get("sync_set_id") or "")
                self._selected_persistent_id = None
                self._selected_persistent_ids = []
            self._apply_store(store)
        self._set_status(result.get("message") or "Operation complete.")

    def _replace_items(self, control, items):
        """Replace a grid source in one notification for large result sets."""
        source = self._object_list_type()
        for item in list(items or []):
            source.Add(item)
        control.ItemsSource = None
        control.ItemsSource = source
        return source

    def _replace_children_items(self, items):
        """Bind the LINKED CHILDREN list through CLR DataRowView objects."""
        table = self._data_table_type("ParameterMonitorLinkedChildren")
        for name in ("family", "type", "origin", "persistent_id", "state"):
            table.Columns.Add(name, self._string_type)
        for item in list(items or []):
            row = table.NewRow()
            for name in ("family", "type", "origin", "persistent_id", "state"):
                row[name] = str(getattr(item, name, "") or "")
            table.Rows.Add(row)
        self.ChildrenGrid.ItemsSource = None
        self.ChildrenGrid.ItemsSource = table.DefaultView
        self._children_table = table
        return table.DefaultView

    def _replace_property_items(self, items):
        """Bind inspector values through CLR DataRowView objects, not Python wrappers."""
        table = self._data_table_type("ParameterMonitorProperties")
        for name in ("name", "scope", "accepted", "current", "state", "key"):
            table.Columns.Add(name, self._string_type)
        table.Columns.Add("changed", self._boolean_type)
        table.Columns.Add("can_resolve", self._boolean_type)
        for item in list(items or []):
            row = table.NewRow()
            for name in ("name", "scope", "accepted", "current", "state", "key"):
                row[name] = str(getattr(item, name, "") or "")
            row["changed"] = bool(getattr(item, "changed", False))
            row["can_resolve"] = bool(getattr(item, "can_resolve", False))
            table.Rows.Add(row)
        self.PropertyGrid.ItemsSource = None
        self.PropertyGrid.ItemsSource = table.DefaultView
        self._property_table = table
        return table.DefaultView

    def _apply_store(self, store):
        self._refreshing = True
        try:
            self._store = store or self._models.new_project_store()
            preferred_set_id = self._selected_set_id
            rows = self._viewmodel.tracking_set_rows(self._store)
            self._log_console("UI", "Applying project store: {} tracking set(s).".format(len(rows)))
            self._tracking_rows = self._replace_items(self.TrackingSetList, rows)
            selected = None
            matched_preferred = False
            for row in rows:
                if row.set_id == preferred_set_id:
                    selected = row
                    matched_preferred = True
                    break
            if selected is None and rows:
                selected = rows[0]
            if not matched_preferred:
                self._selected_persistent_id = None
                self._selected_persistent_ids = []
            self.TrackingSetList.SelectedItem = selected
            self._selected_set_id = selected.set_id if selected is not None else None
        finally:
            self._refreshing = False
        self._refresh_selected_set()

    def _refresh_selected_set(self):
        row = self._selected_set_row()
        if row is None:
            self.SetTitleText.Text = "No Tracking Sets"
            self.SetSubtitleText.Text = "Use Add Set to create an accepted baseline."
            self.SetStatusText.Text = "-"
            self.SetLastCheckText.Text = "Last check: Never"
            for control in (
                self.ChangedCountText, self.AddedCountText,
                self.RemovedCountText, self.UnchangedCountText,
            ):
                control.Text = "0"
            self._element_rows = self._replace_items(self.ElementGrid, [])
            self.ElementCountText.Text = "0 elements"
            self._selected_persistent_id = None
            self._selected_persistent_ids = []
            self._refresh_inspector()
            return
        self._selected_set_id = row.set_id
        tracking_set = row.data
        self.SetTitleText.Text = row.name
        subtitle = "{} | {} | {} tracked properties".format(
            row.source, row.category, row.property_count
        )
        detail = row.source_condition_text or row.status_message
        self.SetSubtitleText.Text = "{} | {}".format(subtitle, detail) if detail else subtitle
        self.SetStatusText.Text = "{} | {}".format(row.status_text, row.active_text)
        self.SetLastCheckText.Text = "Last check: {}".format(row.last_check)
        self.ChangedCountText.Text = str(row.summary.get("changed", 0))
        self.AddedCountText.Text = str(row.summary.get("added", 0))
        self.RemovedCountText.Text = str(row.summary.get("removed", 0))
        self.UnchangedCountText.Text = str(row.summary.get("unchanged", 0))
        self._refresh_elements(tracking_set)

    def _current_filter_key(self):
        try:
            selected = self.FilterCombo.SelectedItem
            return selected.key if selected is not None else self._viewmodel.FILTER_ALL
        except Exception:
            return self._viewmodel.FILTER_ALL

    def _refresh_elements(self, tracking_set=None):
        set_row = self._selected_set_row()
        tracking_set = tracking_set or (set_row.data if set_row is not None else None)
        search = ""
        try:
            search = self.SearchTextBox.Text
        except Exception:
            pass
        base_rows = self._viewmodel.element_rows(
            tracking_set,
            filter_key=self._current_filter_key(),
            search_text=search,
        )
        self._base_element_rows = list(base_rows)
        rows = []
        for row in base_rows:
            include = True
            for field, filter_spec in list(self._column_filters.items()):
                actual = self._filter_value(row, field)
                expected = str(filter_spec.get("value") or "")
                if filter_spec.get("mode") == "contains":
                    if expected.lower() not in actual.lower():
                        include = False
                        break
                elif actual.lower() != expected.lower():
                    include = False
                    break
            if include:
                rows.append(row)
        self._all_element_rows = list(rows)
        preferred_ids = list(self._selected_persistent_ids or [])
        if not preferred_ids and self._selected_persistent_id:
            preferred_ids = [self._selected_persistent_id]
        preferred_id_set = set(preferred_ids)
        self._refreshing = True
        try:
            self._element_rows = self._replace_items(self.ElementGrid, rows)
            selected_rows = []
            for row in rows:
                if row.persistent_id in preferred_id_set:
                    selected_rows.append(row)
            if not selected_rows and rows:
                selected_rows = [rows[0]]
            self.ElementGrid.SelectedItems.Clear()
            self.ElementGrid.SelectedItem = selected_rows[0] if selected_rows else None
            for row in selected_rows[1:]:
                self.ElementGrid.SelectedItems.Add(row)
            self._selected_persistent_ids = [row.persistent_id for row in selected_rows]
            self._selected_persistent_id = (
                selected_rows[0].persistent_id if len(selected_rows) == 1 else None
            )
            self.ElementCountText.Text = (
                "{} of {} element(s) shown".format(len(rows), len(base_rows))
                if self._column_filters else "{} element(s) shown".format(len(rows))
            )
        finally:
            self._refreshing = False
        self._refresh_inspector()
        self._log_console("UI", "Element grid refreshed: {} row(s), selected={}.".format(
            len(rows), self._selected_persistent_ids
        ))

    def _refresh_inspector(self):
        set_row = self._selected_set_row()
        element_rows = self._selected_element_rows()
        tracking_set = set_row.data if set_row is not None else None
        if not element_rows:
            self._log_console("SELECT", "No element selected; inspector cleared.")
            self.ElementTitleText.Text = "Select an element"
            self.ElementContextText.Text = "-"
            self._property_rows = self._replace_property_items([])
            self.RelationshipText.Text = "No host device linked"
            self.ChildrenSummaryText.Text = "No linked children"
            self._replace_children_items([])
            self._set_action_state([], None)
            return
        if len(element_rows) > 1:
            self._selected_persistent_id = None
            self._selected_persistent_ids = [row.persistent_id for row in element_rows]
            self.ElementTitleText.Text = "{} elements selected".format(len(element_rows))
            self.ElementContextText.Text = (
                "Element-specific properties and device details are hidden for multiple selection."
            )
            self._property_rows = self._replace_property_items([])
            self.RelationshipText.Text = "Multiple elements selected"
            self.ChildrenSummaryText.Text = "Multiple elements selected"
            self._replace_children_items([])
            self._set_action_state(element_rows, None)
            self._log_console(
                "SELECT",
                "Multiple selection: {} element(s), persistent_ids={}.".format(
                    len(element_rows), self._selected_persistent_ids
                ),
            )
            return
        element_row = element_rows[0]
        self._selected_persistent_id = element_row.persistent_id
        self._selected_persistent_ids = [element_row.persistent_id]
        self.ElementTitleText.Text = element_row.element
        self.ElementContextText.Text = "{} | ID {} | Level {} | {}".format(
            element_row.family_type,
            element_row.element_id,
            element_row.level,
            element_row.status,
        )
        rows = self._viewmodel.property_rows(tracking_set, element_row)
        selected_property = self._selected_property_row()
        preferred_property_key = self._property_field(selected_property, "key")
        self._refreshing = True
        try:
            self._property_rows = self._replace_property_items(rows)
            selected_property = None
            for row in self._property_rows:
                if self._property_field(row, "key") == preferred_property_key:
                    selected_property = row
                    break
            if selected_property is None and len(self._property_rows):
                selected_property = self._property_rows[0]
            self.PropertyGrid.SelectedItem = selected_property
        finally:
            self._refreshing = False
        context = (element_row.record or {}).get("relationship_context") or {}
        relationship = (element_row.record or {}).get("relationship") or {}
        if relationship:
            self.RelationshipText.Text = "{} | {}".format(
                context.get("device_name") or relationship.get("device_name") or "Linked Device",
                context.get("status_text") or "Relationship retained",
            )
        else:
            self.RelationshipText.Text = "No host device linked"
        children_info = self._viewmodel.linked_children_info(
            tracking_set, element_row.record
        )
        self._replace_children_items(children_info.get("children") or [])
        if not children_info.get("count"):
            self.ChildrenSummaryText.Text = "No linked children"
        elif children_info.get("parent_moved"):
            self.ChildrenSummaryText.Text = (
                "{} linked child(ren) | Parent MOVED - children can follow".format(
                    children_info.get("count")
                )
            )
        else:
            self.ChildrenSummaryText.Text = "{} linked child(ren) | In sync".format(
                children_info.get("count")
            )
        self._set_action_state([element_row], self._selected_property_row())
        value_summary = "; ".join([
            "{}: {} -> {} ({})".format(row.name, row.accepted, row.current, row.state)
            for row in rows
        ])
        self._log_console(
            "SELECT",
            "Element {} | id={} | persistent_id={} | {} property row(s).".format(
                element_row.element, element_row.element_id, element_row.persistent_id, len(rows)
            ),
            value_summary,
        )

    def _set_action_state(self, element_rows, property_row):
        element_rows = list(element_rows or [])
        count = len(element_rows)
        single_row = element_rows[0] if count == 1 else None
        navigable_rows = [row for row in element_rows if bool(row.can_navigate)]
        untrackable_rows = [row for row in element_rows if bool(row.can_untrack)]
        can_navigate = bool(navigable_rows)
        can_untrack = bool(untrackable_rows)
        has_relationship = bool(
            single_row is not None
            and single_row.has_relationship
            and single_row.can_navigate
        )
        self.ShowElementButton.IsEnabled = can_navigate
        self.SelectElementButton.IsEnabled = can_navigate
        self.ShowElementButton.Content = (
            "Show {} in Model".format(len(navigable_rows)) if count > 1 else "Show in Model"
        )
        self.SelectElementButton.Content = (
            "Select {} in Model".format(len(navigable_rows)) if count > 1 else "Select in Model"
        )
        self.ToggleLocationButton.IsEnabled = can_untrack
        all_location_on = bool(untrackable_rows) and all([
            bool((row.record or {}).get("track_location", False))
            for row in untrackable_rows
        ])
        location_verb = "Disable" if all_location_on else "Enable"
        self.ToggleLocationButton.Content = "{} Location ({})".format(
            location_verb,
            len(untrackable_rows),
        ) if count > 1 else "{} Location".format(location_verb)
        self.ResolveElementButton.IsEnabled = bool(
            single_row is not None and single_row.can_resolve
        )
        self.UntrackButton.Visibility = self._visibility.Visible if can_untrack else self._visibility.Collapsed
        self.UntrackButton.Content = (
            "Untrack {} Elements".format(len(untrackable_rows))
            if count > 1 else "Untrack"
        )
        self.RestoreButton.Visibility = self._visibility.Visible if single_row is not None and single_row.can_restore else self._visibility.Collapsed
        self.RemoveRecordButton.Visibility = self._visibility.Visible if single_row is not None and single_row.can_remove_record else self._visibility.Collapsed
        self.ResolvePropertyButton.IsEnabled = bool(
            single_row is not None and property_row is not None
            and bool(self._property_field(property_row, "can_resolve", False))
            and single_row.can_navigate
        )
        self.LinkDeviceButton.IsEnabled = bool(single_row is not None and single_row.can_navigate)
        self.UnlinkDeviceButton.IsEnabled = has_relationship
        self.SelectDeviceButton.IsEnabled = has_relationship
        self.ShowDeviceButton.IsEnabled = has_relationship
        selected_record = (single_row.record or {}) if single_row is not None else {}
        circuits = list(((selected_record.get("relationship_context") or {}).get("circuits") or []))
        self.SelectCircuitButton.IsEnabled = has_relationship and bool(circuits)
        set_row = self._selected_set_row()
        tracking_set = set_row.data if set_row is not None else None
        children_infos = [
            self._viewmodel.linked_children_info(tracking_set, row.record)
            for row in element_rows
            if row.record is not None
        ]
        has_any_children = any([info.get("count") for info in children_infos])
        movable_count = sum([
            len(info.get("movable_child_ids") or []) for info in children_infos
        ])
        self.MoveWithParentButton.Visibility = (
            self._visibility.Visible if has_any_children else self._visibility.Collapsed
        )
        self.MoveWithParentButton.IsEnabled = movable_count > 0
        self.MoveWithParentButton.Content = (
            "Move {} Children with Parent".format(movable_count)
            if movable_count
            else "Move Children with Parent"
        )

    def _payload_for_selection(self, require_element=False, allow_multiple=False):
        set_row = self._selected_set_row()
        if set_row is None:
            self._forms.alert("Select a Tracking Set first.", title=self._title)
            return None
        payload = {"set_id": set_row.set_id}
        if require_element:
            element_rows = self._selected_element_rows()
            if not element_rows:
                self._forms.alert("Select one or more elements first.", title=self._title)
                return None
            if not allow_multiple and len(element_rows) != 1:
                self._forms.alert("Select exactly one element for this action.", title=self._title)
                return None
            payload["persistent_ids"] = [row.persistent_id for row in element_rows]
            if len(element_rows) == 1:
                payload["persistent_id"] = element_rows[0].persistent_id
        return payload

    def tracking_set_selection_changed(self, sender, args):
        if self._refreshing:
            return
        row = self._selected_set_row()
        self._selected_set_id = row.set_id if row is not None else None
        self._selected_persistent_id = None
        self._selected_persistent_ids = []
        self._refresh_selected_set()

    def element_selection_changed(self, sender, args):
        if self._refreshing:
            return
        rows = self._sync_selection_from_grid()
        self._log_console("SELECT", "Middle grid selection changed: {}".format(self._selected_persistent_ids))
        self._refresh_inspector()

    def element_checkbox_preview_mouse_down(self, sender, args):
        """Toggle one row without clearing other selected rows."""
        try:
            row = getattr(sender, "DataContext", None)
            if row is None or not hasattr(row, "persistent_id"):
                return
            selected = row in list(self.ElementGrid.SelectedItems)
            self._refreshing = True
            try:
                if selected:
                    self.ElementGrid.SelectedItems.Remove(row)
                else:
                    self.ElementGrid.SelectedItems.Add(row)
            finally:
                self._refreshing = False
            self._sync_selection_from_grid()
            self._log_console("SELECT", "Checkbox selection changed: {}".format(self._selected_persistent_ids))
            self._refresh_inspector()
            args.Handled = True
        except Exception:
            self._log_ui_exception("Parameter Monitor checkbox selection failed")

    def element_grid_preview_mouse_down(self, sender, args):
        """Intercept checkbox clicks before DataGrid row-selection logic clears peers."""
        try:
            current = getattr(args, "OriginalSource", None)
            for _index in range(8):
                if current is None:
                    return
                if isinstance(current, self._check_box_type):
                    self.element_checkbox_preview_mouse_down(current, args)
                    return
                try:
                    current = self._visual_tree.GetParent(current)
                except Exception:
                    return
        except Exception:
            self._log_ui_exception("Parameter Monitor checkbox click routing failed")

    def select_all_elements_clicked(self, sender, args):
        try:
            select_all = bool(getattr(sender, "IsChecked", False))
            self._refreshing = True
            try:
                self.ElementGrid.SelectedItems.Clear()
                if select_all:
                    for row in self._element_rows:
                        self.ElementGrid.SelectedItems.Add(row)
            finally:
                self._refreshing = False
            self._sync_selection_from_grid()
            self._log_console(
                "SELECT",
                "{} all displayed rows ({} selected).".format(
                    "Selected" if select_all else "Cleared", len(self._selected_persistent_ids)
                ),
            )
            self._refresh_inspector()
        except Exception:
            self._log_ui_exception("Parameter Monitor select-all checkbox failed")

    def element_grid_preview_mouse_move(self, sender, args):
        """Prevent click-drag range selection while leaving scrollbars usable."""
        try:
            if args.LeftButton != self._mouse_button_state.Pressed:
                return
            current = getattr(args, "OriginalSource", None)
            for _index in range(8):
                if current is None:
                    break
                type_name = str(current.GetType().Name or "")
                if type_name in ("ScrollBar", "Thumb", "RepeatButton"):
                    return
                try:
                    current = self._visual_tree.GetParent(current)
                except Exception:
                    break
            args.Handled = True
        except Exception:
            pass

    def collapse_tracking_sets_clicked(self, sender, args):
        try:
            width = float(getattr(self.TrackingSetsColumn, "ActualWidth", 0.0) or 0.0)
            if width > 100.0:
                self._tracking_sets_expanded_width = width
            self.TrackingSetsPanel.Visibility = self._visibility.Collapsed
            self.TrackingSetsSplitter.Visibility = self._visibility.Collapsed
            self.TrackingSetsCollapsedStrip.Visibility = self._visibility.Visible
            self.TrackingSetsColumn.Width = self._grid_length_type(38.0)
            self.TrackingSetsSplitterColumn.Width = self._grid_length_type(0.0)
            self._log_console("UI", "Tracking Sets panel collapsed.")
        except Exception:
            self._log_ui_exception("Parameter Monitor could not collapse Tracking Sets")

    def expand_tracking_sets_clicked(self, sender, args):
        try:
            self.TrackingSetsColumn.Width = self._grid_length_type(
                max(250.0, float(self._tracking_sets_expanded_width or 290.0))
            )
            self.TrackingSetsSplitterColumn.Width = self._grid_length_type(6.0)
            self.TrackingSetsPanel.Visibility = self._visibility.Visible
            self.TrackingSetsSplitter.Visibility = self._visibility.Visible
            self.TrackingSetsCollapsedStrip.Visibility = self._visibility.Collapsed
            self._log_console("UI", "Tracking Sets panel expanded.")
        except Exception:
            self._log_ui_exception("Parameter Monitor could not expand Tracking Sets")

    def parameters_expander_expanded(self, sender, args):
        # Give the parameters section its stretching row back.
        try:
            self.ParametersRowDefinition.Height = self._grid_length_type(
                1.0, self._grid_unit_type.Star
            )
        except Exception:
            pass

    def parameters_expander_collapsed(self, sender, args):
        # Collapse the stretching row so lower sections reclaim the space.
        try:
            self.ParametersRowDefinition.Height = self._grid_length_type.Auto
        except Exception:
            pass

    def console_expander_expanded(self, sender, args):
        try:
            self.ConsoleRowDefinition.MinHeight = 80.0
            self.ConsoleRowDefinition.Height = self._grid_length_type(
                max(80.0, float(getattr(self, "_console_expanded_height", 130.0) or 130.0))
            )
        except Exception:
            pass

    def console_expander_collapsed(self, sender, args):
        try:
            height = float(getattr(self.ConsoleRowDefinition, "ActualHeight", 0.0) or 0.0)
            if height > 80.0:
                self._console_expanded_height = height
            self.ConsoleRowDefinition.MinHeight = 0.0
            self.ConsoleRowDefinition.Height = self._grid_length_type.Auto
        except Exception:
            pass

    def window_preview_key_down(self, sender, args):
        if args.Key == self._key.Escape:
            args.Handled = True
            self._log_console(
                "INFO",
                "Escape pressed. The Parameter Monitor window remains open; active Revit picks may cancel.",
            )

    def property_selection_changed(self, sender, args):
        if self._refreshing:
            return
        self._set_action_state(self._selected_element_rows(), self._selected_property_row())

    def filter_selection_changed(self, sender, args):
        if self._refreshing or not hasattr(self, "ElementGrid"):
            return
        self._selected_persistent_id = None
        self._refresh_elements()

    def search_text_changed(self, sender, args):
        if self._refreshing or not hasattr(self, "ElementGrid"):
            return
        self._refresh_elements()

    def add_set_clicked(self, sender, args):
        self._run("add_set")

    def edit_set_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("edit_set", payload)

    def delete_set_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("delete_set", payload)

    def toggle_active_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("toggle_active", payload)

    def check_set_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("check_set", payload)

    def check_all_clicked(self, sender, args):
        self._run("check_all")

    def refresh_store_clicked(self, sender, args):
        self._run("refresh_store")

    def export_sets_clicked(self, sender, args):
        payload = self._payload_for_selection() or {}
        self._run("export_definitions", payload)

    def import_sets_clicked(self, sender, args):
        self._run("import_definitions")

    def export_report_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("export_report", payload)

    def resolve_set_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("resolve_set", payload)

    def remove_all_removed_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("remove_all_removed", payload)

    def resolve_property_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        property_row = self._selected_property_row()
        if payload and property_row is not None:
            payload["property_key"] = self._property_field(property_row, "key")
            self._run("resolve_property", payload)

    def clear_console_clicked(self, sender, args):
        self._console_lines = []
        try:
            self.OutputConsoleTextBox.Clear()
        except Exception:
            pass

    def resolve_element_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("resolve_element", payload)

    def untrack_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True, allow_multiple=True)
        if payload:
            payload["persistent_ids"] = [
                row.persistent_id for row in self._selected_element_rows()
                if row.can_untrack
            ]
            self._run("untrack_elements", payload)

    def restore_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("restore_element", payload)

    def remove_record_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("remove_record", payload)

    def toggle_location_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True, allow_multiple=True)
        rows = self._selected_element_rows()
        trackable = [row for row in rows if row.can_untrack and row.record is not None]
        if payload and trackable:
            payload["persistent_ids"] = [row.persistent_id for row in trackable]
            payload["enabled"] = not all([
                bool(row.record.get("track_location", False)) for row in trackable
            ])
            self._run("toggle_location", payload)

    def location_all_on_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("location_all_on", payload)

    def location_all_off_clicked(self, sender, args):
        payload = self._payload_for_selection()
        if payload:
            self._run("location_all_off", payload)

    def show_element_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True, allow_multiple=True)
        if payload:
            payload["persistent_ids"] = [
                row.persistent_id for row in self._selected_element_rows()
                if row.can_navigate
            ]
            self._run("show_element", payload)

    def select_element_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True, allow_multiple=True)
        if payload:
            payload["persistent_ids"] = [
                row.persistent_id for row in self._selected_element_rows()
                if row.can_navigate
            ]
            self._run("select_element", payload)

    def link_device_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("pick_device", payload)

    def unlink_device_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("unlink_device", payload)

    def select_device_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("select_device", payload)

    def show_device_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("show_device", payload)

    def select_circuit_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True)
        if payload:
            self._run("select_circuit", payload)

    def sync_element_linker_clicked(self, sender, args):
        self._run("sync_element_linker")

    def move_with_parent_clicked(self, sender, args):
        payload = self._payload_for_selection(require_element=True, allow_multiple=True)
        if not payload:
            return
        set_row = self._selected_set_row()
        tracking_set = set_row.data if set_row is not None else None
        movable_ids = []
        for row in self._selected_element_rows():
            if row.record is None:
                continue
            info = self._viewmodel.linked_children_info(tracking_set, row.record)
            for child_id in info.get("movable_child_ids") or []:
                if child_id not in movable_ids:
                    movable_ids.append(child_id)
        if not movable_ids:
            self._forms.alert(
                "The selected parent has no pending move for its linked children. "
                "Scan the set first if the parent was just moved.",
                title=self._title,
            )
            return
        payload["persistent_ids"] = movable_ids
        self._run("move_with_parent", payload)


def main():
    existing = _find_existing_window()
    if existing is not None and _activate_existing(existing):
        return
    document = getattr(revit, "doc", None)
    if document is None:
        forms.alert("Open a Revit project before launching Parameter Monitor.", title=TITLE)
        return
    if bool(getattr(document, "IsFamilyDocument", False)):
        forms.alert("Parameter Monitor is available only in Revit project documents.", title=TITLE)
        return
    try:
        store = storage_service.load(document, logger=LOGGER)
    except Exception as ex:
        LOGGER.exception("Parameter Monitor could not load project storage")
        forms.alert(
            "Parameter Monitor could not load its project data. No data was overwritten.\n\n{}".format(ex),
            title=TITLE,
        )
        return
    theme_mode, accent_mode = load_theme_state_from_config(
        default_theme="light",
        default_accent="blue",
    )
    gateway = external_events.ParameterMonitorExternalEventGateway(document, logger=LOGGER)
    window = ParameterMonitorWindow(store, gateway, theme_mode, accent_mode)
    window.Show()


if __name__ == "__main__":
    main()
