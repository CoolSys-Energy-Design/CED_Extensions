# -*- coding: utf-8 -*-
"""Create Circuits by Device Parameter - WPF window controller (IronPython / .NET).

The DataGrid is grouped by the GroupVM object itself, so each group header can
host that circuit's Load Name (far left), Panel and Breaker controls. Member
rows are read-only info + an include checkbox + a status.
"""

import os

import clr

# ``ListSortDirection`` / ``SortDescription`` and the ``System.Windows.*``
# WPF types live in assemblies that pyRevit auto-references on .NET
# Framework (Revit <= 2024) but NOT on .NET 8 (Revit 2025 / 2026), where
# the import otherwise fails with "Cannot import name ListSortDirection".
# Reference every candidate assembly explicitly (idempotent, and any name
# that doesn't exist on a given runtime is ignored) BEFORE the imports
# below so they resolve on all three Revit versions.
for _asm in ("WindowsBase", "PresentationFramework", "PresentationCore",
             "System", "System.ObjectModel"):
    try:
        clr.AddReference(_asm)
    except (IOError, ImportError):
        # Some pyRevit/Revit combinations already load these assemblies or do
        # not expose an optional candidate by simple name. Other load errors
        # must remain visible in the pyRevit traceback.
        pass

from System import Action, Math
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import (
    INotifyPropertyChanged, PropertyChangedEventArgs,
)
try:
    from System.ComponentModel import ListSortDirection, SortDescription
except ImportError:
    # Last-resort fallback: the group sort below is skipped, but the
    # window still loads and groups correctly.
    ListSortDirection = None
    SortDescription = None
from System.Windows import DataObject, DragDrop, DragDropEffects, Visibility
from System.Windows.Controls import (
    Button, CheckBox, ComboBox, ContextMenu, DataGridRow, MenuItem, TextBox,
)
from System.Windows.Controls.Primitives import DataGridColumnHeader, ToggleButton
from System.Windows.Data import CollectionViewSource, PropertyGroupDescription
from System.Windows.Input import Key, Keyboard, ModifierKeys, MouseButtonState
from System.Windows.Media import BrushConverter, VisualTreeHelper
from System.Windows.Threading import DispatcherPriority

from pyrevit import forms, revit, DB

from Snippets import revit_helpers
from UIClasses import pathing as ui_pathing

import cg_core

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
_XAML = os.path.join(THIS_DIR, "CircuitGrouperWindow.xaml")
_BULK_XAML = os.path.join(THIS_DIR, "BulkCircuitValuesWindow.xaml")

LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate workspace root for Create Circuits by Device Parameter.",
                title="Create Circuits by Device Parameter")
    raise SystemExit

from UIClasses.ui_bases import CEDWindowBase

_CONV = BrushConverter()


def _brush(hexstr):
    return _CONV.ConvertFromString(hexstr)


BRUSH_READY = _brush("#5BD08A")
BRUSH_WARN = _brush("#E0B020")
BRUSH_BAD = _brush("#E0625A")
BRUSH_OFF = _brush("#7A7A7A")
BRUSH_INFO = _brush("#6AA9E0")

NONSTANDARD_TOOLTIP = ("Non-standard breaker size - this is not a standard frame "
                       "size. Verify it is intended before circuiting.")

# Large grouped views should not eagerly realize every member row. Users can
# expand groups explicitly, including through the header control.
INITIAL_EXPANDED_ROW_LIMIT = 250

class _Notifier(INotifyPropertyChanged):
    def __init__(self):
        self._handlers = []

    def add_PropertyChanged(self, handler):
        self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def notify(self, name):
        args = PropertyChangedEventArgs(name)
        for h in list(self._handlers):
            h(self, args)


# ---------------------------------------------------------------------------
# View-models
# ---------------------------------------------------------------------------
class GroupVM(_Notifier):
    def __init__(self, key, status_brushes=None):
        _Notifier.__init__(self)
        brushes = status_brushes or {}
        self._colors = brushes
        self.key = key
        self.load_name = key
        self.schedule_notes = ""
        self.panel = ""
        self.panel_options = List[str]()
        # set when the multi-panel auto-pick could not choose a panel
        # (no room / unknown names); shown in the validation pane and
        # cleared as soon as the user picks a panel manually
        self.panel_note = ""
        self.rating = cg_core.DEFAULT_RATING
        self.status = "Ready"
        self.status_color = brushes.get("ready", BRUSH_READY)
        self.rating_warning_visibility = Visibility.Collapsed
        self.rating_tooltip = NONSTANDARD_TOOLTIP
        self.status_detail = "Ready"
        self._rating_warning = False
        self.status_icon = u"⚠"
        self.status_icon_color = brushes.get("bad", BRUSH_BAD)
        self.is_editable = True
        self.planning_visibility = Visibility.Visible
        self.result_visibility = Visibility.Collapsed
        self.result_assignment_label = ""
        # collapse state, two-way bound to each circuit's Expander so the
        # user's expand/collapse choice survives view refreshes
        self.is_expanded = True
        self.is_selected = False
        self.is_drag_target = False

    def set_status(self, text, color):
        # skip the PropertyChanged round-trip when nothing changed; at
        # thousands of rows the no-op notifies are a real UI cost
        if self.status != text or self.status_color is not color:
            self.status = text
            self.status_color = color
            self.notify("status")
            self.notify("status_color")
        self._update_status_indicator()

    def _update_status_indicator(self):
        detail = (self.rating_tooltip
                  if self._rating_warning and self.status == "Ready"
                  else self.status)
        if self.status_detail != detail:
            self.status_detail = detail
            self.notify("status_detail")
        vis = (Visibility.Visible
               if self._rating_warning or self.status != "Ready"
               else Visibility.Collapsed)
        if self.rating_warning_visibility != vis:
            self.rating_warning_visibility = vis
            self.notify("rating_warning_visibility")
        icon_color = (
            self._colors.get("warn", BRUSH_WARN)
            if self._rating_warning and self.status == "Ready"
            else self._colors.get("bad", BRUSH_BAD))
        if self.status_icon != u"⚠":
            self.status_icon = u"⚠"
            self.notify("status_icon")
        if self.status_icon_color is not icon_color:
            self.status_icon_color = icon_color
            self.notify("status_icon_color")

    def set_result_status(self, kind, assignment_label, detail, colors):
        kind = str(kind or "error").lower()
        if kind == "success":
            icon = u"✓"
            color = colors.get("ready", BRUSH_READY)
        elif kind == "warning":
            icon = u"⚠"
            color = colors.get("warn", BRUSH_WARN)
        elif kind == "neutral":
            icon = u"—"
            color = colors.get("off", BRUSH_OFF)
        else:
            icon = u"✕"
            color = colors.get("bad", BRUSH_BAD)
        self.is_editable = False
        self.planning_visibility = Visibility.Collapsed
        self.result_visibility = Visibility.Visible
        self.result_assignment_label = str(assignment_label or "No circuit created")
        self.status = (
            "Circuited" if kind == "success" else
            "Created with attention" if kind == "warning" else
            "Not run" if kind == "neutral" else
            "Circuiting failed")
        self.status_detail = str(detail or self.status)
        self.status_icon = icon
        self.status_icon_color = color
        self.rating_warning_visibility = Visibility.Visible
        for property_name in (
                "is_editable", "planning_visibility", "result_visibility",
                "result_assignment_label", "status", "status_detail",
                "status_icon", "status_icon_color",
                "rating_warning_visibility"):
            self.notify(property_name)

    def set_expanded(self, value):
        value = bool(value)
        if self.is_expanded == value:
            return
        self.is_expanded = value
        self.notify("is_expanded")

    def set_selected(self, value):
        value = bool(value)
        if self.is_selected == value:
            return
        self.is_selected = value
        self.notify("is_selected")

    def set_drag_target(self, value):
        value = bool(value)
        if self.is_drag_target == value:
            return
        self.is_drag_target = value
        self.notify("is_drag_target")

    def set_warning(self, show):
        show = bool(show)
        if self._rating_warning == show:
            return
        self._rating_warning = show
        self._update_status_indicator()

    def set_panel_options(self, values):
        prior = [str(x) for x in list(self.panel_options or [])]
        incoming = [str(x) for x in list(values or [])]
        if prior == incoming:
            return
        options = List[str]()
        for value in incoming:
            options.Add(str(value))
        self.panel_options = options
        self.notify("panel_options")


class RowVM(_Notifier):
    def __init__(self, data, group=None, status_brushes=None):
        _Notifier.__init__(self)
        brushes = status_brushes or {}
        self.element_id = data["element_id"]
        self.family_type = data.get("family_type", "")
        self.identity_mark = data.get("identity_mark", "")
        self.voltage_text = data.get("voltage_text", "") or ""
        self.voltage_key = data.get("voltage_key", None)
        self.poles_text = data.get("poles_text", "") or ""
        self.poles_value = data.get("poles_value", None)
        self.space = data.get("space", "") or ""
        self.level = data.get("level", "") or ""
        self.elevation_text = data.get("elevation_text", "") or ""
        self.location = data.get("location", None)
        self.group_values = data.get("group_values", {}) or {}
        self.group_value = ""
        self.already_circuited = bool(data.get("already_circuited", False))
        self.src_panel = (data.get("panel", "") or "").strip()
        self.src_rating = (data.get("rating", "") or "").strip()
        self.src_schedule_notes = (data.get("schedule_notes", "") or "").strip()

        self.group = group
        self.group_key = group.key if group is not None else ""
        self.can_include = not self.already_circuited
        self.results_locked = False
        self.can_edit = bool(self.can_include)
        self.include = not self.already_circuited
        self.is_enabled = bool(self.can_include and self.include)
        self.status = ""
        self.status_color = brushes.get("ready", BRUSH_READY)

    def set_already_circuited(self, value, restore_include=False):
        value = bool(value)
        was_circuited = bool(self.already_circuited)
        if self.already_circuited != value:
            self.already_circuited = value
            self.notify("already_circuited")
        can_include = not value
        if self.can_include != can_include:
            self.can_include = can_include
            self.notify("can_include")
        desired_include = self.include
        if value:
            desired_include = False
        elif restore_include and was_circuited:
            desired_include = True
        if self.include != desired_include:
            self.include = desired_include
            self.notify("include")
        enabled = bool(self.can_include and self.include)
        if self.is_enabled != enabled:
            self.is_enabled = enabled
            self.notify("is_enabled")
        can_edit = bool(self.can_include and not self.results_locked)
        if self.can_edit != can_edit:
            self.can_edit = can_edit
            self.notify("can_edit")

    def lock_results(self):
        self.results_locked = True
        if self.can_edit:
            self.can_edit = False
            self.notify("can_edit")

    def refresh_from_data(self, data):
        """Refresh Revit-backed fields without changing this row's UI group."""
        field_map = (
            ("family_type", "family_type", ""),
            ("identity_mark", "identity_mark", ""),
            ("voltage_text", "voltage_text", ""),
            ("voltage_key", "voltage_key", None),
            ("poles_text", "poles_text", ""),
            ("poles_value", "poles_value", None),
            ("space", "space", ""),
            ("level", "level", ""),
            ("elevation_text", "elevation_text", ""),
            ("location", "location", None),
        )
        for attr_name, data_name, fallback in field_map:
            value = data.get(data_name, fallback)
            if fallback == "":
                value = value or ""
            if getattr(self, attr_name) != value:
                setattr(self, attr_name, value)
                self.notify(attr_name)
        self.group_values = data.get("group_values", {}) or {}
        self.src_panel = (data.get("panel", "") or "").strip()
        self.src_rating = (data.get("rating", "") or "").strip()
        self.src_schedule_notes = (
            data.get("schedule_notes", "") or "").strip()
        self.set_already_circuited(
            bool(data.get("already_circuited", False)),
            restore_include=True,
        )

    def set_status(self, text, color):
        if self.status == text and self.status_color is color:
            return
        self.status = text
        self.status_color = color
        self.notify("status")
        self.notify("status_color")

    def set_include(self, value):
        """Bulk checkbox changes must notify the row binding without forcing a
        full view refresh, which stalls at large row counts."""
        value = bool(value)
        if self.include == value:
            return
        self.include = value
        self.notify("include")
        enabled = bool(self.can_include and self.include)
        if self.is_enabled != enabled:
            self.is_enabled = enabled
            self.notify("is_enabled")

    def move_to(self, group):
        self.group = group
        self.group_key = group.key
        self.notify("group")
        self.notify("group_key")


class BulkCircuitValuesWindow(CEDWindowBase):
    """Small modal editor used to apply optional values to selected groups."""

    def __init__(self, panel_options, rating_options, owner=None):
        CEDWindowBase.__init__(self, xaml_source=_BULK_XAML, theme_aware=True)
        self.PanelOptions = List[str]()
        self.PanelOptions.Add("")
        for value in panel_options or []:
            value = str(value)
            if value and value not in list(self.PanelOptions):
                self.PanelOptions.Add(value)
        self.RatingOptions = List[str]()
        for value in rating_options or []:
            value = str(value)
            if value and value not in list(self.RatingOptions):
                self.RatingOptions.Add(value)
        self.DataContext = self
        self.BulkLoadNameTextBox = self.FindName("BulkLoadNameTextBox")
        self.BulkPanelCombo = self.FindName("BulkPanelCombo")
        self.BulkRatingCombo = self.FindName("BulkRatingCombo")
        self.BulkNotesTextBox = self.FindName("BulkNotesTextBox")
        self.BulkIncrementCheck = self.FindName("BulkIncrementCheck")
        self.result_values = None
        self.reset_clicked(None, None)
        if owner is not None:
            try:
                self.Owner = owner
            except Exception:
                pass

    @staticmethod
    def _text(control):
        try:
            return str(control.Text or "").strip()
        except Exception:
            return ""

    def reset_clicked(self, sender, args):
        self.BulkLoadNameTextBox.Text = ""
        self.BulkPanelCombo.SelectedIndex = -1
        self.BulkPanelCombo.Text = ""
        self.BulkRatingCombo.SelectedIndex = -1
        self.BulkRatingCombo.Text = ""
        self.BulkNotesTextBox.Text = ""
        self.BulkIncrementCheck.IsChecked = False

    def apply_clicked(self, sender, args):
        self.result_values = {
            "load_name": self._text(self.BulkLoadNameTextBox),
            "panel": self._text(self.BulkPanelCombo),
            "rating": self._text(self.BulkRatingCombo),
            "schedule_notes": self._text(self.BulkNotesTextBox),
            "increment_load_name": bool(self.BulkIncrementCheck.IsChecked),
        }
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.result_values = None
        self.DialogResult = False
        self.Close()


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
class CircuitGrouperWindow(CEDWindowBase):
    def __init__(self, rows_data, panel_options, name_to_id, rating_options,
                 group_param_options=None, default_group_param="", title="",
                 panel_info=None, name_param_options=None,
                 default_name_param="", gateway=None, document_key="",
                 document_title="", scope_label="model", scope_ids=None):
        # This is a modeless Revit tool. pyRevit's default Escape handler closes
        # WPF windows, which is especially easy to trigger after a model
        # selection command, so disable it and consume Escape explicitly below.
        CEDWindowBase.__init__(self, xaml_source=_XAML, theme_aware=True,
                               handle_esc=False)

        self._status_brushes = {
            "ready": self._theme_resource("CED.Brush.AccentGreen", BRUSH_READY),
            "warn": self._theme_resource("CED.Brush.BadgeWarningText", BRUSH_WARN),
            "bad": self._theme_resource("CED.Brush.AccentRed", BRUSH_BAD),
            "off": self._theme_resource("CED.Brush.SecondaryText", BRUSH_OFF),
            "info": self._theme_resource("CED.Brush.BadgeInfoText", BRUSH_INFO),
        }

        self.PanelOptions = List[str]()
        for p in panel_options:
            self.PanelOptions.Add(p)
        self.RatingOptions = List[str]()
        for r in rating_options:
            self.RatingOptions.Add(r)
        self.GroupParamOptions = List[str]()
        for gp in (group_param_options or []):
            self.GroupParamOptions.Add(gp)
        self.NameParamOptions = List[str]()
        for np in (name_param_options if name_param_options is not None
                   else (group_param_options or [])):
            self.NameParamOptions.Add(np)

        self._name_to_id = name_to_id
        self._panel_info = panel_info or {}
        self._gateway = gateway
        self._document_key = str(document_key or "")
        self._document_title = str(document_title or "-")
        self._scope_label = str(scope_label or "model")
        self._scope_ids = list(scope_ids) if scope_ids is not None else None
        self._replacing_window = False
        self.result_plans = None
        self._drag_start = None
        self._drag_items = []
        self._drag_started = False
        self._drag_allowed = False
        self._mouse_down_item = None
        self._mouse_down_selection = []
        self._defer_selection = False
        self._pending_checkbox_selection = []
        self._checkbox_selection_restore = []
        self._grid_anchor_item = None
        self._last_group_selection = None
        self._drag_target = None
        self._suppress_regroup = True
        self._checkbox_bulk_edit_in_progress = False
        self._view_refresh_depth = 0
        self._view_refresh_pending = False
        self._view_defer = None
        self._members_by_group = {}

        # which parameter the circuits are grouped on (user-switchable)
        self._group_param = default_group_param or (
            group_param_options[0] if group_param_options else "")
        # which parameter SEEDS each circuit's Load Name (user-switchable,
        # independent of grouping). Must be set before _rebuild_groups, which
        # seeds the load names from it.
        _name_opts = (name_param_options if name_param_options is not None
                      else (group_param_options or []))
        self._name_param = default_name_param or (
            _name_opts[0] if _name_opts else "")
        # The picker stores its visible option text; synthetic options map to
        # the collector's per-row grouping keys before grouping is rebuilt.
        self._active_key = self._group_key_for_option(self._group_param)

        # build the row VMs, then group them by the chosen parameter
        self._rows = []
        for row_order, data in enumerate(rows_data):
            vm = RowVM(data, status_brushes=self._status_brushes)
            # Keep a deterministic member order for the grouped ICollectionView
            # so WPF range selection follows the same order the user sees.
            vm.row_order = row_order
            self._rows.append(vm)
        self._groups = []
        self._rebuild_groups(self._active_key)

        self.DataContext = self

        self.Grid = self.FindName("Grid")
        self.SummaryText = self.FindName("SummaryText")
        self.ValidationText = self.FindName("ValidationText")
        self.GroupByCombo = self.FindName("GroupByCombo")
        self.NameByCombo = self.FindName("NameByCombo")
        self.GroupSelectionText = self.FindName("GroupSelectionText")
        self.GridSelectionText = self.FindName("GridSelectionText")
        self.DocumentText = self.FindName("DocumentText")
        self.ScopeText = self.FindName("ScopeText")
        self.ActionStatusText = self.FindName("ActionStatusText")
        self.RefreshButton = self.FindName("RefreshButton")
        self.RunButton = self.FindName("RunButton")
        self.NewGroupButton = self.FindName("NewGroupButton")
        self.DedicatedGroupButton = self.FindName("DedicatedGroupButton")
        self.BulkSetButton = self.FindName("BulkSetButton")
        self.ToolInfoButton = self.FindName("ToolInfoButton")
        self.ToolInfoPopup = self.FindName("ToolInfoPopup")
        self.CircuitStatusPopup = self.FindName("CircuitStatusPopup")
        self.CircuitStatusPopupText = self.FindName("CircuitStatusPopupText")
        self.BulkValuesPopup = self.FindName("BulkValuesPopup")
        self.BulkPopupLoadName = self.FindName("BulkPopupLoadName")
        self.BulkPopupIncrement = self.FindName("BulkPopupIncrement")
        self.BulkPopupPanel = self.FindName("BulkPopupPanel")
        self.BulkPopupRating = self.FindName("BulkPopupRating")
        self.BulkPopupNotes = self.FindName("BulkPopupNotes")
        if self.BulkPopupPanel is not None:
            self.BulkPopupPanel.ItemsSource = self.PanelOptions
        if self.BulkPopupRating is not None:
            self.BulkPopupRating.ItemsSource = self.RatingOptions
        self._bulk_target_groups = []
        self._column_menu = None
        self._run_in_progress = False
        self._results_mode = False

        self._items = ObservableCollection[object]()
        for vm in self._rows:
            self._items.Add(vm)

        self._cvs = CollectionViewSource()
        self._cvs.Source = self._items
        # One grouping level: each WPF group is one circuit group. Readiness is
        # validation text on the header, never a second visual grouping level.
        self._cvs.GroupDescriptions.Add(PropertyGroupDescription("group"))
        if SortDescription is not None and ListSortDirection is not None:
            self._cvs.SortDescriptions.Add(
                SortDescription("group_key", ListSortDirection.Ascending)
            )
            self._cvs.SortDescriptions.Add(
                SortDescription("row_order", ListSortDirection.Ascending)
            )
        self._view = self._cvs.View
        self.Grid.ItemsSource = self._view

        self._update_group_value_display()
        if self.DocumentText is not None:
            self.DocumentText.Text = "Document: {}".format(self._document_title)
        if self.ScopeText is not None:
            self.ScopeText.Text = "Scope: {}".format(self._scope_label)

        # seed the combos without triggering a redundant regroup/reseed
        if self.GroupByCombo is not None and self._group_param:
            self.GroupByCombo.SelectedItem = self._group_param
        if self.NameByCombo is not None and self._name_param:
            self.NameByCombo.SelectedItem = self._name_param
        self._suppress_regroup = False

        self._update_group_selection_summary()
        self._update_grid_selection_summary()
        self._revalidate()

    def _set_action_status(self, text):
        if self.ActionStatusText is not None:
            self.ActionStatusText.Text = str(text or "")

    def tool_info_toggled(self, sender, args):
        if self.ToolInfoPopup is not None:
            self.ToolInfoPopup.IsOpen = bool(sender.IsChecked)

    def tool_info_popup_closed(self, sender, args):
        if self.ToolInfoButton is not None and self.ToolInfoButton.IsChecked:
            self.ToolInfoButton.IsChecked = False

    def group_status_clicked(self, sender, args):
        group = sender.DataContext
        if not isinstance(group, GroupVM) or self.CircuitStatusPopup is None:
            return
        if self.CircuitStatusPopupText is not None:
            self.CircuitStatusPopupText.Text = str(group.status_detail or group.status)
        self.CircuitStatusPopup.PlacementTarget = sender
        self.CircuitStatusPopup.IsOpen = True
        args.Handled = True

    def grid_header_right_click(self, sender, args):
        """Open a show/hide menu from any visible column header."""
        node = args.OriginalSource
        header = None
        while node is not None:
            if isinstance(node, DataGridColumnHeader):
                header = node
                break
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                node = None
        if header is None:
            return

        menu = ContextMenu()
        for index, column in enumerate(list(self.Grid.Columns)):
            # The first column owns expand/collapse and inclusion controls. It
            # remains visible so the menu is always recoverable.
            if index == 0:
                continue
            item = MenuItem()
            item.Header = str(column.Header or "Column")
            item.IsCheckable = True
            item.IsChecked = column.Visibility == Visibility.Visible
            item.Tag = column
            item.Click += self.column_visibility_clicked
            menu.Items.Add(item)
        menu.PlacementTarget = header
        self._column_menu = menu
        menu.IsOpen = True
        args.Handled = True

    def column_visibility_clicked(self, sender, args):
        column = getattr(sender, "Tag", None)
        if column is None:
            return
        column.Visibility = (Visibility.Visible
                             if bool(sender.IsChecked)
                             else Visibility.Collapsed)

    def window_preview_key_down(self, sender, args):
        """Keep Escape from closing the modeless tool or leaking to Revit."""
        try:
            if args.Key == Key.Escape or args.SystemKey == Key.Escape:
                args.Handled = True
        except Exception:
            pass

    def window_activated(self, sender, args):
        """Retarget after the user switches Revit documents.

        The external event performs the API-side document check. It only
        recollects devices when the active document is actually different, so
        normal focus changes do not rescan a large model.
        """
        if self._replacing_window or self._gateway is None:
            return
        if self._gateway.is_busy():
            return
        self._gateway.raise_sync(self._document_key, self._handle_sync_complete)

    def _handle_sync_complete(self, status, op_name, result, error):
        if status != "ok":
            self._set_action_status(str(error or "Could not track the active document."))
            return
        data = result or {}
        if data.get("changed") and data.get("snapshot"):
            self._replace_with_snapshot(
                data.get("snapshot"), "Target document updated.")

    def _replace_with_snapshot(self, snapshot, completion_message=None):
        """Queue a clean replacement after the Revit external event returns."""
        if not snapshot:
            self._set_action_status("Refresh returned no device data.")
            return
        self._replacing_window = True
        try:
            self.Dispatcher.BeginInvoke(
                Action(lambda: self._finish_snapshot_replacement(
                    snapshot, completion_message)),
                DispatcherPriority.Background,
            )
        except Exception:
            self._finish_snapshot_replacement(snapshot, completion_message)

    def _finish_snapshot_replacement(self, snapshot, completion_message=None):
        try:
            replacement = show_modeless(snapshot, self._gateway, activate=True)
            if replacement is not None:
                if completion_message:
                    replacement._set_action_status(completion_message)
                self.Close()
                return
            raise RuntimeError("The refreshed window could not be opened.")
        except Exception as exc:
            self._replacing_window = False
            message = "Refresh failed: {}".format(exc)
            self._set_action_status(message)
            forms.alert(message, title="Create Circuits by Device Parameter")

    def _replace_string_list(self, target, values, keep_value=""):
        incoming = [str(value) for value in list(values or [])]
        if keep_value and keep_value not in incoming:
            incoming.insert(0, str(keep_value))
        target.Clear()
        for value in incoming:
            target.Add(value)

    def _apply_snapshot_in_place(self, snapshot, completion_message=None):
        """Refresh Revit data while retaining groups and staged group values."""
        data = dict(snapshot or {})
        incoming_rows = list(data.get("rows_data") or [])
        incoming_by_id = {}
        for row_data in incoming_rows:
            try:
                incoming_by_id[int(row_data.get("element_id"))] = row_data
            except Exception:
                continue

        self._begin_view_batch()
        try:
            existing_by_id = dict(
                (int(row.element_id), row) for row in list(self._rows))

            for element_id, row in list(existing_by_id.items()):
                row_data = incoming_by_id.get(element_id)
                if row_data is not None:
                    row.refresh_from_data(row_data)
                    continue
                self._rows.remove(row)
                try:
                    self._items.Remove(row)
                except Exception:
                    pass

            groups_by_key = {}
            for group in list(self._groups):
                groups_by_key.setdefault(str(group.key or ""), group)
            new_groups = []
            for element_id, row_data in list(incoming_by_id.items()):
                if element_id in existing_by_id:
                    continue
                row = RowVM(row_data, status_brushes=self._status_brushes)
                key = str(
                    row.group_values.get(self._active_key, "") or "").strip()
                group = groups_by_key.get(key)
                if group is None:
                    group = GroupVM(key, status_brushes=self._status_brushes)
                    group.is_expanded = (
                        len(incoming_rows) <= INITIAL_EXPANDED_ROW_LIMIT)
                    groups_by_key[key] = group
                    self._groups.append(group)
                    new_groups.append(group)
                row.move_to(group)
                self._rows.append(row)
                self._items.Add(row)

            self._reindex_members()
            self._groups = [
                group for group in self._groups
                if group in self._members_by_group
            ]
            for group in new_groups:
                members = self._members_by_group.get(group, [])
                if members:
                    self._init_group_defaults(group, members)

            self._name_to_id = dict(data.get("name_to_id") or {})
            self._panel_info = dict(data.get("panel_info") or {})
            self._replace_string_list(
                self.PanelOptions, data.get("panel_names") or [])
            self._replace_string_list(
                self.GroupParamOptions,
                data.get("group_param_options") or [],
                self._group_param,
            )
            self._replace_string_list(
                self.NameParamOptions,
                data.get("name_param_options") or [],
                self._name_param,
            )

            self._suppress_regroup = True
            try:
                if self.GroupByCombo is not None:
                    self.GroupByCombo.ItemsSource = self.GroupParamOptions
                    self.GroupByCombo.SelectedItem = self._group_param
                if self.NameByCombo is not None:
                    self.NameByCombo.ItemsSource = self.NameParamOptions
                    self.NameByCombo.SelectedItem = self._name_param
                if self.BulkPopupPanel is not None:
                    self.BulkPopupPanel.ItemsSource = self.PanelOptions
            finally:
                self._suppress_regroup = False

            self._document_title = str(
                data.get("document_title") or self._document_title)
            self._scope_label = str(
                data.get("scope_label") or self._scope_label)
            self._scope_ids = (
                list(data.get("scope_ids"))
                if data.get("scope_ids") is not None else None)
            if self.DocumentText is not None:
                self.DocumentText.Text = "Document: {}".format(
                    self._document_title)
            if self.ScopeText is not None:
                self.ScopeText.Text = "Scope: {}".format(self._scope_label)

            self._update_group_value_display()
            self._refresh_view()
            self._revalidate()
        finally:
            self._end_view_batch()

        if completion_message:
            self._set_action_status(completion_message)

    def _theme_resource(self, key, fallback):
        """Resolve a required CED theme brush."""
        value = self.FindResource(key)
        return value if value is not None else fallback

    @staticmethod
    def _group_key_for_option(option):
        """Translate a visible Group By option into a row-data key."""
        if option == cg_core.SPACE_GROUP_OPTION:
            return cg_core.SPACE_GROUP_KEY
        if option == cg_core.IDENTITY_GROUP_OPTION:
            return cg_core.IDENTITY_GROUP_KEY
        return option

    def _reindex_members(self):
        """Build the group-to-member index once for validation and commands."""
        members_by_group = {}
        for row in self._rows:
            members_by_group.setdefault(row.group, []).append(row)
        self._members_by_group = members_by_group
        return members_by_group

    def _rebuild_groups(self, param_name):
        """(Re)build the GroupVMs by grouping every row on ``param_name`` (a
        parameter name or a synthetic Space/Identity key). Pure model work -
        does not touch the view (callers refresh)."""
        self._active_key = param_name
        old_expansion = {}
        for old_group in self._groups:
            old_members = self._members_by_group.get(old_group, [])
            old_is_done = bool(old_members and all(
                row.already_circuited for row in old_members))
            old_expansion[(old_group.key, old_is_done)] = old_group.is_expanded

        groups_by_key = {}
        members_by_group = {}
        self._groups = []
        for r in self._rows:
            key = (r.group_values.get(param_name, "") or "").strip()
            g = groups_by_key.get(key)
            if g is None:
                g = GroupVM(key, status_brushes=self._status_brushes)
                groups_by_key[key] = g
                self._groups.append(g)
            r.move_to(g)
            members_by_group.setdefault(g, []).append(r)
        for g in self._groups:
            members = members_by_group.get(g, [])
            self._init_group_defaults(g, members)
            # Preserve an existing group's expansion state. New groups are
            # collapsed for large datasets so WPF does not realize every
            # member row and group-header control on first paint.
            is_done = bool(members and all(r.already_circuited for r in members))
            prior = old_expansion.get((g.key, is_done))
            if prior is not None:
                g.is_expanded = bool(prior)
            elif is_done:
                g.is_expanded = False
            else:
                g.is_expanded = len(self._rows) <= INITIAL_EXPANDED_ROW_LIMIT
        self._members_by_group = members_by_group
        self._refresh_group_panel_options(members_by_group)
        self._auto_assign_panels(members_by_group)
        self._refresh_group_panel_options(members_by_group)

    @staticmethod
    def _voltage_matches(left, right):
        try:
            return abs(float(left) - float(right)) <= 0.5
        except Exception:
            return False

    def _compatible_panel_names(self, members):
        """Return panels whose distribution system can serve these members.

        The shared panel repository already resolves the distribution system's
        line-to-ground and line-to-line voltages through Revit's unit APIs.
        A one-pole load can use either reported voltage; two- and three-pole
        loads require line-to-line voltage. Unknown member data is left
        unfiltered so a project can still use Revit's native assignability
        check as the final authority.
        """
        names = [str(x) for x in list(self.PanelOptions or [])]
        effective = cg_core.effective_rows(members) or list(members or [])
        voltages = set(
            r.voltage_key for r in effective
            if getattr(r, "voltage_key", None) is not None
        )
        poles = set(
            int(r.poles_value) for r in effective
            if getattr(r, "poles_value", None)
        )
        if len(voltages) != 1 or len(poles) != 1:
            return names
        voltage = list(voltages)[0]
        pole_count = list(poles)[0]
        compatible = []
        for name in names:
            info = self._panel_info.get(name) or {}
            profile = info.get("profile") or {}
            lg = profile.get("lg_voltage")
            ll = profile.get("ll_voltage")
            phase = str(profile.get("phase", "") or "").lower()
            is_single_phase = "singlephase" in phase.replace(" ", "")
            if pole_count == 3 and is_single_phase:
                continue
            if pole_count == 1:
                candidate_voltages = (lg, ll)
            else:
                candidate_voltages = (ll,)
            if any(self._voltage_matches(voltage, candidate)
                   for candidate in candidate_voltages if candidate is not None):
                compatible.append(name)
        return compatible

    def _refresh_group_panel_options(self, members_by_group=None):
        members_by_group = members_by_group or self._members_by_group
        for group in list(self._groups):
            members = members_by_group.get(group, [])
            options = self._compatible_panel_names(members)
            group.set_panel_options(options)
            current = (group.panel or "").strip()
            # A raw multi-panel source value is resolved by _auto_assign_panels
            # below. Only clear a single selected panel here.
            if (current and current not in options and
                    len(cg_core.parse_panel_candidates(current)) <= 1):
                group.panel = ""
                group.notify("panel")
                group.panel_note = "Selected panel does not match the group's voltage/poles."
                group.notify("panel_note")

    def _auto_assign_panels(self, members_by_group=None):
        """Resolve groups whose source panel value lists MULTIPLE panels
        ('RA, RB, RC, RD') to the closest listed panel that still has room
        for the circuit. Slots are reserved across this session's groups, so
        shared candidate lists overflow to the next-closest panel. When no
        listed panel has room the combo is left blank and the group carries a
        validation note. Single-name values keep today's behavior."""
        if not self._panel_info:
            return
        if members_by_group is None:
            members_by_group = self._members_map()
        requests = []
        group_by_key = {}
        for g in self._groups:
            members = members_by_group.get(g, [])
            if members and all(r.already_circuited for r in members):
                continue
            # multi-panel means the RAW text lists several names, whether or
            # not they all exist in the model (unknown-only lists still get
            # blanked + noted rather than riding along as a bogus combo value)
            if len(cg_core.parse_panel_candidates(g.panel)) < 2:
                continue
            compatible = set(self._compatible_panel_names(members))
            candidates = [candidate for candidate in cg_core.parse_panel_candidates(
                g.panel, self._panel_info.keys()) if candidate in compatible]
            eff = cg_core.effective_rows(members)
            poles = None
            for r in (eff or members):
                if r.poles_value:
                    poles = r.poles_value
                    break
            requests.append({
                "group_key": g.key,
                "candidates": candidates,
                "centroid": cg_core.centroid(
                    [r.location for r in (eff or members)]),
                "poles": poles,
            })
            group_by_key[g.key] = g
        if not requests:
            return
        results = cg_core.resolve_panel_assignments(requests, self._panel_info)
        for key, res in results.items():
            g = group_by_key.get(key)
            if g is None:
                continue
            g.panel = res.get("panel") or ""
            g.panel_note = res.get("note") or ""
            g.notify("panel")

    def _apply_grouping(self, param_name):
        self._begin_view_batch()
        try:
            self._clear_group_selection()
            self._rebuild_groups(param_name)
            self._update_group_value_display()
            self._refresh_view()
            self._revalidate()
        finally:
            self._end_view_batch()

    def group_param_changed(self, sender, args):
        if self._suppress_regroup or self._results_mode:
            return
        item = sender.SelectedItem
        value = str(item) if item is not None else ""
        if not value or value == self._group_param:
            return
        self._group_param = value
        key = self._group_key_for_option(value)
        if key and key != self._active_key:
            self._apply_grouping(key)

    def _update_group_value_display(self):
        """Project the selected Group By value into its own live column."""
        for row in self._rows:
            row.group_value = str(row.group_values.get(self._active_key, "") or "")
            row.notify("group_value")

    def name_param_changed(self, sender, args):
        """Re-seed every circuit's Load Name from the newly chosen name-by
        parameter (replacing any hand edits, since the user asked for a new
        naming source)."""
        if self._suppress_regroup or self._results_mode:
            return
        item = sender.SelectedItem
        value = str(item) if item is not None else ""
        if not value or value == self._name_param:
            return
        self._name_param = value
        members_by_group = self._members_map()
        for g in self._groups:
            self._seed_load_name(g, members_by_group.get(g, []))
            g.notify("load_name")
        # load_name is not a grouping/sort key, so the header bindings update
        # on their own - no view refresh needed here
        self._revalidate()

    # -- setup helpers ----------------------------------------------------
    def _members(self, group):
        return list(self._members_by_group.get(group, []))

    def _members_map(self):
        """All groups' member lists in ONE pass over the rows. Anything that
        touches every group must use this instead of per-group _members()
        scans, which are O(groups x rows) and visibly stall at 5000 items."""
        return self._members_by_group

    def _seed_load_name(self, group, members=None):
        """Seed the group's Load Name from the members' values of the name-by
        parameter; falls back to the Identity-Mark common stem, then the group
        key, when that parameter is blank on every member."""
        if members is None:
            members = self._members(group)
        fallback = cg_core.default_load_name(
            [r.identity_mark for r in members], fallback=group.key)
        group.load_name = cg_core.name_from_values(
            [r.group_values.get(self._name_param, "") for r in members],
            fallback=fallback)

    def _init_group_defaults(self, group, members=None):
        if members is None:
            members = self._members(group)
        self._seed_load_name(group, members)
        for r in members:
            if r.src_panel:
                group.panel = r.src_panel
                break
        for r in members:
            amps, valid, _ = cg_core.parse_rating(r.src_rating)
            if valid:
                group.rating = cg_core.format_amps_number(amps)
                break
        group.schedule_notes = ""
        for r in members:
            if r.src_schedule_notes:
                group.schedule_notes = r.src_schedule_notes
                break

    def _all_group_keys(self):
        return set(g.key for g in self._groups)

    def _refresh_view(self):
        """Queue one grouped-view refresh, coalescing batch changes."""
        self._view_refresh_pending = True
        if self._view_refresh_depth:
            return
        self._perform_view_refresh()

    def _perform_view_refresh(self):
        if not self._view_refresh_pending:
            return
        self._view_refresh_pending = False
        try:
            self._view.Refresh()
        except Exception:
            pass
        self._update_grid_selection_summary()

    def _begin_view_batch(self):
        if self._view_refresh_depth == 0:
            try:
                self._view_defer = self._view.DeferRefresh()
            except Exception:
                self._view_defer = None
        self._view_refresh_depth += 1

    def _end_view_batch(self):
        if self._view_refresh_depth:
            self._view_refresh_depth -= 1
        if not self._view_refresh_depth:
            defer = self._view_defer
            self._view_defer = None
            if defer is not None:
                try:
                    defer.Dispose()
                except Exception:
                    pass
            self._perform_view_refresh()

    def _selected_groups(self):
        """Return selected circuit rows represented by GroupVM headers."""
        return [group for group in self._groups if group.is_selected]

    def _clear_grid_selection(self):
        if self.Grid is None:
            return
        try:
            self.Grid.UnselectAll()
        except Exception:
            try:
                self.Grid.SelectedItems.Clear()
            except Exception:
                pass
        try:
            self.Grid.SelectedItem = None
        except Exception:
            pass
        self._update_grid_selection_summary()

    def _update_grid_selection_summary(self):
        if getattr(self, "GridSelectionText", None) is None:
            return
        count = 0
        try:
            count = len(list(self.Grid.SelectedItems or []))
        except Exception:
            pass
        self.GridSelectionText.Text = "{} element row(s) selected".format(count)

    def grid_selection_changed(self, sender, args):
        """Keep the visible member-row selection count synchronized."""
        self._update_grid_selection_summary()

    def _clear_group_selection(self, keep=None):
        for group in self._groups:
            if group is not keep:
                group.set_selected(False)
        self._update_group_selection_summary()

    def _ordered_groups(self):
        return sorted(
            self._groups,
            key=lambda group: str(group.key or ""),
        )

    def _ordered_visible_rows(self):
        """Return member rows in the same order as the grouped DataGrid view."""
        rows = sorted(
            self._rows,
            key=lambda row: (
                str(row.group_key or ""),
                getattr(row, "row_order", 0),
            ),
        )
        return [row for row in rows if getattr(row.group, "is_expanded", True)]

    def _update_group_selection_summary(self):
        if getattr(self, "GroupSelectionText", None) is None:
            return
        count = len(self._selected_groups())
        if count:
            self.GroupSelectionText.Text = "{} circuit row(s) selected".format(count)
        else:
            self.GroupSelectionText.Text = "No circuit rows selected"

    def _group_header_interactive_source(self, source):
        """Return True when a header click originated in an input/control."""
        node = source
        while node is not None:
            if isinstance(node, (Button, ComboBox, TextBox, ToggleButton, CheckBox)):
                return True
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                return False
        return False

    def group_header_mouse_down(self, sender, args):
        """Select circuit groups independently from DataGrid member rows."""
        group = getattr(sender, "DataContext", None)
        if not isinstance(group, GroupVM):
            return
        if self._group_header_interactive_source(args.OriginalSource):
            self._clear_grid_selection()
            self._grid_anchor_item = None
            # Do not collapse a Ctrl-selected set when the user activates a
            # combo, load-name field, or expand arrow on one of its groups.
            if not self._selected_groups():
                group.set_selected(True)
                self._update_group_selection_summary()
            return
        self._clear_grid_selection()
        self._grid_anchor_item = None
        ctrl_down = bool(Keyboard.Modifiers & ModifierKeys.Control)
        shift_down = bool(Keyboard.Modifiers & ModifierKeys.Shift)
        ordered = self._ordered_groups()
        if shift_down and self._last_group_selection in ordered:
            anchor = ordered.index(self._last_group_selection)
            current = ordered.index(group)
            lo, hi = sorted((anchor, current))
            if not ctrl_down:
                self._clear_group_selection()
            for target in ordered[lo:hi + 1]:
                target.set_selected(True)
        elif ctrl_down:
            group.set_selected(not group.is_selected)
        else:
            self._clear_group_selection(keep=group)
            group.set_selected(True)
        self._last_group_selection = group
        self._update_group_selection_summary()
        args.Handled = True

    def window_mouse_down(self, sender, args):
        """Clicking outside the member grid clears both selection models."""
        if self.Grid is None:
            return
        if not self._is_within(args.OriginalSource, self.Grid):
            # PreviewMouseDown runs before Button.Click. Preserve the member
            # selection for the one toolbar command that consumes it.
            if self._is_selection_command_source(args.OriginalSource):
                return
            self._clear_grid_selection()
            self._clear_group_selection()
            self._grid_anchor_item = None
            self._last_group_selection = None

    def _is_selection_command_source(self, source):
        """Return True for toolbar commands that consume member selection."""
        node = source
        while node is not None:
            if isinstance(node, Button):
                return getattr(node, "Name", "") in (
                    "NewGroupButton", "DedicatedGroupButton", "BulkSetButton",
                    "SelectDeviceButton", "ShowDeviceButton",
                    "SelectGroupButton", "ShowGroupButton",
                )
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                return False
        return False

    def _find_group_from_source(self, source):
        node = source
        while node is not None:
            data = getattr(node, "DataContext", None)
            if isinstance(data, GroupVM):
                return data
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                return None
        return None

    def _find_drop_group_from_source(self, source):
        group = self._find_group_from_source(source)
        if group is not None:
            return group
        row = self._find_row_item(source)
        return row.group if isinstance(row, RowVM) else None

    def expand_collapse_all_clicked(self, sender, args):
        """Toggle every circuit group's expanded state in one UI pass."""
        if not self._groups:
            return
        expand = not all(group.is_expanded for group in self._groups)
        for group in self._groups:
            group.set_expanded(expand)
        try:
            sender.ToolTip = "Collapse all groups" if expand else "Expand all groups"
        except Exception:
            pass

    # -- validation -------------------------------------------------------
    def _revalidate(self):
        members_by_group = self._members_by_group
        self._groups = [g for g in self._groups if g in members_by_group]

        ready = 0
        not_ready = 0
        lines = []
        self._refresh_group_panel_options(members_by_group)

        for g in self._groups:
            members = members_by_group.get(g, [])
            # per-row status
            for r in members:
                if r.already_circuited:
                    r.set_status("Element already circuited", self._status_brushes["info"])
                elif not r.include:
                    r.set_status("Excluded", self._status_brushes["off"])
                else:
                    r.set_status("", self._status_brushes["ready"])

            eff = cg_core.effective_rows(members)
            if not eff:
                if members and all(r.already_circuited for r in members):
                    g.set_status("Element already circuited", self._status_brushes["info"])
                else:
                    g.set_status("No items selected", self._status_brushes["off"])
                g.set_warning(False)
                continue

            problems = cg_core.validate_members(eff)
            if not (g.panel or "").strip():
                problems.append("Panel is not selected")
                if g.panel_note:
                    problems.append(g.panel_note)
            amps, valid, standard = cg_core.parse_rating(g.rating)
            label = g.load_name or g.key
            if problems:
                g.set_status(u"⚠ " + "; ".join(problems), self._status_brushes["bad"])
                g.set_warning(False)
                not_ready += 1
                lines.append(u"'{}': {}".format(label, "; ".join(problems)))
            elif not valid:
                g.set_status("Invalid rating - not ready", self._status_brushes["bad"])
                g.set_warning(False)
                not_ready += 1
                lines.append(u"'{}': invalid breaker rating '{}'".format(label, g.rating))
            else:
                g.set_status("Ready", self._status_brushes["ready"])
                g.set_warning(not standard)
                ready += 1
                if not standard:
                    lines.append(u"'{}': non-standard breaker size ({} A)".format(
                        label, cg_core.format_amps_number(amps)))
        if ready and not_ready:
            summary = "{} circuit{} ready | {} circuit{} {} attention".format(
                ready, "" if ready == 1 else "s",
                not_ready, "" if not_ready == 1 else "s",
                "needs" if not_ready == 1 else "need")
        elif not_ready:
            summary = "{} circuit{} {} attention".format(
                not_ready, "" if not_ready == 1 else "s",
                "needs" if not_ready == 1 else "need")
        elif ready:
            summary = "{} circuit{} ready".format(
                ready, "" if ready == 1 else "s")
        else:
            summary = "No circuits available"
        self.SummaryText.Text = summary
        self.ValidationText.Text = "\n".join(lines) if lines else "No validation issues."
    # -- group header handlers (sender.DataContext == GroupVM) -----------
    def group_load_name_changed(self, sender, args):
        if self._results_mode:
            return
        g = sender.DataContext
        if not isinstance(g, GroupVM):
            return
        g.load_name = sender.Text or ""

    def group_schedule_notes_changed(self, sender, args):
        if self._results_mode:
            return
        g = sender.DataContext
        if not isinstance(g, GroupVM):
            return
        value = sender.Text or ""
        if g.schedule_notes == value:
            return
        g.schedule_notes = value
        g.notify("schedule_notes")

    def group_panel_changed(self, sender, args):
        if self._results_mode:
            return
        g = sender.DataContext
        if not isinstance(g, GroupVM):
            return
        value = ""
        try:
            added = list(args.AddedItems or [])
            if added:
                value = str(added[-1]).strip()
        except Exception:
            pass
        if not value:
            value = (sender.Text or "").strip()
        if not value:
            return
        if g.panel == value:
            return
        g.panel = value
        g.notify("panel")
        if g.panel_note:
            g.panel_note = ""
            g.notify("panel_note")
        self._revalidate()

    def group_rating_changed(self, sender, args):
        if self._results_mode:
            return
        g = sender.DataContext
        if not isinstance(g, GroupVM):
            return
        value = ""
        try:
            added = list(args.AddedItems or [])
            if added:
                value = str(added[-1]).strip()
        except Exception:
            pass
        if not value:
            value = (sender.Text or "").strip()
        if not value:
            return
        if g.rating == value:
            return
        g.rating = value
        g.notify("rating")
        self._revalidate()

    # -- rating select-all on focus --------------------------------------
    def _combo_editbox(self, combo):
        try:
            combo.ApplyTemplate()
            return combo.Template.FindName("PART_EditableTextBox", combo)
        except Exception:
            return None

    def _is_within(self, node, target):
        while node is not None:
            if node is target:
                return True
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                return False
        return False

    def rating_got_focus(self, sender, args):
        tb = self._combo_editbox(sender)
        if tb is not None:
            tb.SelectAll()

    def rating_mouse_down(self, sender, args):
        tb = self._combo_editbox(sender)
        if tb is None or tb.IsKeyboardFocusWithin:
            return
        if self._is_within(args.OriginalSource, tb):
            tb.Focus()
            tb.SelectAll()
            args.Handled = True

    # -- member row handlers ---------------------------------------------
    def row_include_changed(self, sender, args):
        if self._results_mode:
            return
        if self._checkbox_bulk_edit_in_progress:
            return
        vm = sender.DataContext
        if vm is None:
            return
        new_value = bool(sender.IsChecked)
        targets = list(self._pending_checkbox_selection)
        if vm not in targets:
            targets = [vm]
        targets = [target for target in targets if target.can_include]
        if not targets or all(target.include == new_value for target in targets):
            self._pending_checkbox_selection = []
            return
        self._begin_view_batch()
        try:
            self._checkbox_bulk_edit_in_progress = True
            for target in targets:
                target.set_include(new_value)
            self._revalidate()
        finally:
            self._checkbox_bulk_edit_in_progress = False
            self._pending_checkbox_selection = []
            self._end_view_batch()

    # -- toolbar ----------------------------------------------------------
    def _selected_rows(self):
        """Return selected element/member rows from the DataGrid."""
        try:
            selected = list(self.Grid.SelectedItems or [])
        except Exception:
            selected = []
        return [item for item in selected if isinstance(item, RowVM)]

    def _device_target_ids(self):
        return [int(row.element_id) for row in self._selected_rows()]

    def _group_target_ids(self):
        """Return every device in the UI-selected group(s).

        Circuit headers are the primary group selection. When the user has
        selected member rows instead, use the group(s) containing those rows;
        this deliberately makes a single selected device resolve to its whole
        group for the Select Group / Show Group commands.
        """
        groups = list(self._selected_groups())
        if not groups:
            for row in self._selected_rows():
                if row.group is not None and row.group not in groups:
                    groups.append(row.group)
        result = []
        seen = set()
        for group in groups:
            for row in self._members_by_group.get(group, []):
                element_id = int(row.element_id)
                if element_id not in seen:
                    seen.add(element_id)
                    result.append(element_id)
        return result

    def _raise_navigation(self, target_kind, show):
        ids = (self._device_target_ids() if target_kind == "device"
               else self._group_target_ids())
        if not ids:
            label = "device row(s)" if target_kind == "device" else "device or circuit row(s)"
            forms.alert("Select {} in the tool first.".format(label),
                        title="Create Circuits by Device Parameter")
            return
        # Selection and ShowElements are Revit UI operations and do not need an
        # ExternalEvent. Circuit Manager uses the same direct modeless pattern;
        # keeping these out of the write gateway also means a slow refresh can
        # never block model navigation.
        uidoc = getattr(revit, "uidoc", None)
        doc = uidoc.Document if uidoc is not None else None
        if doc is None:
            self._set_action_status("No active Revit document.")
            return
        try:
            current_key = str(doc.GetHashCode())
        except Exception:
            current_key = "{}|{}".format(
                getattr(doc, "Title", "") or "",
                getattr(doc, "PathName", "") or "",
            )
        if current_key != self._document_key:
            self._set_action_status(
                "The active document changed. Refreshing the tool target...")
            if self._gateway is not None:
                self._gateway.raise_refresh(
                    self._document_key, self._scope_ids,
                    self._handle_refresh_complete)
            return

        native_ids = List[DB.ElementId]()
        for value in ids:
            element_id = revit_helpers.elementid_from_value(value)
            if doc.GetElement(element_id) is not None:
                native_ids.Add(element_id)
        if native_ids.Count == 0:
            self._set_action_status("None of the chosen devices still exist.")
            return
        try:
            if show:
                uidoc.ShowElements(native_ids)
                action = "Shown"
            else:
                uidoc.Selection.SetElementIds(native_ids)
                action = "Selected"
            self._set_action_status(
                "{} {} device(s).".format(action, native_ids.Count))
        except Exception as exc:
            message = "Model navigation failed: {}".format(exc)
            self._set_action_status(message)
            forms.alert(message, title="Create Circuits by Device Parameter")

    def _handle_navigation_complete(self, status, op_name, result, error):
        if status == "ok":
            snapshot = (result or {}).get("retarget_snapshot")
            if snapshot:
                self._replace_with_snapshot(snapshot)
                return
            count = int((result or {}).get("count") or 0)
            action = "Shown" if (result or {}).get("show") else "Selected"
            self._set_action_status("{} {} device(s).".format(action, count))
            return
        self._set_action_status(str(error or "Model navigation failed."))
        forms.alert(str(error or "Model navigation failed."),
                    title="Create Circuits by Device Parameter")

    def select_device_clicked(self, sender, args):
        self._raise_navigation("device", False)

    def show_device_clicked(self, sender, args):
        self._raise_navigation("device", True)

    def select_group_clicked(self, sender, args):
        self._raise_navigation("group", False)

    def show_group_clicked(self, sender, args):
        self._raise_navigation("group", True)

    def refresh_clicked(self, sender, args):
        if self._gateway is None:
            return
        self._set_action_status(
            "Re-reading device data for {}...".format(self._scope_label))
        raised = self._gateway.raise_refresh(
            self._document_key, self._scope_ids, self._handle_refresh_complete)
        if raised:
            if self.RefreshButton is not None:
                self.RefreshButton.IsEnabled = False
        else:
            self._set_action_status("Another Revit action is still pending.")

    def _handle_refresh_complete(self, status, op_name, result, error):
        try:
            self.Dispatcher.BeginInvoke(
                Action(lambda: self._finish_refresh_complete(
                    status, result, error)),
                DispatcherPriority.Background,
            )
        except Exception:
            self._finish_refresh_complete(status, result, error)

    def _finish_refresh_complete(self, status, result, error):
        if self.RefreshButton is not None:
            self.RefreshButton.IsEnabled = True
        if status != "ok":
            self._set_action_status(str(error or "Refresh failed."))
            forms.alert(str(error or "Refresh failed."),
                        title="Create Circuits by Device Parameter")
            return
        snapshot = (result or {}).get("snapshot")
        if snapshot:
            if str(snapshot.get("document_key") or "") != self._document_key:
                self._replace_with_snapshot(
                    snapshot, "Target document updated.")
                return
            self._apply_snapshot_in_place(
                snapshot,
                "Device data refreshed; current circuit groups were preserved.",
            )
        else:
            self._set_action_status("Refresh completed without a document snapshot.")

    def bulk_set_selected_clicked(self, sender, args):
        # Circuit rows are represented by selected GroupVM headers, not by
        # DataGrid.SelectedItems (which contains element/member rows).
        groups = self._selected_groups()
        if not groups:
            forms.alert("Select one or more circuit rows (group headers) first.",
                        title="Bulk Set Selected")
            return

        if self.BulkValuesPopup is None:
            self._set_action_status("Bulk Set controls are unavailable.")
            return
        self._bulk_target_groups = list(groups)
        self.bulk_popup_reset_clicked(None, None)
        self.BulkValuesPopup.IsOpen = True
        try:
            self.BulkPopupLoadName.Focus()
        except Exception:
            pass

    @staticmethod
    def _bulk_popup_text(control):
        try:
            return str(control.Text or "").strip()
        except Exception:
            return ""

    def bulk_popup_reset_clicked(self, sender, args):
        self.BulkPopupLoadName.Text = ""
        self.BulkPopupPanel.SelectedIndex = -1
        self.BulkPopupPanel.Text = ""
        self.BulkPopupRating.SelectedIndex = -1
        self.BulkPopupRating.Text = ""
        self.BulkPopupNotes.Text = ""
        self.BulkPopupIncrement.IsChecked = False

    def bulk_popup_cancel_clicked(self, sender, args):
        self.BulkValuesPopup.IsOpen = False
        self._bulk_target_groups = []

    def bulk_popup_apply_clicked(self, sender, args):
        values = {
            "load_name": self._bulk_popup_text(self.BulkPopupLoadName),
            "panel": self._bulk_popup_text(self.BulkPopupPanel),
            "rating": self._bulk_popup_text(self.BulkPopupRating),
            "schedule_notes": self._bulk_popup_text(self.BulkPopupNotes),
            "increment_load_name": bool(self.BulkPopupIncrement.IsChecked),
        }
        groups = list(self._bulk_target_groups)
        self.BulkValuesPopup.IsOpen = False
        self._bulk_target_groups = []
        if groups:
            self._apply_bulk_values(groups, values)

    def _apply_bulk_values(self, groups, values):

        load_name = values.get("load_name", "")
        panel = values.get("panel", "")
        rating = values.get("rating", "")
        notes = values.get("schedule_notes", "")
        increment = bool(values.get("increment_load_name", False))
        members_by_group = self._members_map()

        self._begin_view_batch()
        try:
            for index, group in enumerate(groups, 1):
                if load_name:
                    group.load_name = (
                        "{} {}".format(load_name, index)
                        if increment else load_name
                    )
                    group.notify("load_name")
                if panel:
                    compatible = self._compatible_panel_names(
                        members_by_group.get(group, []))
                    if panel in compatible:
                        group.panel = panel
                        group.panel_note = ""
                    else:
                        group.panel = ""
                        group.panel_note = (
                            "Bulk panel '{}' is incompatible with the group's "
                            "voltage/poles and was cleared."
                        ).format(panel)
                    group.notify("panel")
                    group.notify("panel_note")
                if rating:
                    group.rating = rating
                    group.notify("rating")
                if notes:
                    group.schedule_notes = notes
                    group.notify("schedule_notes")
            self._refresh_group_panel_options(members_by_group)
            self._revalidate()
        finally:
            self._end_view_batch()

    def create_dedicated_clicked(self, sender, args):
        selected = self._selected_rows()
        if not selected:
            forms.alert("Select one or more member rows first.",
                        title="Create Dedicated")
            return

        existing_keys = self._all_group_keys()
        self._begin_view_batch()
        try:
            # Each selected element gets its own GroupVM/circuit row. Capture
            # the source group before moving the row so all circuit inputs are
            # copied exactly, even when several selected members share a group.
            for row in selected:
                source = row.group
                if source is None:
                    continue
                stem = str(source.key or "Group").strip() or "Group"
                key = "{} - DEDICATED".format(stem)
                suffix = 2
                while key in existing_keys:
                    key = "{} - DEDICATED {}".format(stem, suffix)
                    suffix += 1
                existing_keys.add(key)

                dedicated = GroupVM(key, status_brushes=self._status_brushes)
                # Preserve every circuit input from the source group. The key
                # contains DEDICATED, so _build_plans will make one native
                # circuit per selected member while retaining these inputs.
                dedicated.load_name = source.load_name
                dedicated.panel = source.panel
                dedicated.rating = source.rating
                dedicated.schedule_notes = source.schedule_notes
                dedicated.panel_note = source.panel_note
                dedicated.set_panel_options(list(source.panel_options or []))
                self._groups.append(dedicated)
                row.move_to(dedicated)

            self._reindex_members()
            self._refresh_view()
            self._revalidate()
        finally:
            self._end_view_batch()

    def new_group_clicked(self, sender, args):
        selected = self._selected_rows()
        if not selected:
            forms.alert("Select one or more rows first.", title="Create Circuits by Device Parameter")
            return
        new_key = cg_core.next_new_group_key(self._all_group_keys())
        g = GroupVM(new_key, status_brushes=self._status_brushes)
        self._init_group_defaults(g, members=selected)
        self._groups.append(g)
        self._begin_view_batch()
        try:
            for vm in selected:
                vm.move_to(g)
            self._reindex_members()
            self._refresh_view()
            self._revalidate()
        finally:
            self._end_view_batch()

    # -- drag and drop ----------------------------------------------------
    def _find_row_item(self, source):
        node = source
        while node is not None:
            if isinstance(node, DataGridRow):
                return node.Item
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                return None
        return None

    def _find_control(self, source, control_type):
        node = source
        while node is not None:
            if isinstance(node, control_type):
                return node
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                return None
        return None

    def _set_drag_target(self, group):
        if self._drag_target is group:
            return
        if self._drag_target is not None:
            self._drag_target.set_drag_target(False)
        self._drag_target = group
        if self._drag_target is not None:
            self._drag_target.set_drag_target(True)

    def grid_mouse_down(self, sender, args):
        self._drag_start = None
        self._drag_items = []
        self._drag_started = False
        self._drag_allowed = False
        self._mouse_down_item = None
        self._mouse_down_selection = []
        self._defer_selection = False
        self._pending_checkbox_selection = []
        self._checkbox_selection_restore = []

        # Group headers have their own selection model and handler. Do not let
        # the DataGrid-level handler clear that state before the header sees it.
        if self._find_group_from_source(args.OriginalSource) is not None:
            return

        item = self._find_row_item(args.OriginalSource)
        if item is None:
            self._clear_grid_selection()
            self._clear_group_selection()
            self._grid_anchor_item = None
            self._last_group_selection = None
            self._pending_checkbox_selection = []
            return

        current_selection = list(self.Grid.SelectedItems or [])
        checkbox = self._find_control(args.OriginalSource, CheckBox)
        if checkbox is not None:
            # Preserve a multi-row selection while a selected row's checkbox is
            # clicked. row_include_changed uses this snapshot for bulk toggle.
            if (
                len(current_selection) > 1
                and item in current_selection
                and not bool(Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift))
            ):
                saved = list(current_selection)
                self._pending_checkbox_selection = saved
                self._checkbox_selection_restore = saved
            return

        self._clear_group_selection()
        self._last_group_selection = None
        self._mouse_down_item = item
        self._mouse_down_selection = current_selection
        modifiers = Keyboard.Modifiers
        ctrl_down = bool(modifiers & ModifierKeys.Control)
        shift_down = bool(modifiers & ModifierKeys.Shift)

        # DataGrid's native range calculation is unreliable once the view has
        # nested GroupStyles. Use the exact row order used by the view and
        # leave native handling for clicks that do not have a valid anchor.
        if shift_down and self._grid_anchor_item is not None:
            ordered_rows = self._ordered_visible_rows()
            if self._grid_anchor_item in ordered_rows and item in ordered_rows:
                anchor_index = ordered_rows.index(self._grid_anchor_item)
                current_index = ordered_rows.index(item)
                lo, hi = sorted((anchor_index, current_index))
                if not ctrl_down:
                    self._clear_grid_selection()
                for target in ordered_rows[lo:hi + 1]:
                    try:
                        if target not in self.Grid.SelectedItems:
                            self.Grid.SelectedItems.Add(target)
                    except Exception:
                        pass
                self._update_grid_selection_summary()
                args.Handled = True
                return

        self._grid_anchor_item = item
        self._drag_allowed = not bool(
            modifiers & (ModifierKeys.Control | ModifierKeys.Shift))
        try:
            self._drag_start = args.GetPosition(self.Grid)
        except Exception:
            self._drag_start = None

        # Only defer native selection when the mouse-down row is already in a
        # multi-row selection. Clicking a different row must remain a normal
        # DataGrid click, so a subsequent drag cannot carry stale selection.
        if (
            self._drag_allowed
            and len(current_selection) > 1
            and item in current_selection
        ):
            self._defer_selection = True
            args.Handled = True

    def grid_mouse_move(self, sender, args):
        if self._results_mode:
            return
        if (
            args.LeftButton != MouseButtonState.Pressed
            or self._drag_start is None
            or not self._drag_allowed
        ):
            return
        try:
            cur = args.GetPosition(self.Grid)
            if Math.Abs(cur.X - self._drag_start.X) < 4 and Math.Abs(cur.Y - self._drag_start.Y) < 4:
                return
        except Exception:
            return
        if self._defer_selection:
            selected = list(self._mouse_down_selection)
        elif (
            self._mouse_down_item is not None
            and self._mouse_down_item not in self._mouse_down_selection
        ):
            # The native DataGrid selection event may not have completed yet
            # when MouseMove first fires. The clicked row is still the only
            # valid drag source in this case.
            selected = [self._mouse_down_item]
        else:
            selected = list(self.Grid.SelectedItems or [])
        if not selected and self._mouse_down_item is not None:
            selected = [self._mouse_down_item]
        if not selected:
            return
        self._drag_items = selected
        self._drag_started = True
        try:
            data = DataObject("CED.CircuitGrouperRows", "rows")
            DragDrop.DoDragDrop(self.Grid, data, DragDropEffects.Move)
        except Exception:
            pass
        self._drag_start = None
        self._drag_allowed = False
        self._drag_items = []

    def grid_mouse_up(self, sender, args):
        if (
            self._defer_selection
            and not self._drag_started
            and self._mouse_down_item is not None
        ):
            self._clear_grid_selection()
            try:
                self.Grid.SelectedItems.Add(self._mouse_down_item)
                self.Grid.SelectedItem = self._mouse_down_item
            except Exception:
                pass

        if self._checkbox_selection_restore:
            saved = list(self._checkbox_selection_restore)
            self._clear_grid_selection()
            for item in saved:
                try:
                    self.Grid.SelectedItems.Add(item)
                except Exception:
                    pass

        self._checkbox_selection_restore = []
        self._set_drag_target(None)
        self._drag_start = None
        self._drag_items = []
        self._drag_started = False
        self._drag_allowed = False
        self._mouse_down_item = None
        self._mouse_down_selection = []
        self._defer_selection = False

    def grid_drag_over(self, sender, args):
        if self._results_mode:
            args.Effects = getattr(DragDropEffects, "None")
            args.Handled = True
            return
        if not self._drag_items:
            self._set_drag_target(None)
            args.Effects = getattr(DragDropEffects, "None")
            return
        target = self._find_drop_group_from_source(args.OriginalSource)
        self._set_drag_target(target)
        args.Effects = DragDropEffects.Move if target is not None else getattr(
            DragDropEffects, "None")
        args.Handled = True

    def grid_drop(self, sender, args):
        if self._results_mode:
            return
        if not self._drag_items:
            return
        dest = self._find_drop_group_from_source(args.OriginalSource)
        if dest is None:
            return
        self._begin_view_batch()
        try:
            for vm in self._drag_items:
                vm.move_to(dest)
            self._reindex_members()
            self._refresh_view()
            self._revalidate()
        finally:
            self._set_drag_target(None)
            self._end_view_batch()

    # -- finish -----------------------------------------------------------
    def _build_plans(self):
        members_by_group = self._members_by_group
        ok, blocked = [], []
        for g in self._groups:
            eff = cg_core.effective_rows(members_by_group.get(g, []))
            if not eff:
                continue  # all already circuited / excluded
            is_dedicated = (
                cg_core.is_dedicated_group_name(g.key) or
                cg_core.is_dedicated_group_name(g.load_name)
            )
            if is_dedicated:
                # The group remains useful for editing Panel/Breaker values,
                # but each effective member becomes its own native circuit.
                for member in eff:
                    member_label = (str(member.identity_mark or "").strip() or
                                    str(member.element_id))
                    plan_key = "{} / {}".format(g.key, member_label)
                    plan = cg_core.build_group_plan(
                        plan_key, g.load_name or g.key, g.panel, g.rating, [member],
                        g.schedule_notes)
                    plan["dedicated"] = True
                    plan["source_group_key"] = g.key
                    plan["dedicated_element_id"] = int(member.element_id)
                    (ok if plan["ready"] else blocked).append(plan)
            else:
                plan = cg_core.build_group_plan(
                    g.key, g.load_name, g.panel, g.rating, eff,
                    g.schedule_notes)
                (ok if plan["ready"] else blocked).append(plan)
        return ok, blocked

    def run_clicked(self, sender, args):
        if self._run_in_progress:
            return
        if self._gateway is not None and self._gateway.is_busy():
            self._set_action_status("Another Revit action is still pending.")
            return

        ok_plans, blocked = self._build_plans()
        if not ok_plans and not blocked:
            forms.alert("Nothing to circuit (all items are already circuited or excluded).",
                        title="Create Circuits by Device Parameter")
            return

        if blocked:
            total = len(ok_plans) + len(blocked)
            lines = [
                "{} of {} circuit(s) are not ready and will be skipped.".format(
                    len(blocked), total),
            ]
            for p in blocked:
                lines.append(u"  - {}: {}".format(
                    p["load_name"], "; ".join(p["problems"])))
            lines.append("")
            lines.append("Continue with the ready circuit(s)?")
            if not forms.alert("\n".join(lines), title="Create Circuits by Device Parameter", yes=True, no=True):
                return
            if not ok_plans:
                forms.alert("No circuit(s) are ready to run.", title="Create Circuits by Device Parameter")
                return

        if self._gateway is None:
            # Compatibility fallback for the legacy modal entry point.
            self.result_plans = ok_plans
            self.DialogResult = True
            self.Close()
            return

        self._run_in_progress = True
        self._set_action_status("Creating {} circuit(s)...".format(len(ok_plans)))
        if self.RunButton is not None:
            self.RunButton.IsEnabled = False
        if self.RefreshButton is not None:
            self.RefreshButton.IsEnabled = False
        raised = self._gateway.raise_run(
            self._document_key, ok_plans, self._scope_ids,
            self._handle_run_complete)
        if not raised:
            self._run_in_progress = False
            if self.RunButton is not None:
                self.RunButton.IsEnabled = True
            if self.RefreshButton is not None:
                self.RefreshButton.IsEnabled = True
            self._set_action_status("Another Revit action is still pending.")
            return

        # The established panel-assignment service can show native Revit
        # confirmations while the ExternalEvent is running. A floating
        # modeless WPF window can cover those dialogs and make Revit appear to
        # hang. Hide only for the duration of Run; completion is marshalled
        # back to the dispatcher after the ExternalEvent releases Revit.
        try:
            self.Hide()
        except Exception:
            pass

    def _handle_run_complete(self, status, op_name, result, error):
        """Marshal completion out of the Revit ExternalEvent callback."""
        try:
            self.Dispatcher.BeginInvoke(
                Action(lambda: self._finish_run_complete(status, result, error)),
                DispatcherPriority.Background,
            )
        except Exception:
            self._finish_run_complete(status, result, error)

    @staticmethod
    def _result_id_set(values):
        result = set()
        for value in list(values or []):
            try:
                result.add(int(value))
            except Exception:
                continue
        return result

    @staticmethod
    def _assignment_label(assignment_rows):
        numbers_by_panel = {}
        panel_order = []
        for row in list(assignment_rows or []):
            panel = str(row.get("actual_panel") or "(unassigned)")
            if panel.strip().lower() in ("", "(none)", "none"):
                panel = "(Unassigned)"
            number = str(row.get("circuit_number") or "-")
            if panel not in numbers_by_panel:
                numbers_by_panel[panel] = []
                panel_order.append(panel)
            if number not in numbers_by_panel[panel]:
                numbers_by_panel[panel].append(number)
        return "; ".join(
            u"{} – {}".format(panel, ", ".join(numbers_by_panel[panel]))
            for panel in panel_order
        )

    def _lock_results_controls(self):
        self._results_mode = True
        if self.BulkValuesPopup is not None:
            self.BulkValuesPopup.IsOpen = False
        for control in (
                self.NewGroupButton, self.DedicatedGroupButton,
                self.BulkSetButton, self.GroupByCombo, self.NameByCombo,
                self.RefreshButton, self.RunButton):
            if control is not None:
                control.IsEnabled = False
        if self.Grid is not None:
            self.Grid.AllowDrop = False
        for row in list(self._rows):
            row.lock_results()

    def _enter_results_mode(self, report):
        report = dict(report or {})
        assignment = dict(report.get("panel_assignment") or {})
        assignment_rows = list(assignment.get("circuit_status") or [])
        assignment_by_id = {}
        for row in assignment_rows:
            try:
                assignment_by_id[int(row.get("circuit_id"))] = row
            except Exception:
                continue

        circuit_results = list(report.get("circuit_results") or [])
        results_by_group = {}
        result_by_member = {}
        for circuit_result in circuit_results:
            group_key = str(
                circuit_result.get("source_group_key") or
                circuit_result.get("group_key") or "")
            results_by_group.setdefault(group_key, []).append(circuit_result)
            for element_id in self._result_id_set(
                    circuit_result.get("member_element_ids")):
                result_by_member[element_id] = circuit_result

        failures_by_group = {}
        failed_member_ids = set()
        for failure in list(report.get("circuit_failures") or []):
            group_key = str(
                failure.get("source_group_key") or
                failure.get("group_key") or "")
            failures_by_group.setdefault(group_key, []).append(failure)
            failed_member_ids.update(self._result_id_set(
                failure.get("member_element_ids")))

        skipped_no_connector = self._result_id_set(
            report.get("skipped_no_connector"))
        skipped_in_use = self._result_id_set(
            report.get("skipped_unavailable_primary"))
        assignment_errors_by_panel = {}
        for panel_name, message in list(assignment.get("errors") or []):
            assignment_errors_by_panel.setdefault(
                str(panel_name or ""), []).append(str(message or ""))
        attempted_groups = set(
            str(value or "") for value in list(
                report.get("attempted_source_group_keys") or []))

        self._lock_results_controls()
        for row in list(self._rows):
            element_id = int(row.element_id)
            circuit_result = result_by_member.get(element_id)
            if circuit_result is not None:
                row.set_already_circuited(True)
                assignment_row = assignment_by_id.get(
                    int(circuit_result.get("circuit_id") or 0))
                if assignment_row and assignment_row.get("status") == "ASSIGNED":
                    row.set_status(
                        "Circuited: {}".format(
                            self._assignment_label([assignment_row])),
                        self._status_brushes["ready"],
                    )
                else:
                    row.set_status(
                        "Circuit created; panel assignment needs attention",
                        self._status_brushes["warn"],
                    )
            elif element_id in skipped_no_connector:
                row.set_status(
                    "Skipped: no primary power connector",
                    self._status_brushes["warn"],
                )
            elif element_id in skipped_in_use:
                row.set_status(
                    "Skipped: primary power connector already in use",
                    self._status_brushes["warn"],
                )
            elif element_id in failed_member_ids:
                row.set_status(
                    "Circuit creation failed",
                    self._status_brushes["bad"],
                )
            elif row.already_circuited:
                row.set_status(
                    "Already circuited before Run",
                    self._status_brushes["info"],
                )
            else:
                row.set_status("Excluded", self._status_brushes["off"])

        attention_groups = 0
        failed_groups = 0
        result_lines = []
        for group in list(self._groups):
            group_results = list(results_by_group.get(group.key) or [])
            group_failures = list(failures_by_group.get(group.key) or [])
            group_assignment_rows = []
            missing_assignment = False
            for circuit_result in group_results:
                assignment_row = assignment_by_id.get(
                    int(circuit_result.get("circuit_id") or 0))
                if assignment_row is None:
                    missing_assignment = True
                else:
                    group_assignment_rows.append(assignment_row)

            member_ids = self._result_id_set(
                row.element_id for row in self._members_by_group.get(group, []))
            skipped_count = len(member_ids.intersection(
                skipped_no_connector.union(skipped_in_use)))
            assignment_attention = bool(
                missing_assignment or any(
                    row.get("status") != "ASSIGNED"
                    for row in group_assignment_rows))
            panel_errors = []
            for circuit_result in group_results:
                for message in assignment_errors_by_panel.get(
                        str(circuit_result.get("target_panel") or ""), []):
                    if message not in panel_errors:
                        panel_errors.append(message)

            detail_lines = []
            for assignment_row in group_assignment_rows:
                detail_lines.append(
                    u"{} {}: {}".format(
                        u"✓" if assignment_row.get("status") == "ASSIGNED" else u"⚠",
                        self._assignment_label([assignment_row]),
                        assignment_row.get("detail") or assignment_row.get("status"),
                    ))
            for failure in group_failures:
                detail_lines.append(
                    u"✕ {}".format(failure.get("detail") or
                                   "Circuit creation failed."))
            for message in panel_errors:
                if message:
                    detail_lines.append(u"⚠ {}".format(message))
            if skipped_count:
                detail_lines.append(
                    u"⚠ {} selected member(s) were skipped.".format(
                        skipped_count))

            assignment_label = self._assignment_label(group_assignment_rows)
            if not assignment_label and group_results:
                assignment_label = "Created – assignment unavailable"
            if not group_results and group.key not in attempted_groups:
                group.set_result_status(
                    "neutral", "Not run",
                    "This group had no eligible included members when Run started.",
                    self._status_brushes,
                )
            elif not group_results:
                failed_groups += 1
                group.set_result_status(
                    "error", "No circuit created",
                    "\n".join(detail_lines) or "No circuit was created for this group.",
                    self._status_brushes,
                )
                result_lines.append(
                    u"{}: no circuit created.".format(group.load_name or group.key))
            elif assignment_attention or group_failures or skipped_count or panel_errors:
                attention_groups += 1
                group.set_result_status(
                    "warning", assignment_label,
                    "\n".join(detail_lines) or
                    "Circuit was created but needs attention.",
                    self._status_brushes,
                )
                result_lines.append(
                    u"{}: created with attention.".format(
                        group.load_name or group.key))
            else:
                group.set_result_status(
                    "success", assignment_label,
                    "\n".join(detail_lines) or
                    "Circuit was created and assigned as intended.",
                    self._status_brushes,
                )

        created = int(report.get("created") or 0)
        assigned = sum(
            1 for row in assignment_rows if row.get("status") == "ASSIGNED")
        summary_parts = ["{} created".format(created),
                         "{} assigned as intended".format(assigned)]
        if attention_groups:
            summary_parts.append("{} need attention".format(attention_groups))
        if failed_groups:
            summary_parts.append("{} failed".format(failed_groups))
        if assignment.get("fallback_used"):
            result_lines.insert(
                0,
                "Panel capacity workflow used removable SPARE/SPACE rows.",
            )
        self.SummaryText.Text = " | ".join(summary_parts)
        self.ValidationText.Text = (
            "\n".join(result_lines)
            if result_lines else
            "All created circuits were assigned as intended.")

    def _finish_run_complete(self, status, result, error):
        self._run_in_progress = False
        if self.RunButton is not None:
            self.RunButton.IsEnabled = True
        if self.RefreshButton is not None:
            self.RefreshButton.IsEnabled = True

        if status != "ok":
            message = str(error or "Circuiting failed.")
            try:
                self.Show()
                self.Activate()
            except Exception:
                pass
            self._set_action_status(message)
            forms.alert(message, title="Create Circuits by Device Parameter")
            return

        report = result or {}
        snapshot = report.get("snapshot") or report.get("retarget_snapshot")
        created = int(report.get("created") or 0)
        assignment = dict(report.get("panel_assignment") or {})
        not_on_target = int(assignment.get("not_on_target") or 0)
        assignment_errors = list(assignment.get("errors") or [])
        if report.get("retarget_snapshot"):
            completion_message = "The active document changed; no circuits were created."
        elif not_on_target or assignment_errors:
            completion_message = (
                "Created {} circuit(s); {} require panel-assignment attention."
                .format(created, max(not_on_target, len(assignment_errors))))
        else:
            completion_message = "Created {} circuit(s).".format(created)

        if snapshot:
            self._replace_with_snapshot(snapshot, completion_message)
            return

        try:
            self.Show()
            self.Activate()
        except Exception:
            pass
        self._enter_results_mode(report)
        self._set_action_status(completion_message)

    def cancel_clicked(self, sender, args):
        self.Close()


def show_window(rows_data, panel_options, name_to_id, rating_options,
                group_param_options=None, default_group_param="",
                panel_info=None, name_param_options=None,
                default_name_param=""):
    win = CircuitGrouperWindow(
        rows_data, panel_options, name_to_id, rating_options,
        group_param_options, default_group_param, panel_info=panel_info,
        name_param_options=name_param_options,
        default_name_param=default_name_param)
    win.show_dialog()
    return win.result_plans, win._name_to_id


def show_modeless(snapshot, gateway, activate=True):
    """Open a modeless grouper window from a document snapshot."""
    data = dict(snapshot or {})
    rows_data = list(data.get("rows_data") or [])
    win = CircuitGrouperWindow(
        rows_data,
        list(data.get("panel_names") or []),
        dict(data.get("name_to_id") or {}),
        list(data.get("rating_options") or cg_core.RATING_OPTIONS),
        list(data.get("group_param_options") or []),
        data.get("default_group_param") or "",
        panel_info=dict(data.get("panel_info") or {}),
        name_param_options=list(data.get("name_param_options") or []),
        default_name_param=data.get("default_name_param") or "",
        gateway=gateway,
        document_key=data.get("document_key") or "",
        document_title=data.get("document_title") or "-",
        scope_label=data.get("scope_label") or "model",
        scope_ids=data.get("scope_ids"),
    )
    win.Tag = "_ae_circuit_grouper_window_persistent_v6"
    if data.get("empty_message"):
        win._set_action_status(data.get("empty_message"))
    win.Show()
    if activate:
        try:
            win.Activate()
        except Exception:
            pass
    return win
