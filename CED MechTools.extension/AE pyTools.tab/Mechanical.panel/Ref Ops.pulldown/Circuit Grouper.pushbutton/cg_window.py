# -*- coding: utf-8 -*-
"""Circuit Grouper - WPF window controller (IronPython / .NET).

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
    except Exception:
        pass

from System import Math
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
from System.Windows.Controls import DataGridRow
from System.Windows.Data import CollectionViewSource, PropertyGroupDescription
from System.Windows.Input import MouseButtonState
from System.Windows.Media import BrushConverter, VisualTreeHelper

from pyrevit import forms, revit, DB

import cg_core

THIS_DIR = os.path.dirname(__file__)
_XAML = os.path.join(THIS_DIR, "CircuitGrouperWindow.xaml")

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
            try:
                h(self, args)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# View-models
# ---------------------------------------------------------------------------
class GroupVM(_Notifier):
    def __init__(self, key):
        _Notifier.__init__(self)
        self.key = key
        self.load_name = key
        self.panel = ""
        self.rating = cg_core.DEFAULT_RATING
        self.status = "Ready"
        self.status_color = BRUSH_READY
        self.rating_warning_visibility = Visibility.Collapsed
        self.rating_tooltip = NONSTANDARD_TOOLTIP

    def set_status(self, text, color):
        self.status = text
        self.status_color = color
        self.notify("status")
        self.notify("status_color")

    def set_warning(self, show):
        self.rating_warning_visibility = Visibility.Visible if show else Visibility.Collapsed
        self.notify("rating_warning_visibility")


class RowVM(_Notifier):
    def __init__(self, data, group):
        _Notifier.__init__(self)
        self.element_id = data["element_id"]
        self.family_type = data.get("family_type", "")
        self.identity_mark = data.get("identity_mark", "")
        self.voltage_text = data.get("voltage_text", "") or ""
        self.voltage_key = data.get("voltage_key", None)
        self.poles_text = data.get("poles_text", "") or ""
        self.poles_value = data.get("poles_value", None)
        self.already_circuited = bool(data.get("already_circuited", False))
        self.src_panel = (data.get("panel", "") or "").strip()
        self.src_rating = (data.get("rating", "") or "").strip()

        self.group = group
        self.group_key = group.key
        self.can_include = not self.already_circuited
        self.include = not self.already_circuited
        self.status = ""
        self.status_color = BRUSH_READY

    def set_status(self, text, color):
        self.status = text
        self.status_color = color
        self.notify("status")
        self.notify("status_color")

    def move_to(self, group):
        self.group = group
        self.group_key = group.key
        self.notify("group")
        self.notify("group_key")


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
class CircuitGrouperWindow(forms.WPFWindow):
    def __init__(self, rows_data, panel_options, name_to_id, rating_options, title=""):
        forms.WPFWindow.__init__(self, _XAML)

        self.PanelOptions = List[str]()
        for p in panel_options:
            self.PanelOptions.Add(p)
        self.RatingOptions = List[str]()
        for r in rating_options:
            self.RatingOptions.Add(r)

        self._name_to_id = name_to_id
        self.result_plans = None
        self._drag_start = None
        self._drag_items = []

        # build groups + rows (grouped by circuit number)
        self._rows = []
        self._groups = []
        groups_by_key = {}
        for d in rows_data:
            key = d.get("circuit_number", "") or ""
            g = groups_by_key.get(key)
            if g is None:
                g = GroupVM(key)
                groups_by_key[key] = g
                self._groups.append(g)
            self._rows.append(RowVM(d, g))
        for g in self._groups:
            self._init_group_defaults(g)

        self.DataContext = self

        self.Grid = self.FindName("Grid")
        self.SummaryText = self.FindName("SummaryText")
        self.ValidationText = self.FindName("ValidationText")

        self._items = ObservableCollection[object]()
        for vm in self._rows:
            self._items.Add(vm)

        self._cvs = CollectionViewSource()
        self._cvs.Source = self._items
        self._cvs.GroupDescriptions.Add(PropertyGroupDescription("group"))
        if SortDescription is not None and ListSortDirection is not None:
            self._cvs.SortDescriptions.Add(
                SortDescription("group_key", ListSortDirection.Ascending)
            )
        self._view = self._cvs.View
        self.Grid.ItemsSource = self._view

        self._revalidate()

    # -- setup helpers ----------------------------------------------------
    def _members(self, group):
        return [r for r in self._rows if r.group is group]

    def _init_group_defaults(self, group):
        members = self._members(group)
        group.load_name = cg_core.default_load_name(
            [r.identity_mark for r in members], fallback=group.key)
        for r in members:
            if r.src_panel:
                group.panel = r.src_panel
                break
        for r in members:
            amps, valid, _ = cg_core.parse_rating(r.src_rating)
            if valid:
                group.rating = cg_core.format_amps_number(amps)
                break

    def _all_group_keys(self):
        return set(g.key for g in self._groups)

    def _combo_value(self, combo):
        item = combo.SelectedItem
        if item is not None:
            return str(item)
        return combo.Text or ""

    def _refresh_view(self):
        try:
            self._view.Refresh()
        except Exception:
            pass

    # -- validation -------------------------------------------------------
    def _revalidate(self):
        members_by_group = {}
        for r in self._rows:
            members_by_group.setdefault(r.group, []).append(r)
        self._groups = [g for g in self._groups if g in members_by_group]

        ready = 0
        not_ready = 0
        lines = []

        for g in self._groups:
            members = members_by_group.get(g, [])
            # per-row status
            for r in members:
                if r.already_circuited:
                    r.set_status("Already circuited", BRUSH_INFO)
                elif not r.include:
                    r.set_status("Excluded", BRUSH_OFF)
                else:
                    r.set_status("", BRUSH_READY)

            eff = cg_core.effective_rows(members)
            if not eff:
                if members and all(r.already_circuited for r in members):
                    g.set_status("Already circuited", BRUSH_INFO)
                else:
                    g.set_status("No items selected", BRUSH_OFF)
                g.set_warning(False)
                continue

            problems = cg_core.validate_members(eff)
            amps, valid, standard = cg_core.parse_rating(g.rating)
            label = g.load_name or g.key
            if problems:
                g.set_status(u"⚠ " + "; ".join(problems), BRUSH_BAD)
                g.set_warning(False)
                not_ready += 1
                lines.append(u"'{}': {}".format(label, "; ".join(problems)))
            elif not valid:
                g.set_status("Invalid rating - not ready", BRUSH_BAD)
                g.set_warning(False)
                not_ready += 1
                lines.append(u"'{}': invalid breaker rating '{}'".format(label, g.rating))
            else:
                g.set_status("Ready", BRUSH_READY)
                g.set_warning(not standard)
                ready += 1
                if not standard:
                    lines.append(u"'{}': non-standard breaker size ({} A)".format(
                        label, cg_core.format_amps_number(amps)))

        self.SummaryText.Text = "{} circuit(s) ready | {} not ready".format(ready, not_ready)
        self.ValidationText.Text = "\n".join(lines) if lines else "No validation issues."

    # -- group header handlers (sender.DataContext == GroupVM) -----------
    def group_load_name_changed(self, sender, args):
        g = sender.DataContext
        if not isinstance(g, GroupVM):
            return
        g.load_name = sender.Text or ""

    def group_panel_changed(self, sender, args):
        g = sender.DataContext
        if not isinstance(g, GroupVM):
            return
        value = self._combo_value(sender)
        if not value or value == g.panel:
            return
        g.panel = value

    def group_rating_changed(self, sender, args):
        g = sender.DataContext
        if not isinstance(g, GroupVM):
            return
        value = (sender.Text or "").strip()
        if not value and sender.SelectedItem is not None:
            value = str(sender.SelectedItem)
        if not value or value == g.rating:
            return
        g.rating = value
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
        vm = sender.DataContext
        if vm is None:
            return
        new_value = bool(sender.IsChecked)
        if vm.include == new_value:
            return
        vm.include = new_value
        self._revalidate()

    # -- toolbar ----------------------------------------------------------
    def new_group_clicked(self, sender, args):
        selected = list(self.Grid.SelectedItems or [])
        if not selected:
            forms.alert("Select one or more rows first.", title="Circuit Grouper")
            return
        new_key = cg_core.next_new_group_key(self._all_group_keys())
        g = GroupVM(new_key)
        g.load_name = cg_core.default_load_name(
            [r.identity_mark for r in selected], fallback=new_key)
        self._groups.append(g)
        for vm in selected:
            vm.move_to(g)
        self._refresh_view()
        self._revalidate()

    def toggle_checked_clicked(self, sender, args):
        targets = list(self.Grid.SelectedItems or [])
        if not targets:
            targets = list(self._rows)
        targets = [vm for vm in targets if vm.can_include]
        if not targets:
            return
        new_value = not all(vm.include for vm in targets)
        for vm in targets:
            vm.include = new_value
        self._revalidate()
        self._refresh_view()

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

    def grid_mouse_down(self, sender, args):
        try:
            self._drag_start = args.GetPosition(self.Grid)
        except Exception:
            self._drag_start = None

    def grid_mouse_move(self, sender, args):
        if args.LeftButton != MouseButtonState.Pressed or self._drag_start is None:
            return
        try:
            cur = args.GetPosition(self.Grid)
            if Math.Abs(cur.X - self._drag_start.X) < 4 and Math.Abs(cur.Y - self._drag_start.Y) < 4:
                return
        except Exception:
            return
        selected = list(self.Grid.SelectedItems or [])
        if not selected:
            return
        self._drag_items = selected
        try:
            data = DataObject("CED.CircuitGrouperRows", "rows")
            DragDrop.DoDragDrop(self.Grid, data, DragDropEffects.Move)
        except Exception:
            pass
        self._drag_start = None
        self._drag_items = []

    def grid_drag_over(self, sender, args):
        args.Effects = DragDropEffects.Move if self._drag_items else getattr(DragDropEffects, "None")

    def grid_drop(self, sender, args):
        if not self._drag_items:
            return
        target = self._find_row_item(args.OriginalSource)
        if target is None:
            return
        dest = target.group
        for vm in self._drag_items:
            vm.move_to(dest)
        self._refresh_view()
        self._revalidate()

    # -- finish -----------------------------------------------------------
    def _build_plans(self):
        members_by_group = {}
        for r in self._rows:
            members_by_group.setdefault(r.group, []).append(r)
        ok, blocked = [], []
        for g in self._groups:
            eff = cg_core.effective_rows(members_by_group.get(g, []))
            if not eff:
                continue  # all already circuited / excluded
            plan = cg_core.build_group_plan(g.key, g.load_name, g.panel, g.rating, eff)
            (ok if plan["ready"] else blocked).append(plan)
        return ok, blocked

    def run_clicked(self, sender, args):
        ok_plans, blocked = self._build_plans()
        if not ok_plans and not blocked:
            forms.alert("Nothing to circuit (all items are already circuited or excluded).",
                        title="Circuit Grouper")
            return

        if blocked:
            lines = ["{} group(s) are not ready and will be SKIPPED:".format(len(blocked))]
            for p in blocked:
                lines.append(u"  - {}: {}".format(
                    p["load_name"], "; ".join(p["problems"])))
            if ok_plans:
                lines.append("")
                lines.append("Proceed and circuit the {} ready group(s)?".format(len(ok_plans)))
                if not forms.alert("\n".join(lines), title="Circuit Grouper", yes=True, no=True):
                    return
            else:
                forms.alert("\n".join(lines), title="Circuit Grouper")
                return

        self.result_plans = ok_plans
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


def show_window(rows_data, panel_options, name_to_id, rating_options):
    win = CircuitGrouperWindow(rows_data, panel_options, name_to_id, rating_options)
    win.show_dialog()
    return win.result_plans, win._name_to_id
