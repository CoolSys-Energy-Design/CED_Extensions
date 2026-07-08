# -*- coding: utf-8 -*-

import os

import clr

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Action, TimeSpan
from System.ComponentModel import ListSortDirection
from System.Collections.Generic import List
from System.IO import File
from System.Text import Encoding
from System.Windows import Application, Visibility, WindowState
from System.Windows.Controls import ContextMenu, MenuItem
from System.Windows.Threading import DispatcherPriority, DispatcherTimer
from pyrevit import DB, forms, revit, script

from Snippets import revit_helpers
from UIClasses import pathing as ui_pathing

TITLE = "Electrical QC Check"
WINDOW_MARKER = "_ae_electrical_qc_window_persistent_v1"

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
if not LIB_ROOT or not os.path.isdir(LIB_ROOT):
    forms.alert("Could not locate workspace root for Electrical QC Check.", title=TITLE)
    raise SystemExit

from CEDElectrical.Application.services.electrical_qc_service import rows_to_csv
from CEDElectrical.Application.services.electrical_qc_service import scan_document
from UIClasses.ui_bases import CEDWindowBase

LOGGER = script.get_logger()


def _active_doc():
    doc = getattr(revit, "doc", None)
    if doc is not None:
        return doc
    try:
        uidoc = __revit__.ActiveUIDocument
        return uidoc.Document if uidoc else None
    except Exception:
        return getattr(revit, "doc", None)


def _doc_title(doc):
    try:
        return doc.Title if doc is not None else "-"
    except Exception:
        return "-"


def _doc_key(doc):
    if doc is None:
        return ""
    parts = []
    for attr in ("PathName", "Title"):
        try:
            parts.append(str(getattr(doc, attr, "") or ""))
        except Exception:
            parts.append("")
    try:
        info = getattr(doc, "ProjectInformation", None)
        parts.append(str(getattr(info, "UniqueId", "") or ""))
    except Exception:
        parts.append("")
    try:
        parts.append(str(doc.GetHashCode()))
    except Exception:
        pass
    return "|".join(parts)


def _blank_snapshot(doc=None, status="Click Refresh to check active document."):
    return {
        "status": status,
        "doc_title": _doc_title(doc),
        "doc_key": _doc_key(doc),
        "rows": [],
        "count": 0,
    }


class ElectricalQCGateway(object):
    def __init__(self, logger=None):
        self._logger = logger
        self._pending = None
        self._scan_func = scan_document
        self._elementid_from_value = revit_helpers.elementid_from_value
        self._element_id_list_type = List[DB.ElementId]
        self._load_va_pattern = r"([0-9][0-9,]*(?:\.[0-9]+)?)\s*(KVA|VA)\b"
        try:
            from CEDElectrical.refdata.standard_ocp_table import BREAKER_FRAME_SWITCH_TABLE

            self._standard_breaker_sizes = sorted([int(k) for k in BREAKER_FRAME_SWITCH_TABLE.keys()])
        except Exception:
            self._standard_breaker_sizes = []
        self._handler = _ElectricalQCExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def _is_event_pending(self):
        try:
            return bool(self._event.IsPending)
        except Exception:
            return False

    def is_busy(self):
        return self._pending is not None or self._is_event_pending()

    def _raise(self, op_name, payload=None, callback=None):
        if self.is_busy():
            return False
        self._pending = {
            "op": str(op_name or ""),
            "payload": dict(payload or {}),
            "callback": callback,
        }
        try:
            self._event.Raise()
            return True
        except Exception as ex:
            self._pending = None
            if self._logger:
                self._logger.warning("Electrical QC ExternalEvent raise failed: {}".format(ex))
            return False

    def raise_refresh(self, callback=None):
        return self._raise("refresh", callback=callback)

    def raise_select(self, element_ids, callback=None):
        return self._raise("select", payload={"element_ids": list(element_ids or [])}, callback=callback)

    def _consume_pending(self):
        pending = self._pending
        self._pending = None
        return pending

    def scan_document(self, doc):
        if doc is not None:
            try:
                doc.Regenerate()
            except Exception:
                pass
        snapshot = self._scan_func(
            doc,
            logger=self._logger,
            standard_breaker_sizes=self._standard_breaker_sizes,
            load_va_pattern=self._load_va_pattern,
        )
        try:
            snapshot["doc_key"] = _doc_key(doc)
        except Exception:
            pass
        return snapshot

    def elementid_from_value(self, value):
        return self._elementid_from_value(value)

    def make_element_id_list(self):
        return self._element_id_list_type()


class _ElectricalQCExternalEventHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, application):  # noqa: N802
        pending = self._gateway._consume_pending()
        if not pending:
            return
        op_name = pending.get("op")
        payload = dict(pending.get("payload") or {})
        callback = pending.get("callback")
        status = "ok"
        result = None
        error = None
        try:
            uidoc = None
            try:
                uidoc = application.ActiveUIDocument
            except Exception:
                uidoc = None
            if uidoc is None:
                try:
                    uidoc = __revit__.ActiveUIDocument
                except Exception:
                    uidoc = getattr(revit, "uidoc", None)
            doc = uidoc.Document if uidoc is not None else None
            if op_name == "refresh":
                if doc is None:
                    raise Exception("No active document.")
                try:
                    doc.Regenerate()
                except Exception:
                    pass
                result = self._gateway.scan_document(doc)
            elif op_name == "select":
                if uidoc is None or doc is None:
                    raise Exception("No active document.")
                element_ids = []
                for value in list(payload.get("element_ids") or []):
                    try:
                        numeric = int(value or 0)
                    except Exception:
                        numeric = 0
                    if numeric <= 0:
                        continue
                    element = doc.GetElement(self._gateway.elementid_from_value(numeric))
                    if element is None:
                        continue
                    element_ids.append(self._gateway.elementid_from_value(numeric))
                selection = self._gateway.make_element_id_list()
                for element_id in element_ids:
                    selection.Add(element_id)
                uidoc.Selection.SetElementIds(selection)
                result = {"selected": len(element_ids)}
            else:
                raise Exception("Unknown operation: {}".format(op_name))
        except Exception as ex:
            status = "error"
            error = ex
            if self._gateway._logger:
                try:
                    self._gateway._logger.exception("Electrical QC external operation failed: {}".format(ex))
                except Exception:
                    pass
        if callback:
            try:
                callback(status, op_name, result, error)
            except Exception:
                pass

    def GetName(self):  # noqa: N802
        return "CED Electrical QC External Event"


class ElectricalQCWindow(CEDWindowBase):
    theme_aware = True
    use_config_theme = True

    def __init__(self, snapshot, gateway):
        self._gateway = gateway
        self._rows = []
        self._doc_title = "-"
        self._doc_key = ""
        self._column_menu = None
        self._sort_key = "default"
        self._sort_descending = False
        self._grid_reveal_queued = False
        self._grid_revealed = False
        self._loading_timer = None
        self._loading_complete_timer = None
        self._loading_progress_value = 0
        xaml = os.path.abspath(os.path.join(THIS_DIR, "ElectricalSystemCheckWindow.xaml"))
        CEDWindowBase.__init__(self, xaml_source=xaml, theme_aware=True)
        try:
            self.Tag = WINDOW_MARKER
        except Exception:
            pass
        try:
            setattr(self, WINDOW_MARKER, True)
        except Exception:
            pass

        self._grid = self.FindName("QCGrid")
        self._document_text = self.FindName("DocumentText")
        self._count_text = self.FindName("CountText")
        self._status_text = self.FindName("StatusText")
        self._refresh_button = self.FindName("RefreshButton")
        self._columns_button = self.FindName("ColumnsButton")
        self._export_button = self.FindName("ExportButton")
        self._select_primary_button = self.FindName("SelectPrimaryButton")
        self._select_secondary_button = self.FindName("SelectSecondaryButton")
        self._loading_overlay = self.FindName("LoadingOverlay")
        self._loading_progress = self.FindName("LoadingProgress")

        self._apply_snapshot(snapshot or {})
        self._sync_selection_buttons()

    def _set_status(self, text):
        if self._status_text is None:
            return
        try:
            self._status_text.Text = str(text or "")
        except Exception:
            pass

    def _set_busy(self, is_busy, status_text=None, show_loading=False):
        enabled = not bool(is_busy)
        for button in (self._refresh_button, self._export_button, self._columns_button):
            if button is None:
                continue
            try:
                button.IsEnabled = enabled
            except Exception:
                pass
        if status_text:
            self._set_status(status_text)
        if self._loading_overlay is not None:
            try:
                self._loading_overlay.Visibility = (
                    Visibility.Visible
                    if bool(show_loading)
                    else Visibility.Collapsed
                )
            except Exception:
                pass
        if bool(is_busy) and bool(show_loading):
            self._start_loading_progress()
        elif not bool(is_busy) and not bool(show_loading):
            self._stop_loading_progress(show_complete=False)
        self._sync_selection_buttons()

    def _set_loading_progress_value(self, value):
        self._loading_progress_value = int(value or 0)
        if self._loading_progress is None:
            return
        try:
            self._loading_progress.Value = self._loading_progress_value
        except Exception:
            pass

    def _start_loading_progress(self):
        self._set_loading_progress_value(18)
        if self._loading_timer is None:
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(90)
            timer.Tick += self.loading_progress_tick
            self._loading_timer = timer
        try:
            self._loading_timer.Start()
        except Exception:
            pass

    def _stop_loading_progress(self, show_complete=False):
        if self._loading_timer is not None:
            try:
                self._loading_timer.Stop()
            except Exception:
                pass
        if show_complete:
            self._set_loading_progress_value(100)
        else:
            self._set_loading_progress_value(0)

    def _complete_loading_progress(self):
        self._stop_loading_progress(show_complete=True)
        if self._loading_complete_timer is None:
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(350)
            timer.Tick += self.loading_complete_tick
            self._loading_complete_timer = timer
        try:
            self._loading_complete_timer.Stop()
            self._loading_complete_timer.Start()
        except Exception:
            self._hide_loading_overlay()

    def _hide_loading_overlay(self):
        try:
            if self._loading_overlay is not None:
                self._loading_overlay.Visibility = Visibility.Collapsed
        except Exception:
            pass
        self._set_loading_progress_value(0)

    def loading_complete_tick(self, sender, args):
        try:
            if self._loading_complete_timer is not None:
                self._loading_complete_timer.Stop()
        except Exception:
            pass
        self._hide_loading_overlay()

    def loading_progress_tick(self, sender, args):
        current = int(self._loading_progress_value or 0)
        if current < 55:
            current += 9
        elif current < 78:
            current += 4
        elif current < 90:
            current += 1
        else:
            current = 90
        self._set_loading_progress_value(current)

    def begin_initial_refresh(self):
        self._set_busy(True, "Loading QC results...", show_loading=True)
        try:
            self.Dispatcher.BeginInvoke(
                Action(self._start_initial_refresh),
                DispatcherPriority.Background,
            )
        except Exception:
            self._start_initial_refresh()

    def _start_initial_refresh(self):
        if self._gateway is None:
            self._set_status("Ready")
            return
        raised = self._gateway.raise_refresh(callback=self._handle_external_complete)
        if not raised:
            self._set_busy(False, show_loading=False)
            self._set_status("Operation already running")
            return

    def _apply_snapshot(self, snapshot):
        data = dict(snapshot or {})
        self._doc_title = str(data.get("doc_title") or "-")
        self._doc_key = str(data.get("doc_key") or "")
        self._rows = list(data.get("rows") or [])
        if self._document_text is not None:
            self._document_text.Text = "Document: {}".format(self._doc_title)
        if self._count_text is not None:
            count = len(self._rows)
            self._count_text.Text = "{} issue{}".format(count, "" if count == 1 else "s")
        self._apply_rows_to_grid()
        self._build_column_menu()
        status = str(data.get("status") or "Ready")
        if status == "ok":
            status = "Ready"
        self._set_status(status)
        self._sync_selection_buttons()

    def _clear_for_active_document(self, status_text=None):
        doc = _active_doc()
        snapshot = _blank_snapshot(doc, status=status_text or "Click Refresh to check active document.")
        self._apply_snapshot(snapshot)

    def _active_document_changed(self):
        current_key = _doc_key(_active_doc())
        return bool(self._doc_key) and bool(current_key) and current_key != self._doc_key

    def handle_possible_document_switch(self, doc=None):
        current_doc = doc or _active_doc()
        current_key = _doc_key(current_doc)
        if bool(self._doc_key) and bool(current_key) and current_key != self._doc_key:
            self._apply_snapshot(_blank_snapshot(current_doc, status="Active document changed. Click Refresh to check active document."))
            return True
        if not self._doc_key:
            self._doc_key = current_key
            self._doc_title = _doc_title(current_doc)
        return False

    def _severity_sort_value(self, severity):
        order = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3, "INFO": 4}
        return order.get(str(severity or "").upper(), 9)

    def _group_sort_key(self, row):
        if row is None:
            return ""
        if self._sort_key == "category":
            return str(getattr(row, "category", "") or "").lower()
        if self._sort_key == "check_id":
            return str(getattr(row, "check_id", "") or "").lower()
        if self._sort_key == "severity":
            return self._severity_sort_value(getattr(row, "severity", ""))
        if self._sort_key == "issue":
            return str(getattr(row, "generic_issue", getattr(row, "issue", "")) or "").lower()
        if self._sort_key == "recommended_action":
            return str(getattr(row, "recommended_action", "") or "").lower()
        return (
            str(getattr(row, "category", "") or "").lower(),
            str(getattr(row, "check_id", "") or "").lower(),
            self._severity_sort_value(getattr(row, "severity", "")),
            str(getattr(row, "generic_issue", getattr(row, "issue", "")) or "").lower(),
            str(getattr(row, "recommended_action", "") or "").lower(),
        )

    def _row_detail_sort_key(self, row, original_index=0):
        return (
            str(getattr(row, "observed", "") or "").lower(),
            str(getattr(row, "primary_element", "") or "").lower(),
            str(getattr(row, "secondary_element", "") or "").lower(),
            original_index,
        )

    def _prepare_row_display(self, row, category=None, severity=None, issue=None, check_id=None, recommended_action=None):
        if row is None:
            return
        try:
            row.display_category = category if category is not None else getattr(row, "category", "")
        except Exception:
            pass
        try:
            row.display_severity = severity if severity is not None else getattr(row, "severity", "")
        except Exception:
            pass
        try:
            row.display_issue = issue if issue is not None else getattr(row, "generic_issue", getattr(row, "issue", ""))
        except Exception:
            pass
        try:
            row.display_check_id = check_id if check_id is not None else getattr(row, "check_id", "")
        except Exception:
            pass
        try:
            row.display_recommended_action = (
                recommended_action if recommended_action is not None else getattr(row, "recommended_action", "")
            )
        except Exception:
            pass

    def _set_group_flags(self, row, index=0, count=1):
        if row is None:
            return
        try:
            group_count = int(count or 1)
        except Exception:
            group_count = 1
        try:
            group_index = int(index or 0)
        except Exception:
            group_index = 0
        is_single = group_count <= 1
        try:
            row.is_group_single = is_single
            row.is_group_first = (not is_single) and group_index == 0
            row.is_group_last = (not is_single) and group_index == group_count - 1
            row.is_group_continuation = (not is_single) and group_index > 0
            row.is_group_middle = (not is_single) and group_index > 0 and group_index < group_count - 1
        except Exception:
            pass

    def _rows_grouped_by_alert(self):
        indexed = [(index, row) for index, row in enumerate(list(self._rows or []))]
        groups = {}
        group_order = []
        for original_index, row in indexed:
            key = str(getattr(row, "check_id", "") or getattr(row, "issue", "") or "")
            if key not in groups:
                groups[key] = []
                group_order.append(key)
            groups[key].append((original_index, row))

        def append_current_group():
            count = len(current_rows)
            for group_index, group_row in enumerate(current_rows):
                self._set_group_flags(group_row, group_index, count)
                if group_index == 0:
                    self._prepare_row_display(group_row)
                else:
                    self._prepare_row_display(
                        group_row,
                        category="",
                        severity="",
                        issue="",
                        check_id="",
                        recommended_action="",
                    )
                grouped.append(group_row)

        group_headers = []
        for order_index, key in enumerate(group_order):
            rows = [row for _, row in sorted(groups.get(key, []), key=lambda item: self._row_detail_sort_key(item[1], item[0]))]
            header = rows[0] if rows else None
            group_headers.append((order_index, key, header, rows))
        group_headers.sort(
            key=lambda item: (
                self._group_sort_key(item[2]),
                item[0],
            ),
            reverse=bool(self._sort_descending),
        )

        grouped = []
        for _, _, _, current_rows in group_headers:
            append_current_group()
        return grouped

    def _apply_rows_to_grid(self):
        if self._grid is not None:
            try:
                display_rows = self._rows_grouped_by_alert()
                self._grid.CanUserSortColumns = True
                self._grid.ItemsSource = display_rows
            except Exception:
                pass

    def _selected_row(self):
        if self._grid is None:
            return None
        try:
            return getattr(self._grid, "SelectedItem", None)
        except Exception:
            return None

    def _row_element_ids(self, field_name):
        row = self._selected_row()
        if row is None:
            return []
        values = []
        raw = getattr(row, field_name, None)
        if raw is None and field_name.endswith("_ids"):
            raw = getattr(row, field_name[:-1], 0)
        if isinstance(raw, (list, tuple)):
            candidates = list(raw)
        else:
            candidates = [raw]
        for value in candidates:
            try:
                numeric = int(value or 0)
            except Exception:
                numeric = 0
            if numeric > 0 and numeric not in values:
                values.append(numeric)
        return values

    def _sync_selection_buttons(self):
        busy = self._gateway is not None and self._gateway.is_busy()
        button_map = (
            (self._select_primary_button, "primary_element_ids"),
            (self._select_secondary_button, "secondary_element_ids"),
        )
        for button, field_name in button_map:
            if button is None:
                continue
            try:
                button.IsEnabled = (not busy) and len(self._row_element_ids(field_name)) > 0
            except Exception:
                pass

    def _column_is_visible(self, column):
        try:
            return column.Visibility == Visibility.Visible
        except Exception:
            return True

    def _column_header(self, column):
        try:
            return str(column.Header or "")
        except Exception:
            return ""

    def _apply_resource_style(self, target, key):
        if target is None:
            return
        try:
            target.Style = self.FindResource(key)
        except Exception:
            pass

    def _build_column_menu(self):
        if self._grid is None:
            return
        menu = ContextMenu()
        try:
            menu.StaysOpen = True
        except Exception:
            pass
        self._apply_resource_style(menu, "CED.ContextMenu.Base")
        columns = self._grid.Columns
        try:
            count = int(columns.Count)
        except Exception:
            count = 0
        for index in range(count):
            column = columns[index]
            header = self._column_header(column)
            if not header:
                continue
            item = MenuItem()
            self._apply_resource_style(item, "CED.MenuItem.Base")
            item.Header = header
            item.Tag = index
            item.IsCheckable = True
            item.IsChecked = self._column_is_visible(column)
            try:
                item.StaysOpenOnClick = True
            except Exception:
                pass
            item.Click += self.column_menu_clicked
            menu.Items.Add(item)
        self._column_menu = menu
        if self._columns_button is not None:
            self._columns_button.ContextMenu = menu

    def _menu_item_column(self, item):
        if self._grid is None or item is None:
            return None
        try:
            index = int(getattr(item, "Tag", -1))
            columns = self._grid.Columns
            if index < 0 or index >= int(columns.Count):
                return None
            return columns[index]
        except Exception:
            return None

    def _sync_column_menu(self):
        if self._column_menu is None:
            return
        try:
            items = list(self._column_menu.Items)
        except Exception:
            items = []
        for item in items:
            try:
                item.IsChecked = self._column_is_visible(self._menu_item_column(item))
            except Exception:
                pass

    def column_menu_clicked(self, sender, args):
        try:
            column = self._menu_item_column(sender)
            if column is None:
                raise Exception("Column reference not found.")
            make_visible = not self._column_is_visible(column)
            column.Visibility = Visibility.Visible if make_visible else Visibility.Collapsed
            sender.IsChecked = make_visible
            if self._grid is not None:
                self._grid.UpdateLayout()
        except Exception as ex:
            try:
                LOGGER.warning("Electrical QC column toggle failed: {}".format(ex))
            except Exception:
                pass
            self._set_status("Column toggle failed")

    def columns_clicked(self, sender, args):
        self._build_column_menu()
        if self._column_menu is None:
            return
        try:
            self._sync_column_menu()
            self._column_menu.PlacementTarget = sender
            self._column_menu.IsOpen = True
        except Exception:
            pass

    def grid_selection_changed(self, sender, args):
        self._sync_selection_buttons()

    def grid_loaded(self, sender, args):
        if self._grid_revealed or self._grid_reveal_queued:
            return
        self._grid_reveal_queued = True
        try:
            self.Dispatcher.BeginInvoke(
                Action(self._finish_grid_first_layout),
                DispatcherPriority.ContextIdle,
            )
        except Exception:
            self._reveal_grid()

    def _finish_grid_first_layout(self):
        try:
            if self._grid is not None:
                self._grid.UpdateLayout()
            self.UpdateLayout()
        except Exception:
            pass
        try:
            self.Dispatcher.BeginInvoke(
                Action(self._reveal_grid),
                DispatcherPriority.ApplicationIdle,
            )
        except Exception:
            self._reveal_grid()

    def _reveal_grid(self):
        if self._grid_revealed:
            return
        try:
            if self._grid is not None:
                self._grid.UpdateLayout()
            self.UpdateLayout()
        except Exception:
            pass
        try:
            self._grid.Visibility = Visibility.Visible
        except Exception:
            pass
        self._grid_revealed = True
        self._grid_reveal_queued = False

    def grid_sorting(self, sender, args):
        supported = set(["category", "check_id", "severity", "issue", "recommended_action"])
        try:
            args.Handled = True
        except Exception:
            pass
        try:
            column = args.Column
            sort_key = str(getattr(column, "SortMemberPath", "") or "")
        except Exception:
            sort_key = ""
            column = None
        if sort_key not in supported:
            self._set_status("Only grouped alert columns can sort this view")
            return
        if self._sort_key == sort_key:
            self._sort_descending = not bool(self._sort_descending)
        else:
            self._sort_key = sort_key
            self._sort_descending = False
        if self._grid is not None:
            try:
                for grid_column in self._grid.Columns:
                    grid_column.SortDirection = None
            except Exception:
                pass
        if column is not None:
            try:
                column.SortDirection = ListSortDirection.Descending if self._sort_descending else ListSortDirection.Ascending
            except Exception:
                pass
        self._apply_rows_to_grid()
        self._set_status("Sorted by {}".format(self._column_header(column) or sort_key))

    def _handle_external_complete(self, status, op_name, result, error):
        if status == "error":
            self._set_busy(False, show_loading=False)
            self._set_status("Operation failed")
            forms.alert("Electrical QC operation failed:\n\n{}".format(error), title=TITLE)
            return
        if op_name == "refresh":
            self._apply_snapshot(result or {})
            self._set_busy(False, show_loading=True)
            self._complete_loading_progress()
            self._set_status("Refreshed")
            return
        if op_name == "select":
            self._set_busy(False, show_loading=False)
            count = int((result or {}).get("selected") or 0)
            self._set_status("Selected {} element{}".format(count, "" if count == 1 else "s"))
            return
        self._set_busy(False, show_loading=False)

    def refresh_clicked(self, sender, args):
        if self._gateway is None:
            return
        raised = self._gateway.raise_refresh(callback=self._handle_external_complete)
        if not raised:
            self._set_status("Operation already running")
            return
        self._set_busy(True, "Refreshing...", show_loading=True)

    def _raise_select_field(self, field_name):
        if self.handle_possible_document_switch():
            self._set_status("Active document changed. Click Refresh to check active document.")
            return
        element_ids = self._row_element_ids(field_name)
        if not element_ids:
            self._set_status("No element available for this row")
            return
        if self._gateway is None:
            return
        raised = self._gateway.raise_select(element_ids, callback=self._handle_external_complete)
        if not raised:
            self._set_status("Operation already running")
            return
        self._set_busy(True, "Selecting element...", show_loading=False)

    def select_primary_clicked(self, sender, args):
        self._raise_select_field("primary_element_ids")

    def select_secondary_clicked(self, sender, args):
        self._raise_select_field("secondary_element_ids")

    def export_clicked(self, sender, args):
        if self.handle_possible_document_switch():
            self._set_status("Active document changed. Click Refresh before exporting.")
            return
        try:
            path = forms.save_file(
                file_ext="csv",
                title="Export Electrical QC Check",
                default_name="Electrical_QC_Check",
            )
        except Exception as ex:
            forms.alert("Failed to open CSV export dialog:\n\n{}".format(ex), title=TITLE)
            self._set_status("Export dialog failed")
            return
        if not path:
            self._set_status("Export canceled")
            return
        if not str(path).lower().endswith(".csv"):
            path = "{}.csv".format(path)
        try:
            File.WriteAllText(path, rows_to_csv(self._rows), Encoding.UTF8)
            self._set_status("Exported {}".format(path))
        except Exception as ex:
            forms.alert("Failed to export CSV:\n\n{}".format(ex), title=TITLE)
            self._set_status("Export failed")

    def close_clicked(self, sender, args):
        self.Close()

    def window_activated(self, sender, args):
        self.handle_possible_document_switch()


def _find_existing_window():
    app = Application.Current
    if app is None:
        return None
    try:
        windows = list(app.Windows)
    except Exception:
        windows = []
    for window in windows:
        try:
            if bool(getattr(window, WINDOW_MARKER, False)):
                return window
        except Exception:
            pass
        try:
            if str(getattr(window, "Tag", "") or "") == WINDOW_MARKER:
                return window
        except Exception:
            continue
    return None


def _focus_existing_window(window):
    try:
        window.handle_possible_document_switch(_active_doc())
    except Exception:
        pass
    try:
        window.refresh_ced_theme_from_config()
    except Exception:
        pass
    try:
        if getattr(window, "WindowState", None) == WindowState.Minimized:
            window.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        window.Show()
    except Exception:
        pass
    try:
        window.Activate()
    except Exception:
        pass
    try:
        window.Focus()
    except Exception:
        pass


def _show_or_focus_window():
    existing = _find_existing_window()
    if existing is not None:
        _focus_existing_window(existing)
        return
    gateway = ElectricalQCGateway(logger=LOGGER)
    snapshot = _blank_snapshot(_active_doc(), status="Loading QC results...")
    window = ElectricalQCWindow(snapshot=snapshot, gateway=gateway)
    window.Show()
    window.begin_initial_refresh()
    try:
        window.Activate()
    except Exception:
        pass


_show_or_focus_window()
