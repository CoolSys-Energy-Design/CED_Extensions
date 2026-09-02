# -*- coding: utf-8 -*-
"""Modeless revision-cloud review, editing, reporting, and export."""

__title__ = "Revision Manager"
__doc__ = "Review revision clouds, update comments, and create revision reports."

import json
import os
import sys

import clr

for _assembly in ("System", "PresentationFramework", "PresentationCore", "WindowsBase", "System.Windows.Forms"):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass


def _load_clr_system_types():
    """Load CLR System types, repairing a file-backed namespace collision once."""
    try:
        from System import EventHandler as ClrEventHandler
        from System import Uri as ClrUri
        from System import UriKind as ClrUriKind
        return ClrEventHandler, ClrUri, ClrUriKind
    except ImportError:
        loaded_system = sys.modules.get("System")
        if loaded_system is None or not getattr(loaded_system, "__file__", None):
            raise
        # CLR namespaces do not have __file__. A file-backed System entry is a
        # Python module collision left in the shared pyRevit engine.
        del sys.modules["System"]
        clr.AddReference("System")
        from System import EventHandler as ClrEventHandler
        from System import Uri as ClrUri
        from System import UriKind as ClrUriKind
        return ClrEventHandler, ClrUri, ClrUriKind


EventHandler, Uri, UriKind = _load_clr_system_types()

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Events import ViewActivatedEventArgs
from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
from System import Action
from System.Collections.Generic import List
from System.Windows import Application, FontWeights, GridLength, HorizontalAlignment, TextAlignment, Thickness, Visibility, WindowState
from System.Windows.Threading import DispatcherPriority
from System.Windows.Controls import Border, Image, TextBox
from System.Windows.Controls.Primitives import DocumentPageView
from System.Windows.Documents import BlockUIContainer, FlowDocument, IDocumentPaginatorSource, Paragraph, Run, Table, TableCell, TableColumn, TableRow, TableRowGroup
from System.Windows.Forms import DialogResult, SaveFileDialog
from System.Windows.Media import Brushes, Color, FontFamily, SolidColorBrush, Stretch
from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage
from pyrevit import DB, coreutils, forms, revit, script
from pyrevit.revit import query


THIS_DIR = os.path.abspath(os.path.dirname(__file__))


def _find_lib_root(start_dir):
    current = os.path.abspath(start_dir)
    while True:
        candidate = os.path.join(current, "CEDLib.lib")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(current)
        if parent == current:
            return None
        current = parent


LIB_ROOT = _find_lib_root(THIS_DIR)
if not LIB_ROOT:
    forms.alert("Could not locate CEDLib.lib for Revision Manager.", title="Revision Cloud Manager")
    raise SystemExit
if LIB_ROOT in sys.path:
    sys.path.remove(LIB_ROOT)
sys.path.insert(0, LIB_ROOT)
if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

from Snippets import revit_helpers
from UIClasses.ui_bases import CEDWindowBase
from revision_manager_core import (
    CloudRow,
    RevisionOption,
    build_report_column_options,
    build_report_rows,
    default_report_column_width,
    filter_clouds,
    group_report_rows,
    report_column_value,
    selected_report_columns,
    selected_revision_ids,
    text,
)
from revision_manager_exporters import copy_report_to_clipboard, export_pdf, export_xlsx


TITLE = "Revision Cloud Manager"
WINDOW_MARKER = "_ced_revision_cloud_manager_window_v1"
LOGGER = script.get_logger()
LOGO_PATH = os.path.join(THIS_DIR, "CED_Logo_H.png")
REPORT_CONFIG_SECTION = "revision_manager_report_schema"
REPORT_BLUE_BRUSH = SolidColorBrush(Color.FromRgb(3, 67, 123))
# Printed page colours are fixed, not themed - the preview mirrors paper, where a
# pale amber row reads as "needs attention" without going muddy in greyscale.
REPORT_ISSUE_ROW_BRUSH = SolidColorBrush(Color.FromRgb(255, 244, 214))
# Toggleable review-grid columns: (key, checkbox name, CloudGrid column index).
# Rev (0), Sheet (1) and View (3) are always shown and deliberately absent here -
# they are the identity of a row, so there is nothing useful to hide.
REVIEW_COLUMN_DEFINITIONS = (
    ("sheet_name", "ReviewSheetNameCheck", 2),
    ("comments", "ReviewCommentsCheck", 4),
    ("placement", "ReviewPlacementCheck", 5),
    ("issues", "ReviewIssuesCheck", 6),
    ("created_by", "ReviewCreatedByCheck", 7),
    ("edited_by", "ReviewEditedByCheck", 8),
    ("owned_by", "ReviewOwnedByCheck", 9),
    ("cloud_id", "ReviewCloudIdCheck", 10),
)
# Columns that stay hidden until asked for. A saved preference always wins; this
# only decides the first run, and existing saved sets simply lack these keys.
REVIEW_COLUMNS_DEFAULT_OFF = ("created_by", "edited_by", "owned_by")


def _id_value(value, default=0):
    return revit_helpers.get_elementid_value(value, default=default)


def _id_from_value(value):
    return revit_helpers.elementid_from_value(value)


def _document_key(doc):
    if doc is None:
        return ""
    try:
        return "{}|{}".format(doc.PathName or "", doc.Title or "")
    except Exception:
        return ""


def _document_label(doc):
    if doc is None:
        return "No active document"
    try:
        return text(doc.Title, "Untitled")
    except Exception:
        return "Untitled"


def _active_uidoc():
    try:
        return __revit__.ActiveUIDocument
    except Exception:
        return None


def _parameter_text(element, name, default=""):
    return revit_helpers.get_parameter_text(element, name, default=default)


def _parameter_display_value(parameter):
    """Convert a Revit parameter to stable UI/report text at the DTO boundary."""
    if parameter is None:
        return ""
    try:
        storage_type = parameter.StorageType
    except Exception:
        return ""
    try:
        if storage_type == DB.StorageType.String:
            return text(parameter.AsString() or parameter.AsValueString() or "")
        if storage_type == DB.StorageType.ElementId:
            native_id = parameter.AsElementId()
            display = parameter.AsValueString()
            return text(display) if display else text(_id_value(native_id))
        display = parameter.AsValueString()
        if display:
            return text(display)
        if storage_type == DB.StorageType.Integer:
            return text(parameter.AsInteger())
        if storage_type == DB.StorageType.Double:
            return text(parameter.AsDouble())
    except Exception:
        return ""
    return ""


def _revision_cloud_revision_display(cloud, fallback="N/A"):
    """Return the exact revision value Revit displays for a revision cloud."""
    if cloud is None:
        return text(fallback, "N/A")
    try:
        parameter = cloud.get_Parameter(DB.BuiltInParameter.REVISION_CLOUD_REVISION)
        value = parameter.AsValueString() if parameter is not None else None
        if value:
            return text(value)
    except Exception:
        pass
    return text(fallback, "N/A")


def _element_parameter_map(element):
    values = {}
    if element is None:
        return values
    try:
        parameters = list(element.Parameters or [])
    except Exception:
        parameters = []
    for parameter in parameters:
        try:
            definition = parameter.Definition
            name = text(definition.Name).strip() if definition else ""
        except Exception:
            name = ""
        if not name or name in values:
            continue
        values[name] = _parameter_display_value(parameter)
    return values


def _worksharing_owner_state(doc, element):
    """Return current ownership display state without crossing the DTO boundary with IDs."""
    current_user = ""
    try:
        current_user = text(doc.Application.Username)
    except Exception:
        pass
    try:
        if not bool(doc.IsWorkshared):
            return "", "", "", current_user, False
    except Exception:
        return "", "", "", current_user, False

    owner = ""
    creator = ""
    edited_by = ""
    status = None
    try:
        status = DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id)
    except Exception:
        pass
    try:
        tooltip = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id)
        owner = text(tooltip.Owner)
        creator = text(tooltip.Creator)
        edited_by = text(tooltip.LastChangedBy)
    except Exception:
        pass
    try:
        owned_by_other = status == DB.CheckoutStatus.OwnedByOtherUser
    except Exception:
        owned_by_other = False
    if not owned_by_other and owner and current_user:
        owned_by_other = owner.strip().lower() != current_user.strip().lower()
    return owner, creator, edited_by, current_user, bool(owned_by_other)


def _revision_values(revision):
    if revision is None:
        return ("N/A", "", "", 0)
    number = _parameter_text(revision, "Revision Number", "N/A")
    date_value = _parameter_text(revision, "Revision Date", "")
    description = _parameter_text(revision, "Revision Description", "")
    try:
        sequence = int(revision.SequenceNumber)
    except Exception:
        sequence = 0
    return number, date_value, description, sequence


def _element_name(element, default=""):
    if element is None:
        return default
    try:
        value = query.get_name(element)
        if value:
            return text(value)
    except Exception:
        pass
    try:
        return text(DB.Element.Name.__get__(element), default)
    except Exception:
        return default


def _sheet_values(sheet):
    if sheet is None:
        return ("N/A", "N/A")
    number = _parameter_text(sheet, "Sheet Number", "N/A") or "N/A"
    name = _parameter_text(sheet, "Sheet Name", "N/A") or "N/A"
    return number, name


def _project_metadata(doc):
    values = {
        "project_number": "",
        "client": "",
        "project_name": "",
        "report_date": coreutils.current_date(),
    }
    try:
        project_info = revit.query.get_project_info()
        values["project_number"] = text(project_info.number)
        values["client"] = text(project_info.client_name)
        values["project_name"] = text(project_info.name)
        return values
    except Exception:
        pass
    try:
        info = doc.ProjectInformation
    except Exception:
        info = None
    if info is not None:
        values["project_number"] = _parameter_text(info, "Project Number", "")
        values["client"] = _parameter_text(info, "Client Name", "")
        values["project_name"] = _parameter_text(info, "Project Name", "")
    return values


def build_snapshot(doc):
    if doc is None:
        return {
            "document_key": "",
            "document_label": "No active document",
            "clouds": [],
            "revisions": [],
            "revision_parameter_names": [],
            "cloud_parameter_names": [],
            "metadata": _project_metadata(doc),
        }

    revisions = []
    revision_lookup = {}
    try:
        revision_elements = DB.FilteredElementCollector(doc).OfClass(DB.Revision).ToElements()
    except Exception:
        revision_elements = []
    for revision in list(revision_elements or []):
        revision_id = _id_value(revision.Id)
        number, date_value, description, sequence = _revision_values(revision)
        revision_lookup[revision_id] = revision
        revisions.append(RevisionOption(
            revision_id=revision_id,
            number=number,
            description=description,
            date_value=date_value,
            sequence=sequence,
            selected=True,
        ))
    revisions = sorted(revisions, key=lambda x: (x.sequence, x.number))

    try:
        cloud_elements = (DB.FilteredElementCollector(doc)
                          .OfCategory(DB.BuiltInCategory.OST_RevisionClouds)
                          .WhereElementIsNotElementType()
                          .ToElements())
    except Exception:
        cloud_elements = []

    clouds = []
    cloud_parameter_names = set()
    for cloud in list(cloud_elements or []):
        cloud_id = _id_value(cloud.Id)
        revision_id = _id_value(cloud.RevisionId)
        revision = revision_lookup.get(revision_id)
        if revision is None:
            try:
                revision = doc.GetElement(cloud.RevisionId)
            except Exception:
                revision = None
        number, date_value, description, sequence = _revision_values(revision)
        revision_display = _revision_cloud_revision_display(cloud, number)

        try:
            owner_view = doc.GetElement(cloud.OwnerViewId)
        except Exception:
            owner_view = None
        owner_view_id = _id_value(getattr(cloud, "OwnerViewId", None))
        view_name = _element_name(owner_view, "Unknown View")
        try:
            view_type = text(owner_view.ViewType)
        except Exception:
            view_type = ""

        native_sheet_ids = []
        try:
            native_sheet_ids = list(cloud.GetSheetIds() or [])
        except Exception:
            native_sheet_ids = []
        if isinstance(owner_view, DB.ViewSheet):
            owner_numeric = _id_value(owner_view.Id)
            if owner_numeric not in [_id_value(item) for item in native_sheet_ids]:
                native_sheet_ids.append(owner_view.Id)

        sheet_ids = []
        sheet_numbers = []
        sheet_names = []
        for native_sheet_id in native_sheet_ids:
            numeric_sheet_id = _id_value(native_sheet_id)
            if numeric_sheet_id in sheet_ids:
                continue
            try:
                sheet = doc.GetElement(native_sheet_id)
            except Exception:
                sheet = None
            if not isinstance(sheet, DB.ViewSheet):
                continue
            sheet_ids.append(numeric_sheet_id)
            sheet_number, sheet_name = _sheet_values(sheet)
            sheet_numbers.append(sheet_number)
            sheet_names.append(sheet_name)

        comments = _parameter_text(cloud, "Comments", "")
        cloud_parameters = _element_parameter_map(cloud)
        cloud_parameter_names.update(cloud_parameters.keys())
        worksharing_owner, created_by, edited_by, current_user, owned_by_other = _worksharing_owner_state(doc, cloud)
        cloud_in_view = owner_view is not None and not isinstance(owner_view, DB.ViewSheet)
        not_on_sheet = cloud_in_view and not bool(sheet_ids)
        if isinstance(owner_view, DB.ViewSheet):
            placement = "On Sheet"
        elif sheet_ids:
            placement = "In View"
        else:
            placement = "Not on Sheet"

        clouds.append(CloudRow(
            cloud_id=cloud_id,
            revision_id=revision_id,
            owner_view_id=owner_view_id,
            revision_number=number,
            revision_display=revision_display,
            revision_sort=sequence,
            revision_date=date_value,
            revision_description=description,
            sheet_ids=sheet_ids,
            sheet_numbers_list=sheet_numbers,
            sheet_names_list=sheet_names,
            view_name=view_name,
            view_type=view_type,
            comments=comments,
            placement=placement,
            missing_comment=not bool(comments.strip()),
            cloud_in_view=cloud_in_view,
            not_on_sheet=not_on_sheet,
            cloud_parameters=cloud_parameters,
            worksharing_owner=worksharing_owner,
            created_by=created_by,
            edited_by=edited_by,
            current_user=current_user,
            owned_by_other=owned_by_other,
        ))

    clouds = sorted(clouds, key=lambda x: (x.revision_sort, x.sheet_sort, x.cloud_id))
    return {
        "document_key": _document_key(doc),
        "document_label": _document_label(doc),
        "clouds": clouds,
        "revisions": revisions,
        "revision_parameter_names": [],
        "cloud_parameter_names": sorted(cloud_parameter_names),
        "metadata": _project_metadata(doc),
    }


class RevisionManagerGateway(object):
    def __init__(self, ui_application, document_key, logger=None):
        self.ui_application = ui_application
        self.document_key = text(document_key)
        self.logger = logger
        self._pending = None
        self._handler = _RevisionManagerHandler(self)
        self._event = ExternalEvent.Create(self._handler)
        self._view_activated_handler = None
        self._document_changed_handler = None
        self._document_changed_callback = None
        self._document_edited_callback = None
        self._suppress_document_changed = False
        self._subscribe_document_events()

    def _subscribe_document_events(self):
        try:
            self._view_activated_handler = EventHandler[ViewActivatedEventArgs](self._on_view_activated)
            self.ui_application.ViewActivated += self._view_activated_handler
        except Exception as ex:
            if self.logger:
                self.logger.debug("Revision Manager ViewActivated subscription failed: %s", ex)
        try:
            self._document_changed_handler = EventHandler[DocumentChangedEventArgs](self._on_document_edited)
            self.ui_application.Application.DocumentChanged += self._document_changed_handler
        except Exception as ex:
            if self.logger:
                self.logger.debug("Revision Manager DocumentChanged subscription failed: %s", ex)

    def set_document_changed_callback(self, callback):
        self._document_changed_callback = callback

    def set_document_edited_callback(self, callback):
        self._document_edited_callback = callback

    def _on_view_activated(self, sender, args):
        doc = None
        try:
            doc = args.CurrentActiveView.Document
        except Exception:
            try:
                uidoc = self.ui_application.ActiveUIDocument
                doc = uidoc.Document if uidoc else None
            except Exception:
                doc = None
        if self._document_changed_callback:
            try:
                self._document_changed_callback(_document_key(doc), _document_label(doc))
            except Exception:
                pass

    def _on_document_edited(self, sender, args):
        if self._suppress_document_changed or self._document_edited_callback is None:
            return
        try:
            doc = args.GetDocument()
        except Exception:
            doc = None
        if _document_key(doc) != self.document_key:
            return
        try:
            self._document_edited_callback(_document_key(doc), _document_label(doc))
        except Exception:
            pass

    def is_busy(self):
        try:
            pending = bool(self._event.IsPending)
        except Exception:
            pending = False
        return self._pending is not None or pending

    def _raise(self, operation, payload=None, callback=None):
        if self.is_busy():
            return False
        self._pending = {
            "operation": text(operation),
            "payload": dict(payload or {}),
            "callback": callback,
        }
        try:
            self._event.Raise()
            return True
        except Exception as ex:
            self._pending = None
            if self.logger:
                self.logger.warning("Revision Manager ExternalEvent raise failed: %s", ex)
            return False

    def raise_refresh(self, callback):
        return self._raise("refresh", callback=callback)

    def raise_show(self, cloud_id, owner_view_id, callback):
        return self._raise("show", {
            "document_key": self.document_key,
            "cloud_id": int(cloud_id or 0),
            "owner_view_id": int(owner_view_id or 0),
        }, callback)

    def raise_select(self, cloud_ids, callback):
        return self._raise("select", {
            "document_key": self.document_key,
            "cloud_ids": [int(item or 0) for item in list(cloud_ids or [])],
        }, callback)

    def raise_navigate(self, view_id, callback):
        return self._raise("navigate", {
            "document_key": self.document_key,
            "view_id": int(view_id or 0),
        }, callback)

    def raise_update_comment(self, cloud_id, original_comment, new_comment, callback):
        return self._raise("update_comment", {
            "document_key": self.document_key,
            "cloud_id": int(cloud_id or 0),
            "original_comment": text(original_comment),
            "new_comment": text(new_comment),
        }, callback)

    def consume(self):
        pending = self._pending
        self._pending = None
        return pending

    def dispose(self):
        if self._view_activated_handler is not None:
            try:
                self.ui_application.ViewActivated -= self._view_activated_handler
            except Exception:
                pass
            self._view_activated_handler = None
        if self._document_changed_handler is not None:
            try:
                self.ui_application.Application.DocumentChanged -= self._document_changed_handler
            except Exception:
                pass
            self._document_changed_handler = None
        try:
            self._event.Dispose()
        except Exception:
            pass


class _RevisionManagerHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self.gateway = gateway

    def _validated_context(self, application, payload):
        uidoc = application.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        if doc is None:
            raise Exception("No active Revit document.")
        expected = text(payload.get("document_key"))
        if expected and _document_key(doc) != expected:
            raise Exception("The active document changed. Refresh Revision Manager before continuing.")
        return uidoc, doc

    def _execute_show(self, application, payload):
        uidoc, doc = self._validated_context(application, payload)
        cloud_id = _id_from_value(payload.get("cloud_id"))
        cloud = doc.GetElement(cloud_id)
        if cloud is None:
            raise Exception("The selected revision cloud no longer exists.")
        owner_view_id = _id_from_value(payload.get("owner_view_id"))
        owner_view = doc.GetElement(owner_view_id)
        if owner_view is not None:
            try:
                uidoc.ActiveView = owner_view
            except Exception:
                pass
        selection_ids = List[DB.ElementId]()
        selection_ids.Add(cloud.Id)
        uidoc.Selection.SetElementIds(selection_ids)
        try:
            uidoc.ShowElements(cloud.Id)
        except Exception:
            uidoc.ShowElements(selection_ids)
        return {"shown": 1}

    def _execute_select(self, application, payload):
        uidoc, doc = self._validated_context(application, payload)
        selection_ids = List[DB.ElementId]()
        missing = 0
        for numeric_id in list(payload.get("cloud_ids") or []):
            native_id = _id_from_value(numeric_id)
            cloud = doc.GetElement(native_id)
            if cloud is None:
                missing += 1
                continue
            selection_ids.Add(cloud.Id)
        uidoc.Selection.SetElementIds(selection_ids)
        return {"selected": selection_ids.Count, "skipped": 0, "missing": missing}

    def _execute_navigate(self, application, payload):
        uidoc, doc = self._validated_context(application, payload)
        view_id = _id_from_value(payload.get("view_id"))
        target_view = doc.GetElement(view_id)
        if target_view is None or not isinstance(target_view, DB.View):
            raise Exception("The selected view or sheet is no longer available.")
        try:
            if bool(target_view.IsTemplate):
                raise Exception("View templates cannot be opened.")
        except AttributeError:
            pass
        uidoc.ActiveView = target_view
        return {"view_id": _id_value(target_view.Id), "view_name": _element_name(target_view, "View")}

    def _execute_update_comment(self, application, payload):
        unused_uidoc, doc = self._validated_context(application, payload)
        cloud_id = _id_from_value(payload.get("cloud_id"))
        cloud = doc.GetElement(cloud_id)
        if cloud is None:
            raise Exception("The selected revision cloud no longer exists.")
        owner, unused_creator, unused_editor, unused_current_user, owned_by_other = _worksharing_owner_state(doc, cloud)
        if owned_by_other:
            if owner:
                raise Exception("The selected revision cloud is currently owned by {}.".format(owner))
            raise Exception("The selected revision cloud is currently owned by another Revit user.")
        parameter = revit_helpers.get_parameter(cloud, "Comments", include_type=False)
        if parameter is None:
            raise Exception("The selected revision cloud does not have a Comments parameter.")
        try:
            if parameter.IsReadOnly:
                raise Exception("The Comments parameter is read-only.")
        except AttributeError:
            pass
        current = text(revit_helpers.get_parameter_value(parameter, default=""))
        original = text(payload.get("original_comment"))
        proposed = text(payload.get("new_comment"))
        if current != original:
            raise Exception("The cloud comment changed in Revit after it was loaded. Refresh before overwriting it.")
        transaction = DB.Transaction(doc, "Update Revision Cloud Comment")
        self.gateway._suppress_document_changed = True
        try:
            transaction.Start()
            parameter.Set(proposed)
            transaction.Commit()
        except Exception:
            try:
                if transaction.HasStarted():
                    transaction.RollBack()
            except Exception:
                pass
            raise
        finally:
            self.gateway._suppress_document_changed = False
        return {"cloud_id": _id_value(cloud.Id), "comment": proposed}

    def Execute(self, application):  # noqa: N802
        pending = self.gateway.consume()
        if not pending:
            return
        operation = pending.get("operation")
        payload = pending.get("payload") or {}
        callback = pending.get("callback")
        status = "ok"
        result = None
        error = None
        try:
            if operation == "refresh":
                uidoc = application.ActiveUIDocument
                doc = uidoc.Document if uidoc else None
                result = build_snapshot(doc)
                self.gateway.document_key = text(result.get("document_key"))
            elif operation == "show":
                result = self._execute_show(application, payload)
            elif operation == "select":
                result = self._execute_select(application, payload)
            elif operation == "navigate":
                result = self._execute_navigate(application, payload)
            elif operation == "update_comment":
                result = self._execute_update_comment(application, payload)
            else:
                raise Exception("Unknown Revision Manager operation: {}".format(operation))
        except Exception as ex:
            status = "error"
            error = ex
            if self.gateway.logger:
                self.gateway.logger.exception("Revision Manager operation failed: %s", ex)
        if callback:
            try:
                callback(status, operation, result, error)
            except Exception:
                pass

    def GetName(self):  # noqa: N802
        return "CED Revision Cloud Manager External Event"


class RevisionManagerWindow(CEDWindowBase):
    theme_aware = True
    auto_wire_textboxes = False

    def __init__(self, snapshot, gateway):
        self._initializing = True
        self._setting_comment = False
        self._updating_report = False
        self._updating_column_editor = False
        self._snapshot = {}
        self._clouds = []
        self._filtered_clouds = []
        self._revision_options = []
        self._report_columns = []
        self._saved_column_schema = []
        self._legacy_parameter_names = []
        self._metadata = {}
        self._document_key = ""
        self._stale = False
        self._comment_original = ""
        self._comment_cloud_id = 0
        self._comment_dirty = False
        self._queue = []
        self._queue_index = 0
        self._preview_page_width = 1056.0
        self._preview_page_height = 816.0
        self._preview_document = None
        self._preview_zoom = 75.0
        self._preview_page_index = 0
        self._preview_page_count = 0
        self._issue_toast_key = None
        self._issue_toast_dismissed_key = None
        self._issue_warning_active = False
        self._issue_warning_text = ""
        self._gateway = gateway
        CEDWindowBase.__init__(self, xaml_source=os.path.join(THIS_DIR, "RevisionManagerWindow.xaml"))
        try:
            self.Tag = WINDOW_MARKER
        except Exception:
            pass
        self._gateway.set_document_changed_callback(self._active_document_changed)
        self._gateway.set_document_edited_callback(self._document_edited)
        self._load_report_preferences()
        self._apply_snapshot(snapshot, reset_report_selection=True)
        self._initializing = False
        self._update_action_states()
        self._queue_initial_populate()

    def _queue_initial_populate(self):
        """Fill the review grid and preview after the window's first paint.

        Building the review grid is by far the most expensive layout in this
        window. Doing it inline here means the shell paints around a grid whose
        column widths have not resolved yet, so the user watches it reflow from
        collapsed to correct. Queuing at DispatcherPriority.Loaded lets the
        window come up complete and populate once. Falls back to populating
        inline if the dispatcher is unavailable for any reason.
        """
        try:
            self.QueueText.Text = "Loading clouds..."
        except Exception:
            pass
        try:
            self.Dispatcher.BeginInvoke(
                DispatcherPriority.Loaded, Action(self._initial_populate)
            )
        except Exception:
            self._initial_populate()

    def _initial_populate(self):
        try:
            self._apply_filters()
            self._update_report_preview()
        finally:
            self._update_action_states()

    def _style(self, key):
        try:
            return self.FindResource(key)
        except Exception:
            return None

    def _set_textbox_style(self, box, key):
        if box is None:
            return
        try:
            box.ClearValue(TextBox.BorderBrushProperty)
            box.ClearValue(TextBox.BorderThicknessProperty)
        except Exception:
            pass
        style_key = "CED.Input.TextBox" if key == "CommentChangedTextBox" else key
        style = self._style(style_key)
        if box is not None and style is not None:
            box.Style = style
        # Opacity and Foreground belong to the style. Setting them here writes local
        # values, which outrank every setter in CED.Input.TextBox.ReadOnly - that is
        # what made read-only metadata fields on the Report tab look editable: they
        # were pinned to full-strength PrimaryText at 0.9 opacity, identical to an
        # editable field in dark themes. Clear them and let the style decide.
        try:
            box.ClearValue(TextBox.OpacityProperty)
            box.ClearValue(TextBox.ForegroundProperty)
        except Exception:
            pass
        if key == "CommentChangedTextBox":
            try:
                box.BorderBrush = self.FindResource("CED.Brush.AccentBlue")
                box.BorderThickness = Thickness(2)
            except Exception:
                pass

    def _load_report_preferences(self):
        self._report_config = script.get_config(REPORT_CONFIG_SECTION)
        raw_schema = text(getattr(self._report_config, "column_schema_json", ""))
        try:
            parsed = json.loads(raw_schema) if raw_schema else []
            self._saved_column_schema = parsed if isinstance(parsed, list) else []
        except Exception:
            self._saved_column_schema = []
        if not self._saved_column_schema:
            try:
                legacy_config = script.get_config("revision_parameters_config")
                legacy_raw = text(getattr(legacy_config, "selected_param_names", ""))
                self._legacy_parameter_names = [item for item in legacy_raw.split(",") if item]
            except Exception:
                self._legacy_parameter_names = []
        include_logo = text(getattr(self._report_config, "include_logo", "true")).strip().lower() == "true"
        deduplicate = text(getattr(self._report_config, "deduplicate", "true")).strip().lower() != "false"
        show_blank_comment_rows = text(getattr(
            self._report_config, "show_blank_comment_rows", "true")).strip().lower() != "false"
        checked_only = text(getattr(self._report_config, "show_checked_columns_only", "false")).strip().lower() == "true"
        orientation = text(getattr(self._report_config, "page_orientation", "landscape")).strip().lower()
        font_name = text(getattr(self._report_config, "report_font_name", "Arial"), "Arial")
        table_font_size = text(getattr(self._report_config, "report_table_font_size", "10.5"), "10.5")
        raw_review_columns = text(getattr(self._report_config, "review_columns_json", ""))
        default_review_columns = set(
            item[0] for item in REVIEW_COLUMN_DEFINITIONS
            if item[0] not in REVIEW_COLUMNS_DEFAULT_OFF)
        try:
            review_columns = (set(json.loads(raw_review_columns)) if raw_review_columns
                              else set(default_review_columns))
        except Exception:
            review_columns = set(default_review_columns)
        self.IncludeLogoCheck.IsChecked = include_logo
        self.DeduplicateCheck.IsChecked = deduplicate
        self.ShowBlankCommentRowsCheck.IsChecked = show_blank_comment_rows
        self.ShowCheckedColumnsOnlyCheck.IsChecked = checked_only
        self.PageOrientationCombo.SelectedIndex = 1 if orientation == "portrait" else 0
        self._select_combo_content(self.ReportFontCombo, font_name, fallback_index=0)
        self._select_combo_content(
            self.ReportTableFontSizeCombo,
            "{} pt".format(table_font_size.replace(".0", "")),
            fallback_index=1)
        self._set_review_column_visibility(review_columns)

    def _save_report_preferences(self):
        schema = [
            {"key": item.key, "selected": bool(item.is_selected),
             "width": float(item.width_weight), "bold": bool(item.is_bold)}
            for item in self._report_columns
        ]
        self._saved_column_schema = schema
        self._report_config.column_schema_json = json.dumps(schema, separators=(",", ":"))
        self._report_config.include_logo = "true" if bool(self.IncludeLogoCheck.IsChecked) else "false"
        self._report_config.deduplicate = "true" if bool(self.DeduplicateCheck.IsChecked) else "false"
        self._report_config.show_blank_comment_rows = (
            "true" if bool(self.ShowBlankCommentRowsCheck.IsChecked) else "false")
        self._report_config.show_checked_columns_only = (
            "true" if bool(self.ShowCheckedColumnsOnlyCheck.IsChecked) else "false")
        self._report_config.page_orientation = self._page_orientation()
        self._report_config.report_font_name = self._report_font_name()
        self._report_config.report_table_font_size = text(self._report_table_font_size())
        self._report_config.review_columns_json = json.dumps(
            self._selected_review_column_keys(), separators=(",", ":"))
        script.save_config()

    def _select_combo_content(self, combo, desired, fallback_index=0):
        desired = text(desired).strip().lower()
        try:
            for index, item in enumerate(list(combo.Items or [])):
                content = text(getattr(item, "Content", item)).strip().lower()
                if content == desired:
                    combo.SelectedIndex = index
                    return
            combo.SelectedIndex = int(fallback_index)
        except Exception:
            pass

    def _report_font_name(self):
        try:
            return text(getattr(self.ReportFontCombo.SelectedItem, "Content", "Arial"), "Arial")
        except Exception:
            return "Arial"

    def _report_table_font_size(self):
        try:
            value = text(getattr(self.ReportTableFontSizeCombo.SelectedItem, "Content", "10.5"))
            return max(8.0, min(14.0, float(value.replace("pt", "").replace("px", "").strip())))
        except Exception:
            return 10.5

    def _selected_review_column_keys(self):
        selected = []
        for key, checkbox_name, column_index in REVIEW_COLUMN_DEFINITIONS:
            try:
                if bool(getattr(self, checkbox_name).IsChecked):
                    selected.append(key)
            except Exception:
                selected.append(key)
        return selected

    def _set_review_column_visibility(self, selected_keys):
        selected = set(selected_keys or [])
        for key, checkbox_name, column_index in REVIEW_COLUMN_DEFINITIONS:
            is_visible = key in selected
            try:
                getattr(self, checkbox_name).IsChecked = is_visible
                self.CloudGrid.Columns[column_index].Visibility = (
                    Visibility.Visible if is_visible else Visibility.Collapsed)
            except Exception:
                pass

    def _page_orientation(self):
        try:
            selected = self.PageOrientationCombo.SelectedItem
            value = text(getattr(selected, "Content", selected)).strip().lower()
            return "portrait" if value == "portrait" else "landscape"
        except Exception:
            return "landscape"

    def _visible_report_columns(self):
        if bool(self.ShowCheckedColumnsOnlyCheck.IsChecked):
            return [item for item in self._report_columns if item.is_selected]
        return list(self._report_columns)

    def _refresh_report_column_grid(self, selected=None):
        visible = self._visible_report_columns()
        self.ReportColumnGrid.ItemsSource = None
        self.ReportColumnGrid.ItemsSource = visible
        if selected in visible:
            self.ReportColumnGrid.SelectedItem = selected
            try:
                self.ReportColumnGrid.ScrollIntoView(selected)
            except Exception:
                pass
        self.report_column_selection_changed(None, None)

    def _rebuild_report_columns(self, reset=False):
        saved = [] if reset else self._saved_column_schema
        self._report_columns = build_report_column_options(
            self._snapshot.get("revision_parameter_names") or [],
            self._snapshot.get("cloud_parameter_names") or [],
            saved_schema=saved,
        )
        if not saved and self._legacy_parameter_names:
            legacy_names = set(self._legacy_parameter_names)
            for column in self._report_columns:
                if column.source == "Revision Cloud" and column.parameter_name in legacy_names:
                    column.is_selected = True
        self._refresh_report_column_grid()

    def _set_status(self, value):
        try:
            self.StatusText.Text = text(value)
        except Exception:
            pass

    def _apply_snapshot(self, snapshot, reset_report_selection=True):
        self._snapshot = dict(snapshot or {})
        self._clouds = list(self._snapshot.get("clouds") or [])
        incoming_options = list(self._snapshot.get("revisions") or [])
        if not reset_report_selection and self._revision_options:
            selected_ids = selected_revision_ids(self._revision_options)
            for option in incoming_options:
                option.is_selected = option.revision_id in selected_ids
        self._revision_options = incoming_options
        self._metadata = dict(self._snapshot.get("metadata") or {})
        self._document_key = text(self._snapshot.get("document_key"))
        self._gateway.document_key = self._document_key
        self._set_stale_visual_state(False)
        self.DocumentText.Text = "Active document: {}".format(text(self._snapshot.get("document_label"), "-"))
        self._populate_filter_controls()
        self.RevisionChecklist.ItemsSource = None
        self.RevisionChecklist.ItemsSource = self._revision_options
        self._rebuild_report_columns(reset=False)
        if bool(self.UseProjectValuesCheck.IsChecked):
            self._load_project_metadata_fields()
        self._discard_comment_draft(clear_selection=True)

    def _populate_filter_controls(self):
        current_revision = text(getattr(self.RevisionFilter, "SelectedItem", "All revisions"), "All revisions")
        current_issue = text(getattr(self.IssueFilter, "SelectedItem", "All issue status"), "All issue status")
        current_placement = text(getattr(self.PlacementFilter, "SelectedItem", "All placement"), "All placement")
        current_sheet = text(getattr(self.SheetFilter, "SelectedItem", "All sheets"), "All sheets")

        revisions = ["All revisions"] + sorted(set([row.revision_number for row in self._clouds]))
        issues = ["All issue status", "Issues only", "No issues", "Missing Comment", "Review Placement", "Not on Sheet"]
        placements = ["All placement", "On Sheet", "In View", "Not on Sheet"]
        sheets = ["All sheets"] + sorted(set([
            sheet for row in self._clouds for sheet in row.sheet_numbers_list if text(sheet)
        ]))
        self._assign_combo(self.RevisionFilter, revisions, current_revision)
        self._assign_combo(self.IssueFilter, issues, current_issue)
        self._assign_combo(self.PlacementFilter, placements, current_placement)
        self._assign_combo(self.SheetFilter, sheets, current_sheet)

    def _assign_combo(self, combo, values, preferred):
        combo.ItemsSource = list(values)
        try:
            combo.SelectedItem = preferred if preferred in values else values[0]
        except Exception:
            combo.SelectedIndex = 0

    def _selected_rows(self):
        try:
            return list(self.CloudGrid.SelectedItems or [])
        except Exception:
            return []

    def _selected_report_state(self):
        revision_ids = selected_revision_ids(self._revision_options)
        deduplicate = bool(self.DeduplicateCheck.IsChecked)
        rows = build_report_rows(self._clouds, revision_ids, deduplicate=deduplicate)
        if not bool(self.ShowBlankCommentRowsCheck.IsChecked):
            rows = [row for row in rows if text(row.get("comments")).strip()]
        groups = group_report_rows(rows)
        existing_ids = set([group["revision_id"] for group in groups])
        for option in self._revision_options:
            if option.is_selected and option.revision_id not in existing_ids:
                groups.append({
                    "revision_id": option.revision_id,
                    "number": option.number,
                    "date": option.date,
                    "description": option.description,
                    "sequence": option.sequence,
                    "rows": [],
                })
        groups = sorted(groups, key=lambda group: group["sequence"])
        columns = selected_report_columns(self._report_columns)
        logo_path = LOGO_PATH if bool(self.IncludeLogoCheck.IsChecked) and os.path.isfile(LOGO_PATH) else None
        return (
            self._current_metadata(), groups, columns, logo_path,
            self._page_orientation(), self._report_font_name(),
            self._report_table_font_size())

    def _current_metadata(self):
        return {
            "project_number": text(self.ProjectNumberBox.Text),
            "client": text(self.ClientBox.Text),
            "project_name": text(self.ProjectNameBox.Text),
            "report_date": text(self.ReportDateBox.Text),
        }

    def _load_project_metadata_fields(self):
        self._updating_report = True
        try:
            self.ProjectNumberBox.Text = text(self._metadata.get("project_number"))
            self.ClientBox.Text = text(self._metadata.get("client"))
            self.ProjectNameBox.Text = text(self._metadata.get("project_name"))
            self.ReportDateBox.Text = text(self._metadata.get("report_date"))
        finally:
            self._updating_report = False

    def _set_metadata_editable(self, editable):
        for box in (self.ProjectNumberBox, self.ClientBox, self.ProjectNameBox, self.ReportDateBox):
            box.IsReadOnly = not bool(editable)
            self._set_textbox_style(box, "CED.Input.TextBox" if editable else "CED.Input.TextBox.ReadOnly")

    def _discard_comment_draft(self, clear_selection=False):
        self._setting_comment = True
        try:
            self._comment_original = ""
            self._comment_cloud_id = 0
            self._comment_dirty = False
            self.CommentStateText.Text = ""
            self.CommentBox.Text = ""
            self.CommentBox.IsReadOnly = True
            self._set_textbox_style(self.CommentBox, "CED.Input.TextBox.ReadOnly")
            self.CloudIdBox.Text = ""
            self.SelectedRevisionBox.Text = ""
            self.OwnerBox.Text = ""
            self.RevitOwnerBox.Text = ""
            self.CreatedByBox.Text = ""
            if clear_selection:
                try:
                    self.CloudGrid.SelectedItems.Clear()
                except Exception:
                    pass
        finally:
            self._setting_comment = False

    def _load_selection_inspector(self):
        rows = self._selected_rows()
        self._queue = list(self._filtered_clouds)
        if rows:
            try:
                self._queue_index = self._queue.index(rows[0])
            except ValueError:
                self._queue_index = 0
        else:
            self._queue_index = 0
        self._setting_comment = True
        try:
            self._comment_dirty = False
            self._comment_original = ""
            self._comment_cloud_id = 0
            if len(rows) == 1:
                row = rows[0]
                self.CloudIdBox.Text = text(row.cloud_id)
                self.SelectedRevisionBox.Text = text(row.revision_display)
                self.OwnerBox.Text = row.owner_display
                self.RevitOwnerBox.Text = row.worksharing_owner_display
                self.CreatedByBox.Text = row.created_by_display
                self.CommentBox.Text = row.comments
                self._comment_original = row.comments
                self._comment_cloud_id = row.cloud_id
                comment_locked = bool(self._stale or row.owned_by_other)
                self.CommentBox.IsReadOnly = comment_locked
                if row.owned_by_other:
                    self.CommentStateText.Text = "Comment editing is unavailable while another Revit user owns this cloud."
                elif self._stale:
                    self.CommentStateText.Text = "Refresh required before editing."
                else:
                    self.CommentStateText.Text = ""
                self._set_textbox_style(
                    self.CommentBox,
                    "CED.Input.TextBox.ReadOnly" if comment_locked else "CED.Input.TextBox")
            elif len(rows) > 1:
                self.CloudIdBox.Text = "{} selected".format(len(rows))
                self.SelectedRevisionBox.Text = "Multiple"
                self.OwnerBox.Text = "Multiple sheets / views"
                self.RevitOwnerBox.Text = "Multiple"
                self.CreatedByBox.Text = "Multiple"
                self.CommentBox.Text = "Comments can be edited only when one cloud row is selected."
                self.CommentBox.IsReadOnly = True
                self.CommentStateText.Text = "Select one cloud to edit its comment."
                self._set_textbox_style(self.CommentBox, "CED.Input.TextBox.ReadOnly")
            else:
                self.CloudIdBox.Text = ""
                self.SelectedRevisionBox.Text = ""
                self.OwnerBox.Text = ""
                self.RevitOwnerBox.Text = ""
                self.CreatedByBox.Text = ""
                self.CommentBox.Text = ""
                self.CommentBox.IsReadOnly = True
                self.CommentStateText.Text = ""
                self._set_textbox_style(self.CommentBox, "CED.Input.TextBox.ReadOnly")
        finally:
            self._setting_comment = False
        self._update_queue_text()
        self._update_action_states()

    def _update_queue_text(self):
        count = len(self._queue)
        if count and self._selected_rows():
            self.QueueText.Text = "{} of {}".format(self._queue_index + 1, count)
        elif count:
            self.QueueText.Text = "{} visible".format(count)
        else:
            self.QueueText.Text = "No visible rows"

    def _update_action_states(self):
        selected_count = len(self._selected_rows())
        busy = self._gateway.is_busy()
        enabled = not self._stale and not busy
        self.ShowButton.IsEnabled = bool(selected_count and enabled)
        self.SelectButton.IsEnabled = bool(selected_count and enabled)
        can_review = bool(selected_count and len(self._filtered_clouds) > 1 and enabled)
        self.PreviousButton.IsEnabled = can_review and self._queue_index > 0
        self.NextButton.IsEnabled = can_review and self._queue_index < len(self._filtered_clouds) - 1
        self.ApplyCommentButton.IsEnabled = bool(
            selected_count == 1 and self._comment_dirty and enabled and not self.CommentBox.IsReadOnly)
        self.RefreshButton.IsEnabled = not busy
        self.StaleRefreshButton.IsEnabled = not busy

    def _apply_filters(self):
        if self._initializing:
            return
        self._filtered_clouds = filter_clouds(
            self._clouds,
            search=text(self.SearchBox.Text),
            revision=text(self.RevisionFilter.SelectedItem, "All revisions"),
            issue=text(self.IssueFilter.SelectedItem, "All issue status"),
            placement=text(self.PlacementFilter.SelectedItem, "All placement"),
            sheet=text(self.SheetFilter.SelectedItem, "All sheets"),
        )
        self.CloudGrid.ItemsSource = None
        self.CloudGrid.ItemsSource = self._filtered_clouds
        issue_count = len([row for row in self._filtered_clouds if row.has_issue])
        self.CloudCountText.Text = "{} Clouds".format(len(self._filtered_clouds))
        self.IssueCountText.Text = "{} Issues".format(issue_count)
        self.SelectedCountText.Text = "0 Selected"
        self._discard_comment_draft(clear_selection=False)
        self._queue = list(self._filtered_clouds)
        self._queue_index = 0
        self._update_queue_text()
        self._update_action_states()

    def _update_report_preview(self):
        if self._initializing or self._updating_report:
            return
        # Decide the warning before building, because the table rows are shaded
        # only while the toast is on screen.
        self._evaluate_issue_warning()
        metadata, groups, columns, logo_path, orientation, font_name, table_font_size = self._selected_report_state()
        document = FlowDocument()
        page_width = 816.0 if orientation == "portrait" else 1056.0
        page_height = 1056.0 if orientation == "portrait" else 816.0
        self._preview_page_width = page_width
        self._preview_page_height = page_height
        page_padding = 36.0
        content_width = page_width - (page_padding * 2.0)
        document.PageWidth = page_width
        document.PageHeight = page_height
        document.ColumnWidth = content_width
        document.PagePadding = Thickness(page_padding)
        document.FontFamily = FontFamily(font_name)
        document.FontSize = table_font_size * (96.0 / 72.0)
        document.Background = Brushes.White
        document.Foreground = Brushes.Black

        header = Table()
        header.CellSpacing = 0
        header.Columns.Add(TableColumn())
        header.Columns[0].Width = GridLength(content_width - (175.0 if logo_path else 0.0))
        if logo_path:
            header.Columns.Add(TableColumn())
            header.Columns[1].Width = GridLength(175.0)
        header_group = TableRowGroup()
        header.RowGroups.Add(header_group)
        header_row = TableRow()
        header_group.Rows.Add(header_row)
        title_cell = TableCell()
        title = Paragraph()
        title.Margin = Thickness(0, 0, 0, 2)
        title_run = Run("Project Revision Summary")
        title_run.FontSize = 18
        title_run.FontWeight = FontWeights.Bold
        title_run.Foreground = REPORT_BLUE_BRUSH
        title.Inlines.Add(title_run)
        title_cell.Blocks.Add(title)
        company = Paragraph(Run("CoolSys Energy Design"))
        company.Margin = Thickness(0, 0, 0, 10)
        company.FontWeight = FontWeights.Bold
        title_cell.Blocks.Add(company)
        header_row.Cells.Add(title_cell)
        if logo_path:
            try:
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.CacheOption = BitmapCacheOption.OnLoad
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                logo = Image()
                logo.Source = bitmap
                logo.Width = 155
                logo.Height = 44
                logo.Stretch = Stretch.Uniform
                logo.HorizontalAlignment = HorizontalAlignment.Right
                logo_container = BlockUIContainer(logo)
                logo_container.Margin = Thickness(0)
                logo_cell = TableCell()
                logo_cell.Blocks.Add(logo_container)
                header_row.Cells.Add(logo_cell)
            except Exception:
                header_row.Cells.Add(TableCell(Paragraph()))
        document.Blocks.Add(header)
        for label, key in (
            ("Project Number", "project_number"),
            ("Client", "client"),
            ("Project Name", "project_name"),
            ("Report Date", "report_date"),
        ):
            paragraph = Paragraph()
            paragraph.Margin = Thickness(0, 1, 0, 1)
            paragraph.TextAlignment = TextAlignment.Left
            label_run = Run("{}: ".format(label))
            label_run.FontWeight = FontWeights.Bold
            paragraph.Inlines.Add(label_run)
            paragraph.Inlines.Add(Run(text(metadata.get(key))))
            document.Blocks.Add(paragraph)

        if not columns:
            empty_columns = Paragraph(Run("Select at least one report column or parameter."))
            empty_columns.Margin = Thickness(0, 20, 0, 0)
            empty_columns.TextAlignment = TextAlignment.Left
            document.Blocks.Add(empty_columns)
        elif not groups:
            empty = Paragraph(Run("Select at least one revision to preview a report."))
            empty.Margin = Thickness(0, 24, 0, 0)
            empty.FontStyle = self.FontStyle
            document.Blocks.Add(empty)
        for group in groups:
            heading = Paragraph()
            heading.Margin = Thickness(0, 22, 0, 7)
            heading.TextAlignment = TextAlignment.Left
            heading_run = Run("Revision {} | Date: {} | Description: {}".format(
                group["number"], group["date"], group["description"]))
            heading_run.FontWeight = FontWeights.Bold
            heading_run.FontSize = 14
            heading_run.Foreground = REPORT_BLUE_BRUSH
            heading.Inlines.Add(heading_run)
            document.Blocks.Add(heading)
            if columns:
                document.Blocks.Add(self._preview_table(group["rows"], columns, content_width))
        self._preview_document = document
        self._preview_page_index = 0
        self._refresh_preview_pages()
        self._update_preview_issue_toast()

    def _evaluate_issue_warning(self):
        """Work out whether the report covers clouds that still have issues.

        Runs before the FlowDocument is built so _preview_table knows whether to
        shade the offending rows. The report exports either way - this is a "you
        may not want to issue this yet" nudge, not a block. Keyed on the revision
        selection plus the flagged count, so dismissing sticks for that exact
        state while a new selection or a newly resolved cloud brings it back.
        """
        self._issue_warning_active = False
        self._issue_warning_text = ""
        try:
            revision_ids = set(selected_revision_ids(self._revision_options))
            flagged = [row for row in self._clouds
                       if row.revision_id in revision_ids and row.has_issue]
            self._issue_toast_key = (tuple(sorted(revision_ids)), len(flagged))
            if not flagged or self._issue_toast_key == self._issue_toast_dismissed_key:
                return
            self._issue_warning_text = (
                "1 cloud in the selected revisions still has an issue. Please review."
                if len(flagged) == 1 else
                "{} clouds in the selected revisions still have issues. Please review.".format(
                    len(flagged)))
            self._issue_warning_active = True
        except Exception:
            self._issue_warning_active = False
            self._issue_warning_text = ""

    def _update_preview_issue_toast(self):
        try:
            if self._issue_warning_active:
                self.PreviewIssueToastText.Text = self._issue_warning_text
                self.PreviewIssueToast.Visibility = Visibility.Visible
            else:
                self.PreviewIssueToast.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def dismiss_preview_issue_toast_clicked(self, sender, args):
        # Dismissing also drops the row shading, so the document is rebuilt.
        self._issue_toast_dismissed_key = self._issue_toast_key
        self.PreviewIssueToast.Visibility = Visibility.Collapsed
        self._update_report_preview()

    def _preview_table(self, rows, columns, available_width):
        table = Table()
        table.CellSpacing = 0
        table.BorderBrush = Brushes.LightSlateGray
        table.BorderThickness = Thickness(0.5)
        usable_width = max(1.0, float(available_width) - 2.0)
        width_weights = [max(0.5, min(4.0, float(column.width_weight))) for column in columns]
        total_weight = sum(width_weights) or 1.0
        for column, weight in zip(columns, width_weights):
            table_column = TableColumn()
            table_column.Width = GridLength(usable_width * float(weight) / total_weight)
            table.Columns.Add(table_column)
        group = TableRowGroup()
        table.RowGroups.Add(group)
        header_row = TableRow()
        header_row.Background = Brushes.LightSteelBlue
        group.Rows.Add(header_row)
        for column_index, column in enumerate(columns):
            header_paragraph = Paragraph(Run(column.label))
            header_paragraph.TextAlignment = TextAlignment.Left
            cell = TableCell(header_paragraph)
            cell.FontWeight = FontWeights.Bold
            cell.Padding = Thickness(5)
            cell.BorderBrush = Brushes.LightSlateGray
            right_border = 0.75 if column_index == len(columns) - 1 else 0.25
            cell.BorderThickness = Thickness(0.25, 0.25, right_border, 0.25)
            header_row.Cells.Add(cell)
        if not rows:
            empty_paragraph = Paragraph(Run("No revision clouds found for this revision."))
            empty_paragraph.TextAlignment = TextAlignment.Left
            cell = TableCell(empty_paragraph)
            cell.ColumnSpan = len(columns)
            cell.Padding = Thickness(5)
            cell.BorderBrush = Brushes.LightSlateGray
            cell.BorderThickness = Thickness(0.25, 0.25, 0.75, 0.25)
            group.Rows.Add(TableRow())
            group.Rows[group.Rows.Count - 1].Cells.Add(cell)
        for item in rows:
            row = TableRow()
            # While the warning toast is up, tint the rows it is talking about so
            # "4 clouds have issues" points at something. Cleared on dismissal.
            if self._issue_warning_active and text(item.get("issues")).strip():
                row.Background = REPORT_ISSUE_ROW_BRUSH
            group.Rows.Add(row)
            for column_index, column in enumerate(columns):
                value = report_column_value(item, column)
                cell_paragraph = Paragraph(Run(text(value)))
                cell_paragraph.TextAlignment = TextAlignment.Left
                cell = TableCell(cell_paragraph)
                cell.Padding = Thickness(5)
                cell.BorderBrush = Brushes.LightSlateGray
                right_border = 0.75 if column_index == len(columns) - 1 else 0.25
                cell.BorderThickness = Thickness(0.25, 0.25, right_border, 0.25)
                cell.TextAlignment = TextAlignment.Left
                if bool(getattr(column, "is_bold", False)):
                    cell.FontWeight = FontWeights.Bold
                if column.key == "issues" and text(value):
                    cell.Foreground = Brushes.Firebrick
                row.Cells.Add(cell)
        return table

    def _set_preview_zoom(self, value):
        try:
            self._preview_zoom = max(25.0, min(400.0, float(value)))
            if abs(float(self.PreviewZoomSlider.Value) - self._preview_zoom) > 0.01:
                self.PreviewZoomSlider.Value = self._preview_zoom
            self.PreviewZoomText.Text = "{}%".format(int(round(self._preview_zoom)))
            self._size_preview_page()
        except Exception:
            pass

    def _document_paginator(self):
        if self._preview_document is None:
            return None
        try:
            return clr.Convert(self._preview_document, IDocumentPaginatorSource).DocumentPaginator
        except Exception:
            try:
                return self._preview_document.DocumentPaginator
            except Exception:
                return None

    def _refresh_preview_pages(self):
        paginator = self._document_paginator()
        if paginator is None:
            self._preview_page_count = 0
            self.PreviewPageText.Text = "0 of 0"
            return
        try:
            paginator.ComputePageCount()
            self._preview_page_count = max(1, int(paginator.PageCount))
        except Exception as ex:
            LOGGER.debug("Revision Manager preview pagination unavailable: %s", ex)
            self._preview_page_count = 1
        self._preview_page_index = max(0, min(
            int(self._preview_page_index), self._preview_page_count - 1))
        try:
            self.PreviewPageView.DocumentPaginator = paginator
            self.PreviewPageView.PageNumber = self._preview_page_index
        except Exception as ex:
            LOGGER.debug("Revision Manager preview page unavailable: %s", ex)
        self._size_preview_page()
        self._update_preview_navigation()

    def _size_preview_page(self):
        if not hasattr(self, "PreviewPageView"):
            return
        scale = max(0.25, min(4.0, float(self._preview_zoom) / 100.0))
        try:
            self.PreviewPageView.Width = self._preview_page_width
            self.PreviewPageView.Height = self._preview_page_height
            self.PreviewPageBorder.Width = self._preview_page_width + 2.0
            self.PreviewPageBorder.Height = self._preview_page_height + 2.0
            self.PreviewScaleTransform.ScaleX = scale
            self.PreviewScaleTransform.ScaleY = scale

            scaled_width = (self._preview_page_width + 2.0) * scale
            scaled_height = (self._preview_page_height + 2.0) * scale
            viewport_width = max(0.0, float(self.PreviewScroll.ViewportWidth))
            viewport_height = max(0.0, float(self.PreviewScroll.ViewportHeight))
            extent_width = max(viewport_width, scaled_width + 36.0)
            extent_height = max(viewport_height, scaled_height + 36.0)
            left = max(18.0, (extent_width - scaled_width) / 2.0)
            top = max(18.0, (extent_height - scaled_height) / 2.0)
            self.PreviewExtent.Width = extent_width
            self.PreviewExtent.Height = extent_height
            self.PreviewScaleHost.Margin = Thickness(left, top, 0, 0)
        except Exception:
            pass

    def _update_preview_navigation(self):
        count = max(0, int(self._preview_page_count))
        current = self._preview_page_index + 1 if count else 0
        self.PreviewPageText.Text = "{} of {}".format(current, count)
        self.PreviewPreviousPageButton.IsEnabled = bool(count and self._preview_page_index > 0)
        self.PreviewNextPageButton.IsEnabled = bool(
            count and self._preview_page_index < count - 1)

    def _show_preview_page(self, page_index):
        if self._preview_page_count <= 0:
            return
        self._preview_page_index = max(0, min(int(page_index), self._preview_page_count - 1))
        try:
            self.PreviewPageView.PageNumber = self._preview_page_index
            self.PreviewScroll.ScrollToHome()
        except Exception:
            pass
        self._update_preview_navigation()

    def _set_stale_visual_state(self, stale, message=None):
        self._stale = bool(stale)
        visibility = Visibility.Visible if self._stale else Visibility.Collapsed
        self.StaleOverlay.Visibility = visibility
        self.StaleToast.Visibility = visibility
        self.MainTabs.IsEnabled = not self._stale
        self.ExportButtons.IsEnabled = not self._stale
        if message:
            self.StaleToastMessage.Text = text(message)

    def _mark_stale(self, active_label, message):
        if self._stale:
            return
        self._set_stale_visual_state(True, message)
        self.DocumentText.Text = "Active document: {} | Refresh required".format(text(active_label, "-"))
        self._set_status(message)
        self._load_selection_inspector()
        self._update_action_states()

    def _active_document_changed(self, active_key, active_label):
        different = text(active_key) != self._document_key
        if different:
            self._mark_stale(
                active_label,
                "The active Revit document changed. Refresh Revision Manager to load its current revision-cloud data.")
        elif not self._stale:
            self.DocumentText.Text = "Active document: {}".format(text(active_label, "-"))
            self._set_status("Ready")
        self._update_action_states()

    def _document_edited(self, document_key, document_label):
        if text(document_key) != self._document_key:
            return
        self._mark_stale(
            document_label,
            "Model changes were detected. Refresh Revision Manager to display current revision-cloud data.")

    def _operation_started(self, message):
        self._set_status(message)
        self._update_action_states()

    def _operation_failed(self, error):
        message = text(error, "The operation failed.")
        self._set_status(message)
        forms.alert(message, title=TITLE)
        self._update_action_states()

    def _current_queue_row(self):
        if not self._queue:
            return None
        self._queue_index = max(0, min(self._queue_index, len(self._queue) - 1))
        return self._queue[self._queue_index]

    def _show_queue_row(self):
        row = self._current_queue_row()
        if row is None:
            return
        if self._gateway.raise_show(row.cloud_id, row.owner_view_id, self._model_operation_complete):
            self._operation_started("Opening revision cloud {}...".format(row.cloud_id))
        else:
            self._set_status("Another Revision Manager action is already pending.")

    def _model_operation_complete(self, status, operation, result, error):
        if status != "ok":
            self._operation_failed(error)
            return
        if operation == "show":
            self._set_status("Revision cloud shown in model.")
        elif operation == "select":
            selected = int((result or {}).get("selected") or 0)
            missing = int((result or {}).get("missing") or 0)
            self._set_status("Selected {} cloud(s); {} missing.".format(selected, missing))
        elif operation == "navigate":
            self._set_status("Opened {}.".format(text((result or {}).get("view_name"), "view or sheet")))
        self._update_action_states()

    def filters_changed(self, sender, args):
        self._apply_filters()

    def clear_filters_clicked(self, sender, args):
        self._initializing = True
        try:
            self.SearchBox.Text = ""
            self.RevisionFilter.SelectedIndex = 0
            self.IssueFilter.SelectedIndex = 0
            self.PlacementFilter.SelectedIndex = 0
            self.SheetFilter.SelectedIndex = 0
        finally:
            self._initializing = False
        self._apply_filters()

    def review_column_toggled(self, sender, args):
        if self._initializing:
            return
        self._set_review_column_visibility(self._selected_review_column_keys())
        self._save_report_preferences()

    def cloud_selection_changed(self, sender, args):
        if self._initializing:
            return
        rows = self._selected_rows()
        self.SelectedCountText.Text = "{} Selected".format(len(rows))
        self._load_selection_inspector()

    def location_link_double_clicked(self, sender, args):
        try:
            if int(args.ClickCount) != 2:
                return
            target_id = int(getattr(sender, "Tag", 0) or 0)
        except Exception:
            return
        if target_id <= 0 or self._stale:
            return
        if self._gateway.raise_navigate(target_id, self._model_operation_complete):
            self._operation_started("Opening view or sheet...")
        else:
            self._set_status("Another Revision Manager action is already pending.")

    def comment_text_changed(self, sender, args):
        if self._setting_comment or self._initializing:
            return
        if len(self._selected_rows()) != 1 or self.CommentBox.IsReadOnly:
            return
        self._comment_dirty = text(self.CommentBox.Text) != self._comment_original
        self.CommentStateText.Text = "Unsaved comment change" if self._comment_dirty else ""
        self._set_textbox_style(self.CommentBox, "CommentChangedTextBox" if self._comment_dirty else "CED.Input.TextBox")
        self._update_action_states()

    def apply_comment_clicked(self, sender, args):
        rows = self._selected_rows()
        if len(rows) != 1 or rows[0].cloud_id != self._comment_cloud_id:
            self._load_selection_inspector()
            return
        proposed = text(self.CommentBox.Text)
        if self._gateway.raise_update_comment(
                self._comment_cloud_id, self._comment_original, proposed, self._comment_update_complete):
            self._operation_started("Applying revision cloud comment...")
        else:
            self._set_status("Another Revision Manager action is already pending.")

    def _comment_update_complete(self, status, operation, result, error):
        if status != "ok":
            self._operation_failed(error)
            return
        cloud_id = int((result or {}).get("cloud_id") or 0)
        comment = text((result or {}).get("comment"))
        for row in self._clouds:
            if row.cloud_id == cloud_id:
                row.update_comment(comment)
                break
        self._comment_original = comment
        self._comment_dirty = False
        self.CommentStateText.Text = ""
        self._set_textbox_style(self.CommentBox, "CED.Input.TextBox")
        self._apply_filters()
        for row in self._filtered_clouds:
            if row.cloud_id == cloud_id:
                try:
                    self.CloudGrid.SelectedItem = row
                    self.CloudGrid.ScrollIntoView(row)
                except Exception:
                    pass
                break
        self._update_report_preview()
        self._set_status("Revision cloud comment updated.")
        self._update_action_states()

    def show_clicked(self, sender, args):
        self._show_queue_row()

    def select_clicked(self, sender, args):
        ids = [row.cloud_id for row in self._selected_rows()]
        if self._gateway.raise_select(ids, self._model_operation_complete):
            self._operation_started("Selecting revision clouds...")
        else:
            self._set_status("Another Revision Manager action is already pending.")

    def previous_clicked(self, sender, args):
        if not self._queue:
            return
        self._select_review_row(max(0, self._queue_index - 1))

    def next_clicked(self, sender, args):
        if not self._queue:
            return
        self._select_review_row(min(len(self._queue) - 1, self._queue_index + 1))

    def _select_review_row(self, index):
        if not self._queue:
            return
        self._queue_index = max(0, min(int(index), len(self._queue) - 1))
        row = self._queue[self._queue_index]
        try:
            self.CloudGrid.SelectedItems.Clear()
            self.CloudGrid.SelectedItem = row
            self.CloudGrid.ScrollIntoView(row)
        except Exception:
            pass

    def refresh_clicked(self, sender, args):
        if self._gateway.raise_refresh(self._refresh_complete):
            self._operation_started("Refreshing revision cloud data...")
        else:
            self._set_status("Another Revision Manager action is already pending.")

    def _refresh_complete(self, status, operation, result, error):
        if status != "ok":
            self._operation_failed(error)
            return
        document_changed = text((result or {}).get("document_key")) != self._document_key
        self._initializing = True
        try:
            if document_changed:
                self.UseProjectValuesCheck.IsChecked = True
            self._apply_snapshot(result, reset_report_selection=True)
            self._set_metadata_editable(not bool(self.UseProjectValuesCheck.IsChecked))
        finally:
            self._initializing = False
        self._apply_filters()
        self._update_report_preview()
        self._set_status("Revision cloud data refreshed.")
        self._update_action_states()

    def project_values_toggled(self, sender, args):
        if self._initializing:
            return
        use_project = bool(self.UseProjectValuesCheck.IsChecked)
        if use_project:
            self._load_project_metadata_fields()
        self._set_metadata_editable(not use_project)
        self._update_report_preview()

    def report_revision_toggled(self, sender, args):
        if self._initializing:
            return
        try:
            revision_id = int(sender.Tag)
            is_selected = bool(sender.IsChecked)
        except Exception:
            return
        for option in self._revision_options:
            if option.revision_id == revision_id:
                option.is_selected = is_selected
                break
        self._update_report_preview()

    def select_all_revisions_clicked(self, sender, args):
        for option in self._revision_options:
            option.is_selected = True
        self.RevisionChecklist.ItemsSource = None
        self.RevisionChecklist.ItemsSource = self._revision_options
        self._update_report_preview()

    def clear_revisions_clicked(self, sender, args):
        for option in self._revision_options:
            option.is_selected = False
        self.RevisionChecklist.ItemsSource = None
        self.RevisionChecklist.ItemsSource = self._revision_options
        self._update_report_preview()

    def report_column_toggled(self, sender, args):
        if self._initializing:
            return
        key = text(getattr(sender, "Tag", ""))
        selected = bool(getattr(sender, "IsChecked", False))
        for column in self._report_columns:
            if column.key == key:
                column.is_selected = selected
                break
        selected_column = None
        for column in self._report_columns:
            if column.key == key:
                selected_column = column
                break
        self._refresh_report_column_grid(selected=selected_column)
        self._save_report_preferences()
        self._update_report_preview()

    def report_column_filter_changed(self, sender, args):
        if self._initializing:
            return
        self._refresh_report_column_grid()
        self._save_report_preferences()

    def _move_report_column(self, delta):
        selected = self.ReportColumnGrid.SelectedItem
        if selected is None:
            self._set_status("Select a report column to reorder it.")
            return
        visible = self._visible_report_columns()
        try:
            visible_index = visible.index(selected)
        except ValueError:
            return
        target_visible_index = visible_index + int(delta)
        if target_visible_index < 0 or target_visible_index >= len(visible):
            return
        target_item = visible[target_visible_index]
        index = self._report_columns.index(selected)
        target = self._report_columns.index(target_item)
        self._report_columns[index], self._report_columns[target] = (
            self._report_columns[target], self._report_columns[index])
        self._refresh_report_column_grid(selected=selected)
        self._save_report_preferences()
        self._update_report_preview()

    def move_column_up_clicked(self, sender, args):
        self._move_report_column(-1)

    def move_column_down_clicked(self, sender, args):
        self._move_report_column(1)

    def reset_columns_clicked(self, sender, args):
        self._rebuild_report_columns(reset=True)
        self._save_report_preferences()
        self._update_report_preview()

    def report_preference_changed(self, sender, args):
        if self._initializing:
            return
        self._save_report_preferences()
        self._update_report_preview()

    def report_format_changed(self, sender, args):
        if self._initializing:
            return
        self._save_report_preferences()
        self._update_report_preview()

    def report_column_selection_changed(self, sender, args):
        if not hasattr(self, "SelectedColumnWidthLabel"):
            return
        column = self.ReportColumnGrid.SelectedItem
        enabled = column is not None
        self._updating_column_editor = True
        try:
            self.SelectedColumnWidthSlider.IsEnabled = enabled
            self.ColumnWidthResetButton.IsEnabled = enabled
            self.SelectedColumnBoldCheck.IsEnabled = enabled
            if not enabled:
                self.SelectedColumnWidthLabel.Text = "Select a column"
                self.SelectedColumnWidthText.Text = "—"
                self.SelectedColumnBoldCheck.IsChecked = False
                return
            self.SelectedColumnWidthLabel.Text = text(column.label)
            self.SelectedColumnWidthSlider.Value = float(column.width_weight)
            self.SelectedColumnBoldCheck.IsChecked = bool(column.is_bold)
            self._update_selected_column_width_text(column.width_weight)
        finally:
            self._updating_column_editor = False

    def _update_selected_column_width_text(self, value):
        self.SelectedColumnWidthText.Text = "{:.2f}x".format(float(value))

    def report_column_width_slider_changed(self, sender, args):
        if self._initializing or self._updating_column_editor:
            return
        column = self.ReportColumnGrid.SelectedItem
        if column is None:
            return
        column.width_weight = max(0.5, min(4.0, round(float(sender.Value), 2)))
        self._update_selected_column_width_text(column.width_weight)
        self._update_report_preview()

    def report_column_width_committed(self, sender, args):
        if self._initializing or self._updating_column_editor:
            return
        if self.ReportColumnGrid.SelectedItem is not None:
            self._save_report_preferences()

    def column_width_reset_clicked(self, sender, args):
        column = self.ReportColumnGrid.SelectedItem
        if column is None:
            return
        column.width_weight = default_report_column_width(column.key)
        self.report_column_selection_changed(None, None)
        self._save_report_preferences()
        self._update_report_preview()

    def report_column_bold_toggled(self, sender, args):
        if self._initializing or self._updating_column_editor:
            return
        column = self.ReportColumnGrid.SelectedItem
        if column is None:
            return
        column.is_bold = bool(sender.IsChecked)
        self._save_report_preferences()
        self._update_report_preview()

    def preview_zoom_slider_changed(self, sender, args):
        if self._initializing or not hasattr(self, "PreviewPageView"):
            return
        self._set_preview_zoom(sender.Value)

    def preview_zoom_out_clicked(self, sender, args):
        self._set_preview_zoom(self._preview_zoom - 10.0)

    def preview_zoom_in_clicked(self, sender, args):
        self._set_preview_zoom(self._preview_zoom + 10.0)

    def preview_zoom_fit_clicked(self, sender, args):
        try:
            available_height = max(100.0, float(self.PreviewScroll.ViewportHeight) - 38.0)
            self._set_preview_zoom(100.0 * available_height / float(self._preview_page_height))
        except Exception:
            self._set_preview_zoom(75.0)

    def preview_zoom_fit_width_clicked(self, sender, args):
        try:
            available_width = max(100.0, float(self.PreviewScroll.ViewportWidth) - 38.0)
            self._set_preview_zoom(100.0 * available_width / float(self._preview_page_width))
        except Exception:
            self._set_preview_zoom(75.0)

    def preview_previous_page_clicked(self, sender, args):
        self._show_preview_page(self._preview_page_index - 1)

    def preview_next_page_clicked(self, sender, args):
        self._show_preview_page(self._preview_page_index + 1)

    def preview_host_size_changed(self, sender, args):
        self._size_preview_page()

    def page_orientation_changed(self, sender, args):
        if self._initializing:
            return
        self._save_report_preferences()
        self._update_report_preview()

    def report_input_changed(self, sender, args):
        self._update_report_preview()

    def main_tab_changed(self, sender, args):
        if not hasattr(self, "ExportButtons"):
            return
        report_tab = int(getattr(self.MainTabs, "SelectedIndex", 0) or 0) == 1
        self.ExportButtons.Visibility = Visibility.Visible if report_tab else Visibility.Collapsed
        if report_tab:
            self._update_report_preview()

    def _save_path(self, extension, default_name, description):
        dialog = SaveFileDialog()
        dialog.DefaultExt = extension
        dialog.AddExtension = True
        dialog.FileName = default_name
        dialog.Filter = "{} (*.{})|*.{}".format(description, extension, extension)
        if dialog.ShowDialog() == DialogResult.OK:
            return text(dialog.FileName)
        return ""

    def export_excel_clicked(self, sender, args):
        metadata, groups, columns, logo_path, unused_orientation, font_name, unused_table_size = self._selected_report_state()
        if not columns:
            forms.alert("Select at least one report column before exporting.", title=TITLE)
            return
        path = self._save_path("xlsx", "Revision_Cloud_Report.xlsx", "Excel Workbook")
        if not path:
            return
        try:
            export_xlsx(path, metadata, groups, columns, logo_path=logo_path, font_name=font_name)
            self._set_status("Excel report saved: {}".format(path))
        except Exception as ex:
            self._operation_failed(ex)

    def export_pdf_clicked(self, sender, args):
        metadata, groups, columns, logo_path, orientation, font_name, table_font_size = self._selected_report_state()
        if not columns:
            forms.alert("Select at least one report column before exporting.", title=TITLE)
            return
        path = self._save_path("pdf", "Revision_Cloud_Report.pdf", "PDF Document")
        if not path:
            return
        try:
            export_pdf(
                path, metadata, groups, columns, logo_path=logo_path,
                orientation=orientation, font_name=font_name,
                table_font_size=table_font_size)
            self._set_status("PDF report saved: {}".format(path))
        except Exception as ex:
            self._operation_failed(ex)

    def copy_word_clicked(self, sender, args):
        metadata, groups, columns, unused_logo_path, unused_orientation, font_name, table_font_size = self._selected_report_state()
        if not columns:
            forms.alert("Select at least one report column before copying.", title=TITLE)
            return
        try:
            copy_report_to_clipboard(
                metadata, groups, columns, logo_path=None,
                font_name=font_name, table_font_size=table_font_size)
            self._set_status("Formatted report copied. Paste into Word with Ctrl+V.")
        except Exception as ex:
            self._operation_failed(ex)

    def close_clicked(self, sender, args):
        self.Close()

    def window_closing(self, sender, args):
        try:
            self._save_report_preferences()
        except Exception:
            pass
        try:
            self._gateway.dispose()
        except Exception:
            pass


def _find_existing_window():
    try:
        windows = Application.Current.Windows
    except Exception:
        return None
    for window in windows:
        try:
            if text(getattr(window, "Tag", "")) == WINDOW_MARKER:
                return window
        except Exception:
            continue
    return None


def _focus_existing_window(window):
    try:
        window.refresh_ced_theme_from_config()
    except Exception:
        pass
    try:
        if window.WindowState == WindowState.Minimized:
            window.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        window.Show()
        window.Activate()
        window.Focus()
    except Exception:
        pass


def _show_or_focus():
    existing = _find_existing_window()
    if existing is not None:
        _focus_existing_window(existing)
        return existing
    uidoc = _active_uidoc()
    doc = uidoc.Document if uidoc else None
    if doc is None:
        forms.alert("Open a Revit project before starting Revision Manager.", title=TITLE)
        return None
    snapshot = build_snapshot(doc)
    gateway = RevisionManagerGateway(__revit__, snapshot.get("document_key"), logger=LOGGER)
    window = RevisionManagerWindow(snapshot, gateway)
    window.Show()
    try:
        window.Activate()
    except Exception:
        pass
    return window


_WINDOW = _show_or_focus()
