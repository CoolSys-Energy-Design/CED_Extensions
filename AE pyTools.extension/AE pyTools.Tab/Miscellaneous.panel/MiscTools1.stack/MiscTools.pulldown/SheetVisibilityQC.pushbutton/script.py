# -*- coding: utf-8 -*-
"""Sheet Visibility QC.

Finds selected model-category elements which are not returned by a view-scoped
FilteredElementCollector for any selected, non-drafting view placed on the
selected sheets. Elements owned by Revit design options are excluded.

The collector is intentionally used as Revit's visibility test.  Revit's
collector documentation notes that its result is a graphics candidate set,
not a pixel-perfect test: crop-region edge cases and occlusion can still
produce false positives.  The UI calls this out so the report is not mistaken
for a rendered-image audit.

This file is self-contained so it can be copied into a pyRevit pushbutton.
It uses CEDLib's existing Snippets.categories electrical-category helper and
Revit API filters rather than LINQ.
"""

from __future__ import print_function

import codecs
import json
import os
import time

import clr

for _assembly in ("System", "PresentationCore", "PresentationFramework", "WindowsBase"):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System import Action
from System.Collections.Generic import List
from System.Windows import (
    Application,
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    ResizeMode,
    RoutedEventHandler,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Visibility,
    Window,
    WindowStartupLocation,
)
from System.Windows.Controls import (
    Border,
    Button,
    ColumnDefinition,
    Grid,
    GridView,
    GridViewColumn,
    GridViewColumnHeader,
    Label,
    ListBox,
    ListView,
    Orientation,
    RowDefinition,
    ScrollViewer,
    SelectionMode,
    StackPanel,
    TabControl,
    TabItem,
    TextBlock,
    TextBox,
    ProgressBar,
    WrapPanel,
    CheckBox,
    ComboBox,
    GroupStyle,
)
from System.Windows.Data import Binding, BindingMode, CollectionViewSource, PropertyGroupDescription
from System.Windows import DataTemplate, FrameworkElementFactory
from System.Windows.Media import Brushes
from System.Windows.Threading import DispatcherPriority
from pyrevit import DB, forms, revit, script


TITLE = "Sheet Visibility QC"
WINDOW_MARKER = "_ced_sheet_visibility_qc_window_v1"


CATEGORY_GROUPS = (
    ("Mechanical / HVAC", (
        ("Air Terminals", "OST_DuctTerminal"),
        ("Duct Accessories", "OST_DuctAccessory"),
        ("Duct Fittings", "OST_DuctFitting"),
        ("Ducts", "OST_DuctCurves"),
        ("Flex Ducts", "OST_FlexDuctCurves"),
        ("Mechanical Control Devices", "OST_MechanicalControlDevices"),
        ("Mechanical Equipment", "OST_MechanicalEquipment"),
    )),
    ("Plumbing / Piping", (
        ("Pipe Accessories", "OST_PipeAccessory"),
        ("Plumbing Equipment", "OST_PlumbingEquipment"),
        ("Pipe Fittings", "OST_PipeFitting"),
        ("Pipes", "OST_PipeCurves"),
        ("Flex Pipes", "OST_FlexPipeCurves"),
        ("Plumbing Fixtures", "OST_PlumbingFixtures"),
        ("Sprinklers", "OST_Sprinklers"),
    )),
    ("Electrical", (
        ("Cable Trays", "OST_CableTray"),
        ("Cable Tray Fittings", "OST_CableTrayFitting"),
        ("Conduits", "OST_Conduit"),
        ("Conduit Fittings", "OST_ConduitFitting"),
        ("Electrical Equipment", "OST_ElectricalEquipment"),
        ("Electrical Fixtures", "OST_ElectricalFixtures"),
        ("Lighting Devices", "OST_LightingDevices"),
        ("Lighting Fixtures", "OST_LightingFixtures"),
        ("Communication Devices", "OST_CommunicationDevices"),
        ("Data Devices", "OST_DataDevices"),
        ("Fire Alarm Devices", "OST_FireAlarmDevices"),
        ("Nurse Call Devices", "OST_NurseCallDevices"),
        ("Security Devices", "OST_SecurityDevices"),
        ("Telephone Devices", "OST_TelephoneDevices"),
    )),
    ("Multi-discipline / MEP-supporting", (
        ("Generic Models", "OST_GenericModel"),
        ("Specialty Equipment", "OST_SpecialityEquipment"),
    )),
)

SCAN_STATUS_TAGS = (
    "Needs review",
    "Accepted exception",
    "Resolved",
    "Investigate",
    "Covered",
)

THIS_DIR = os.path.abspath(os.path.dirname(__file__))

# The pushbutton is normally inside a CED extension, so this resolves the
# sibling CEDLib.lib directory without hard-coding the user's repository path.
try:
    from UIClasses import pathing as ui_pathing

    LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
except Exception:
    LIB_ROOT = None

if not LIB_ROOT:
    forms.alert(
        "Could not locate CEDLib.lib. Copy this pushbutton into a CED extension "
        "so the shared electrical category helper can be imported.",
        title=TITLE,
    )
    raise SystemExit

from Snippets import categories as category_utils  # noqa: E402
from Snippets import revit_helpers  # noqa: E402


LOGGER = script.get_logger()


def _text(value, fallback=""):
    if value is None:
        return fallback
    try:
        result = str(value)
    except Exception:
        result = fallback
    return result if result else fallback


def _html_escape(value):
    return (
        _text(value, "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def _id_value(value, default=-1):
    """Return a stable integer value for old and new Revit ElementId APIs."""
    if value is None:
        return int(default)
    # Result rows and operation payloads intentionally use plain integers.
    # The shared helper only understands ElementId-like objects and otherwise
    # returns its fallback, which made valid selected rows look like id 0.
    try:
        if not hasattr(value, "Value") and not hasattr(value, "IntegerValue"):
            return int(value)
    except Exception:
        pass
    try:
        return int(revit_helpers.get_elementid_value(value, default=default))
    except Exception:
        pass
    for attr in ("IntegerValue", "Value"):
        try:
            return int(getattr(value, attr))
        except Exception:
            pass
    try:
        return int(value)
    except Exception:
        return int(default)


def _element_id(value):
    try:
        return revit_helpers.elementid_from_value(int(value))
    except Exception:
        return DB.ElementId(int(value))


def _active_uidoc(application=None):
    if application is not None:
        try:
            if application.ActiveUIDocument is not None:
                return application.ActiveUIDocument
        except Exception:
            pass
    try:
        if __revit__.ActiveUIDocument is not None:
            return __revit__.ActiveUIDocument
    except Exception:
        pass
    try:
        return revit.uidoc
    except Exception:
        return None


def _active_doc(application=None):
    uidoc = _active_uidoc(application)
    if uidoc is not None:
        try:
            return uidoc.Document
        except Exception:
            pass
    try:
        return revit.doc
    except Exception:
        return None


def _document_title(doc):
    try:
        return _text(doc.Title, "Untitled")
    except Exception:
        return "Untitled"


def _document_path(doc):
    try:
        return _text(getattr(doc, "PathName", None), "")
    except Exception:
        return ""


def _document_key(doc):
    if doc is None:
        return ""
    values = []
    for attr in ("PathName", "Title"):
        try:
            values.append(_text(getattr(doc, attr), ""))
        except Exception:
            values.append("")
    try:
        values.append(str(doc.GetHashCode()))
    except Exception:
        pass
    return "|".join(values)


def _sorted_sheet_key(record):
    return (_text(record.get("number"), "").lower(), _text(record.get("name"), "").lower())


def _sheet_browser_group(doc, sheet):
    """Read the current Revit Sheets browser organization path when available."""
    try:
        organization = DB.BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc)
        if organization is not None:
            folder_items = list(organization.GetFolderItems(sheet.Id) or [])
            names = []
            for folder_item in folder_items:
                name = _text(getattr(folder_item, "Name", None), "")
                if name and name not in names:
                    names.append(name)
            if names:
                return " / ".join(names)
    except Exception:
        pass
    return "Ungrouped"


def _view_name(view):
    try:
        return _text(view.Name, "Unnamed view")
    except Exception:
        return "Unnamed view"


def _view_type_name(view):
    try:
        return _text(view.ViewType, "Unknown")
    except Exception:
        return "Unknown"


def _is_drafting_view(view):
    try:
        return view.ViewType == DB.ViewType.DraftingView
    except Exception:
        return _view_type_name(view).lower().replace(" ", "") == "draftingview"


def _category_records(doc):
    """Return the supported category choices and their default state."""
    electrical_ids = set()
    try:
        electrical_ids = set(
            _id_value(item)
            for item in list(category_utils.get_all_electrical_category_ids(doc=doc) or [])
            if _id_value(item) not in (-1, 0)
        )
    except Exception:
        pass

    records = []
    for group_name, definitions in CATEGORY_GROUPS:
        for display_name, enum_name in definitions:
            built_in = getattr(DB.BuiltInCategory, enum_name, None)
            if built_in is None:
                continue
            try:
                element_id = DB.ElementId(built_in)
                category = DB.Category.GetCategory(doc, element_id)
            except Exception:
                category = None
                element_id = None
            if category is None or element_id is None:
                continue
            value = _id_value(element_id, default=0)
            if value in (-1, 0):
                continue
            records.append(
                {
                    "key": enum_name,
                    "name": display_name,
                    "group": group_name,
                    "category_id": value,
                    "checked": value in electrical_ids and group_name == "Electrical",
                    "display": "{} - {}".format(group_name, display_name),
                }
            )
    return records


def _category_filter(doc, category_keys=None):
    """Build a selected-category/main-model Revit filter once per scan."""
    records = _category_records(doc)
    by_key = dict((item["key"], item) for item in records)
    if category_keys is None:
        selected = [item["key"] for item in records if bool(item.get("checked"))]
    else:
        selected = [
            _text(item, "")
            for item in list(category_keys or [])
            if _text(item, "") in by_key
        ]
    if not selected:
        raise Exception("Select at least one model category before scanning.")

    revit_ids = List[DB.ElementId]()
    seen = set()
    for key in selected:
        value = int(by_key[key]["category_id"])
        # BuiltInCategory ElementIds are negative in Revit.  Only the invalid
        # sentinel (normally -1) and zero should be rejected here.
        if value in (-1, 0) or value in seen:
            continue
        seen.add(value)
        revit_ids.Add(_element_id(value))
    if revit_ids.Count == 0:
        raise Exception("The selected categories are not available in this document.")
    category_filter = DB.ElementMulticategoryFilter(revit_ids)
    try:
        # Keep circuits out even if a future shared-category helper expands its
        # electrical set.  The UI intentionally does not expose this category.
        circuit_built_in = getattr(DB.BuiltInCategory, "OST_ElectricalCircuit", None)
        if circuit_built_in is not None:
            try:
                circuit_filter = DB.ElementCategoryFilter(
                    DB.ElementId(circuit_built_in),
                    True,
                )
                category_filter = DB.LogicalAndFilter(category_filter, circuit_filter)
            except Exception:
                # The selected multicategory filter already excludes circuits;
                # keep older Revit API overloads usable if this optional guard
                # is unavailable.
                pass
        # InvalidElementId is Revit's documented main-model design-option
        # filter.  Elements owned by any design option are excluded before
        # they can become candidates or view-collector results.
        main_model_filter = DB.ElementDesignOptionFilter(DB.ElementId.InvalidElementId)
        return DB.LogicalAndFilter(category_filter, main_model_filter), int(revit_ids.Count), selected
    except Exception as ex:
        raise Exception("Could not build the main-model design-option filter: {}".format(ex))


def _electrical_filter(doc):
    """Compatibility wrapper for callers that still request electrical defaults."""
    category_filter, category_count, _ = _category_filter(doc, None)
    return category_filter, category_count


def _collect_sheet_records(doc):
    records = []
    collector = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSheet)
        .WhereElementIsNotElementType()
    )
    for sheet in collector:
        sheet_id = _id_value(getattr(sheet, "Id", None))
        if sheet_id <= 0:
            continue
        number = _text(getattr(sheet, "SheetNumber", None), "")
        name = _text(getattr(sheet, "Name", None), "Unnamed sheet")
        non_drafting_count = 0
        try:
            placed_ids = list(sheet.GetAllPlacedViews() or [])
        except Exception:
            placed_ids = []
        for placed_id in placed_ids:
            try:
                view = doc.GetElement(placed_id)
            except Exception:
                view = None
            if view is None or _is_drafting_view(view):
                continue
            non_drafting_count += 1
        suffix = "{} non-drafting view{}".format(
            non_drafting_count,
            "" if non_drafting_count == 1 else "s",
        )
        records.append(
            {
                "element_id": sheet_id,
                "number": number,
                "name": name,
                "view_count": non_drafting_count,
                "browser_group": _sheet_browser_group(doc, sheet),
                "display": "{} - {} ({})".format(number, name, suffix),
            }
        )
    records.sort(key=_sorted_sheet_key)
    return records


def _collect_view_records(doc, sheet_ids, warnings=None):
    """Return unique non-drafting views placed on the selected sheets."""
    warning_list = warnings if warnings is not None else []
    by_view_id = {}
    order = []

    for sheet_value in list(sheet_ids or []):
        sheet_id = _element_id(sheet_value)
        try:
            sheet = doc.GetElement(sheet_id)
        except Exception:
            sheet = None
        if sheet is None:
            warning_list.append("A selected sheet could not be found in the active document.")
            continue
        sheet_number = _text(getattr(sheet, "SheetNumber", None), "")
        sheet_name = _text(getattr(sheet, "Name", None), "Unnamed sheet")
        sheet_label = "{} - {}".format(sheet_number, sheet_name)
        try:
            placed_ids = list(sheet.GetAllPlacedViews() or [])
        except Exception as ex:
            warning_list.append("Could not read views on {}: {}".format(sheet_label, ex))
            continue

        for placed_id in placed_ids:
            view_id = _id_value(placed_id)
            if view_id <= 0:
                continue
            try:
                view = doc.GetElement(placed_id)
            except Exception:
                view = None
            if view is None:
                continue
            if _is_drafting_view(view):
                continue
            try:
                if bool(view.IsTemplate):
                    continue
            except Exception:
                pass

            if view_id not in by_view_id:
                record = {
                    "element_id": view_id,
                    "view": view,
                    "name": _view_name(view),
                    "type": _view_type_name(view),
                    "sheet_labels": [],
                }
                by_view_id[view_id] = record
                order.append(record)
            record = by_view_id[view_id]
            if sheet_label not in record["sheet_labels"]:
                record["sheet_labels"].append(sheet_label)

    for record in order:
        record["display"] = "{} - {} [{}]".format(
            "; ".join(record["sheet_labels"]),
            record["name"],
            record["type"],
        )
    order.sort(key=lambda x: _text(x.get("display"), "").lower())
    return order


def _collect_ids_from_view(doc, view, category_filter, target_values=None, warnings=None):
    """Collect element ids that pass the selected-category and optional id filters."""
    warning_list = warnings if warnings is not None else []

    # Schedules and a few other placed view types do not expose drawn-element
    # iteration even though they are valid views on a sheet.  Use Revit's own
    # validity test before constructing the view-scoped collector.
    validity_check = getattr(DB.FilteredElementCollector, "IsViewValidForElementIteration", None)
    if validity_check is not None and not bool(validity_check(doc, view.Id)):
        warning_list.append(
            "Skipped {} because Revit does not support drawn-element "
            "iteration for this view type.".format(_view_name(view))
        )
        return set()

    collector = DB.FilteredElementCollector(doc, view.Id)
    filter_to_apply = category_filter
    used_id_filter = False

    target_values = set(target_values or [])
    if target_values:
        # The id filter keeps the secondary coverage scan small.  It is an API
        # filter, not a Python/LINQ post-filter.  The fallback keeps the tool
        # compatible with Revit versions that do not expose ElementIdSetFilter.
        id_filter_type = getattr(DB, "ElementIdSetFilter", None)
        if id_filter_type is not None:
            id_list = List[DB.ElementId]()
            for value in target_values:
                id_list.Add(_element_id(value))
            try:
                filter_to_apply = DB.LogicalAndFilter(
                    category_filter,
                    id_filter_type(id_list),
                )
                used_id_filter = True
            except Exception as ex:
                warning_list.append(
                    "The id filter was unavailable for {} ({}); using the "
                    "selected-category view collector instead.".format(_view_name(view), ex)
                )

    try:
        collector.WherePasses(filter_to_apply).WhereElementIsNotElementType()
        collected_ids = collector.ToElementIds()
    except Exception as ex:
        if used_id_filter:
            warning_list.append(
                "The optimized id filter failed for {} ({}); retrying the "
                "selected-category view collector.".format(_view_name(view), ex)
            )
            try:
                collector = DB.FilteredElementCollector(doc, view.Id)
                collector.WherePasses(category_filter).WhereElementIsNotElementType()
                collected_ids = collector.ToElementIds()
            except Exception:
                raise
        else:
            raise

    values = set()
    for element_id in collected_ids:
        value = _id_value(element_id)
        if value <= 0:
            continue
        if target_values and value not in target_values:
            # Only the compatibility fallback reaches this Python guard.
            continue
        values.add(value)
    return values


def _collect_candidate_ids(doc, category_filter):
    collector = DB.FilteredElementCollector(doc)
    collector.WherePasses(category_filter).WhereElementIsNotElementType()
    values = set()
    for element_id in collector.ToElementIds():
        value = _id_value(element_id)
        if value > 0:
            values.add(value)
    return values


def _family_type_parts(doc, element):
    symbol = None
    try:
        symbol = getattr(element, "Symbol", None)
    except Exception:
        symbol = None
    family_name = ""
    type_name = ""
    if symbol is not None:
        try:
            type_name = _text(symbol.Name, "")
        except Exception:
            pass
        try:
            family = getattr(symbol, "Family", None)
            family_name = _text(getattr(family, "Name", None), "")
        except Exception:
            pass
    if not type_name:
        try:
            type_name = _text(element.Name, "")
        except Exception:
            pass
    return family_name or "-", type_name or "-"


def _family_type_label(doc, element):
    family_name, type_name = _family_type_parts(doc, element)
    if family_name and type_name:
        return "{} : {}".format(family_name, type_name)
    return family_name or type_name or "-"


def _level_label(doc, element):
    try:
        level_id = getattr(element, "LevelId", None)
        level_value = _id_value(level_id)
        if level_value > 0:
            level = doc.GetElement(_element_id(level_value))
            if level is not None:
                return _text(getattr(level, "Name", None), "-")
    except Exception:
        pass
    return "-"


def _element_row(doc, element_id, visible_view_count=0, views_text=""):
    try:
        element = doc.GetElement(_element_id(element_id))
    except Exception:
        element = None
    category_name = "-"
    if element is not None:
        try:
            category_name = _text(element.Category.Name, "-")
        except Exception:
            pass
    family_name, type_name = _family_type_parts(doc, element) if element is not None else ("-", "-")
    family_type = _family_type_label(doc, element) if element is not None else "-"
    level = _level_label(doc, element) if element is not None else "-"
    return {
        "element_id": int(element_id),
        "id": str(int(element_id)),
        "category": category_name,
        "family": family_name,
        "type": type_name,
        "family_type": family_type,
        "level": level,
        "view_count": int(visible_view_count or 0),
        "views": views_text or "-",
        "passed_sheet": views_text or "-",
        "display": "{} | {} | {}".format(category_name, family_type, element_id),
    }


def _row_sort_key(row):
    return (
        _text(row.get("category"), "").lower(),
        _text(row.get("family_type"), "").lower(),
        int(row.get("element_id") or 0),
    )


def _new_scan_state(doc, sheet_ids, category_keys=None):
    """Prepare a resumable primary scan without iterating every view yet."""
    warnings = []
    category_filter, category_count, selected_category_keys = _category_filter(doc, category_keys)
    candidate_values = _collect_candidate_ids(doc, category_filter)
    view_records = _collect_view_records(doc, sheet_ids, warnings=warnings)
    return {
        "doc_key": _document_key(doc),
        "sheet_count": len(list(sheet_ids or [])),
        "category_filter": category_filter,
        "category_count": int(category_count),
        "category_keys": selected_category_keys,
        "candidate_values": candidate_values,
        "view_records": view_records,
        "visible_values": set(),
        "provenance": {},
        "warnings": warnings,
        "index": 0,
        "total": max(1, len(view_records) + 1),
    }


def _finish_scan_state(doc, state):
    candidate_values = state.get("candidate_values") or set()
    visible_values = state.get("visible_values") or set()
    missing_values = []
    covered_values = []
    for value in sorted(candidate_values):
        if value in visible_values:
            covered_values.append(value)
        else:
            missing_values.append(value)

    missing_rows = []
    covered_rows = []
    for value in missing_values:
        missing_rows.append(_element_row(doc, value, visible_view_count=0))
    for value in covered_values:
        covered_rows.append(
            _element_row(
                doc,
                value,
                visible_view_count=1,
                views_text=state.get("provenance", {}).get(value, "-"),
            )
        )
    missing_rows.sort(key=_row_sort_key)
    covered_rows.sort(key=_row_sort_key)

    return {
        "status": "ok",
        "doc_title": _document_title(doc),
        "doc_key": _document_key(doc),
        "sheet_count": int(state.get("sheet_count") or 0),
        "view_count": len(state.get("view_records") or []),
        "category_count": int(state.get("category_count") or 0),
        "candidate_count": len(candidate_values),
        "covered_count": len(covered_values),
        "missing_count": len(missing_values),
        "missing_rows": missing_rows,
        "covered_rows": covered_rows,
        "warnings": list(state.get("warnings") or []),
    }


def _advance_scan_state(doc, state):
    """Process one placed view, or finalize after the last view."""
    if _document_key(doc) != _text(state.get("doc_key"), ""):
        raise Exception("The active Revit document changed while the scan was running.")

    view_records = state.get("view_records") or []
    index = int(state.get("index") or 0)
    if index < len(view_records):
        record = view_records[index]
        view = record.get("view")
        if view is not None:
            try:
                visible_values = _collect_ids_from_view(
                    doc,
                    view,
                    state.get("category_filter"),
                    warnings=state.get("warnings"),
                )
                state["visible_values"].update(visible_values)
                for element_id in visible_values:
                    if element_id not in state["provenance"]:
                        state["provenance"][element_id] = record.get("display", "-")
            except Exception as ex:
                state["warnings"].append(
                    "Could not collect selected-category elements from {}: {}".format(
                        record.get("display", _view_name(view)),
                        ex,
                    )
                )
        state["index"] = index + 1
        return {
            "done": False,
            "current": int(state["index"]),
            "total": int(state["total"]),
            "view_display": record.get("display", "view"),
        }

    return {
        "done": True,
        "current": int(state.get("total") or 1),
        "total": int(state.get("total") or 1),
        "result": _finish_scan_state(doc, state),
    }


def scan_document(doc, sheet_ids, logger=None, category_keys=None):
    """Synchronous compatibility wrapper around the resumable scan."""
    state = _new_scan_state(doc, sheet_ids, category_keys=category_keys)
    while True:
        step = _advance_scan_state(doc, state)
        if bool(step.get("done")):
            return step.get("result")


def _new_target_coverage_state(
    doc,
    sheet_ids,
    selected_values,
    empty_label,
    category_keys=None,
):
    """Prepare a targeted coverage scan for only the selected element ids."""
    warnings = []
    category_filter, category_count, selected_category_keys = _category_filter(doc, category_keys)
    ordered_values = []
    target_values = set()
    for value in list(selected_values or []):
        numeric = _id_value(value, default=0)
        if numeric <= 0 or numeric in target_values:
            continue
        target_values.add(numeric)
        ordered_values.append(numeric)

    view_records = _collect_view_records(doc, sheet_ids, warnings=warnings)
    return {
        "doc_key": _document_key(doc),
        "category_filter": category_filter,
        "category_count": int(category_count),
        "category_keys": selected_category_keys,
        "sheet_count": len(list(sheet_ids or [])),
        "ordered_values": ordered_values,
        "target_values": target_values,
        "view_records": view_records,
        "view_names_by_element": dict((value, []) for value in ordered_values),
        "warnings": warnings,
        "empty_label": empty_label,
        "index": 0,
        "total": max(1, len(view_records) + 1),
    }


def _finish_target_coverage_state(doc, state):
    view_names_by_element = state.get("view_names_by_element") or {}
    rows = []
    for value in list(state.get("ordered_values") or []):
        names = view_names_by_element.get(value, [])
        rows.append(
            _element_row(
                doc,
                value,
                visible_view_count=len(names),
                views_text="; ".join(names) if names else state.get("empty_label", "No selected view passed this element"),
            )
        )
    rows.sort(key=_row_sort_key)
    return {
        "status": "ok",
        "doc_title": _document_title(doc),
        "doc_key": _document_key(doc),
        "category_count": int(state.get("category_count") or 0),
        "sheet_count": int(state.get("sheet_count") or 0),
        "view_count": len(state.get("view_records") or []),
        "rows": rows,
        "warnings": list(state.get("warnings") or []),
    }


def _advance_target_coverage_state(doc, state):
    if _document_key(doc) != _text(state.get("doc_key"), ""):
        raise Exception("The active Revit document changed while the targeted scan was running.")

    view_records = state.get("view_records") or []
    index = int(state.get("index") or 0)
    if index < len(view_records):
        record = view_records[index]
        view = record.get("view")
        if view is not None:
            try:
                visible_values = _collect_ids_from_view(
                    doc,
                    view,
                    state.get("category_filter"),
                    target_values=state.get("target_values"),
                    warnings=state.get("warnings"),
                )
                for value in visible_values:
                    if value in state["view_names_by_element"]:
                        state["view_names_by_element"][value].append(record.get("display", "-"))
            except Exception as ex:
                state["warnings"].append(
                    "Could not collect selected elements from {}: {}".format(
                        record.get("display", _view_name(view)),
                        ex,
                    )
                )
        state["index"] = index + 1
        return {
            "done": False,
            "current": int(state["index"]),
            "total": int(state["total"]),
            "view_display": record.get("display", "view"),
        }

    return {
        "done": True,
        "current": int(state.get("total") or 1),
        "total": int(state.get("total") or 1),
        "result": _finish_target_coverage_state(doc, state),
    }


def scan_element_view_coverage(doc, sheet_ids, selected_values, logger=None, category_keys=None):
    """Synchronous compatibility wrapper around targeted coverage scanning."""
    state = _new_target_coverage_state(
        doc,
        sheet_ids,
        selected_values,
        "No selected view passed this element",
        category_keys=category_keys,
    )
    while True:
        step = _advance_target_coverage_state(doc, state)
        if bool(step.get("done")):
            return step.get("result")


class SheetRow(object):
    def __init__(self, data):
        self.element_id = int(data.get("element_id") or 0)
        self.number = _text(data.get("number"), "")
        self.name = _text(data.get("name"), "")
        self.view_count = int(data.get("view_count") or 0)
        self.browser_group = _text(data.get("browser_group"), "Ungrouped")
        self.is_included = bool(data.get("is_included", True))
        self.group_sort = ""
        self.display = _text(data.get("display"), "")


class CategoryRow(object):
    def __init__(self, data):
        self.key = _text(data.get("key"), "")
        self.name = _text(data.get("name"), "")
        self.group = _text(data.get("group"), "")
        self.category_id = int(data.get("category_id") or 0)
        self.is_checked = bool(data.get("checked", False))
        self.display = _text(data.get("display"), self.name)


class ResultRow(object):
    def __init__(self, data):
        self.element_id = int(data.get("element_id") or 0)
        self.id = _text(data.get("id"), str(self.element_id))
        self.category = _text(data.get("category"), "-")
        self.family = _text(data.get("family"), "-")
        self.type = _text(data.get("type"), "-")
        self.family_type = _text(data.get("family_type"), "-")
        self.level = _text(data.get("level"), "-")
        self.view_count = int(data.get("view_count") or 0)
        self.views = _text(data.get("views"), "-")
        self.passed_sheet = _text(data.get("passed_sheet"), self.views)
        self.status = _text(
            data.get("status"),
            "Covered" if self.view_count > 0 else "Needs review",
        )
        self.display = _text(data.get("display"), "")


class SheetVisibilityGateway(object):
    def __init__(self, logger=None):
        self.logger = logger
        self.pending = None
        self.scan_state = None
        self.coverage_state = None
        self.cancelled = False
        self.cancel_requested = False
        self.handler = _SheetVisibilityExternalEventHandler(self)
        self.event = ExternalEvent.Create(self.handler)

    def busy(self):
        if self.pending is not None:
            return True
        try:
            return bool(self.event.IsPending)
        except Exception:
            return False

    def active_work(self):
        return bool(self.pending is not None or self.scan_state is not None or self.coverage_state is not None)

    def cancel(self, closing=False):
        """Cancel queued and resumable work; a current API call ends normally."""
        self.cancel_requested = True
        self.cancelled = bool(closing)
        self.pending = None
        self.scan_state = None
        self.coverage_state = None

    def raise_operation(self, name, payload=None, callback=None):
        if self.cancelled or self.busy():
            return False
        self.cancel_requested = False
        self.pending = {
            "name": _text(name),
            "payload": dict(payload or {}),
            "callback": callback,
        }
        try:
            self.event.Raise()
            return True
        except Exception as ex:
            self.pending = None
            if self.logger:
                self.logger.warning("Sheet Visibility QC event failed: {}".format(ex))
            return False

    def consume(self):
        pending = self.pending
        self.pending = None
        return pending


class _SheetVisibilityExternalEventHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self.gateway = gateway

    def Execute(self, application):  # noqa: N802
        pending = self.gateway.consume()
        if not pending:
            return
        name = pending.get("name")
        payload = pending.get("payload") or {}
        callback = pending.get("callback")
        status = "ok"
        result = None
        error = None
        try:
            uidoc = _active_uidoc(application)
            doc = uidoc.Document if uidoc is not None else None
            if doc is None:
                raise Exception("No active Revit document is available.")
            if name == "load_sheets":
                result = {
                    "doc_title": _document_title(doc),
                    "doc_key": _document_key(doc),
                    "sheets": _collect_sheet_records(doc),
                    "categories": _category_records(doc),
                }
            elif name == "scan_start":
                self.gateway.scan_state = _new_scan_state(
                    doc,
                    list(payload.get("sheet_ids") or []),
                    category_keys=payload.get("category_keys"),
                )
                state = self.gateway.scan_state
                result = {
                    "status": "ok",
                    "doc_title": _document_title(doc),
                    "doc_key": _document_key(doc),
                    "candidate_count": len(state.get("candidate_values") or []),
                    "sheet_count": int(state.get("sheet_count") or 0),
                    "view_count": len(state.get("view_records") or []),
                    "category_count": int(state.get("category_count") or 0),
                    "category_keys": list(state.get("category_keys") or []),
                    "progress_current": 0,
                    "progress_total": int(state.get("total") or 1),
                    "warnings": list(state.get("warnings") or []),
                }
            elif name == "scan_step":
                if self.gateway.scan_state is None:
                    raise Exception("No active sheet visibility scan exists.")
                step = _advance_scan_state(doc, self.gateway.scan_state)
                result = step
                if bool(step.get("done")):
                    self.gateway.scan_state = None
            elif name == "targeted_coverage_start":
                target_sheet_ids = list(payload.get("sheet_ids") or [])
                if bool(payload.get("all_sheets")):
                    target_sheet_ids = [
                        int(item.get("element_id") or 0)
                        for item in _collect_sheet_records(doc)
                        if int(item.get("element_id") or 0) > 0
                    ]
                self.gateway.coverage_state = _new_target_coverage_state(
                    doc,
                    target_sheet_ids,
                    list(payload.get("element_ids") or []),
                    payload.get("empty_label") or "No selected view passed this element",
                    category_keys=payload.get("category_keys"),
                )
                state = self.gateway.coverage_state
                result = {
                    "status": "ok",
                    "doc_title": _document_title(doc),
                    "doc_key": _document_key(doc),
                    "sheet_count": int(state.get("sheet_count") or 0),
                    "view_count": len(state.get("view_records") or []),
                    "target_count": len(state.get("ordered_values") or []),
                    "progress_current": 0,
                    "progress_total": int(state.get("total") or 1),
                    "warnings": list(state.get("warnings") or []),
                }
            elif name == "targeted_coverage_step":
                if self.gateway.coverage_state is None:
                    raise Exception("No active targeted coverage scan exists.")
                step = _advance_target_coverage_state(doc, self.gateway.coverage_state)
                result = step
                if bool(step.get("done")):
                    self.gateway.coverage_state = None
            elif name == "scan":
                result = scan_document(
                    doc,
                    list(payload.get("sheet_ids") or []),
                    logger=self.gateway.logger,
                    category_keys=payload.get("category_keys"),
                )
            elif name == "coverage":
                result = scan_element_view_coverage(
                    doc,
                    list(payload.get("sheet_ids") or []),
                    list(payload.get("element_ids") or []),
                    logger=self.gateway.logger,
                    category_keys=payload.get("category_keys"),
                )
            elif name in ("select", "show"):
                element_ids = List[DB.ElementId]()
                seen = set()
                for raw_value in list(payload.get("element_ids") or []):
                    value = _id_value(raw_value, default=0)
                    if value <= 0 or value in seen:
                        continue
                    element_id = _element_id(value)
                    try:
                        element = doc.GetElement(element_id)
                    except Exception:
                        element = None
                    if element is None:
                        continue
                    seen.add(value)
                    element_ids.Add(element_id)
                if uidoc is None:
                    raise Exception("No active Revit UI document is available.")
                if name == "show" and element_ids.Count > 0:
                    try:
                        uidoc.ShowElements(element_ids)
                    except Exception as ex:
                        if self.gateway.logger:
                            self.gateway.logger.debug("ShowElements failed: {}".format(ex))
                uidoc.Selection.SetElementIds(element_ids)
                result = {"selected": int(element_ids.Count)}
            else:
                raise Exception("Unknown operation: {}".format(name))
        except Exception as ex:
            status = "error"
            error = ex
            if name in ("scan_start", "scan_step", "scan"):
                self.gateway.scan_state = None
            if name in (
                "targeted_coverage_start",
                "targeted_coverage_step",
                "coverage",
            ):
                self.gateway.coverage_state = None
            if self.gateway.logger:
                try:
                    self.gateway.logger.exception("Sheet Visibility QC failed: {}".format(ex))
                except Exception:
                    pass
        if self.gateway.cancel_requested and name in (
            "load_sheets", "scan_start", "scan_step", "scan",
            "targeted_coverage_start", "targeted_coverage_step", "coverage",
        ):
            status = "cancelled"
            result = None
            error = None
            self.gateway.cancel_requested = False
        if callback:
            try:
                callback(status, name, result, error)
            except Exception:
                pass

    def GetName(self):  # noqa: N802
        return "CED Sheet Visibility QC"


def _make_button(text, width=120, tooltip=None):
    button = Button()
    button.Content = text
    button.Width = width
    button.Height = 28
    button.Margin = Thickness(3, 2, 3, 2)
    button.Padding = Thickness(8, 2, 8, 2)
    if tooltip:
        button.ToolTip = tooltip
    return button


def _make_header(text):
    block = TextBlock()
    block.Text = text
    block.FontWeight = FontWeights.Bold
    block.Margin = Thickness(0, 0, 0, 4)
    return block


def _make_checkbox_template(property_name, click_handler, preview_handler=None):
    template = DataTemplate()
    factory = FrameworkElementFactory(CheckBox)
    binding = Binding(property_name)
    binding.Mode = BindingMode.TwoWay
    factory.SetBinding(
        CheckBox.IsCheckedProperty,
        binding,
    )
    try:
        if preview_handler is not None:
            factory.AddHandler(
                CheckBox.PreviewMouseLeftButtonDownEvent,
                RoutedEventHandler(preview_handler),
            )
        factory.AddHandler(
            CheckBox.ClickEvent,
            RoutedEventHandler(click_handler),
        )
    except Exception:
        pass
    template.VisualTree = factory
    return template


def _make_check_list_view(check_property, click_handler, columns, preview_handler=None):
    view = ListView()
    view.SelectionMode = SelectionMode.Extended
    view.HorizontalContentAlignment = HorizontalAlignment.Stretch
    view.Margin = Thickness(0)
    grid_view = GridView()
    check_column = GridViewColumn()
    check_column.Header = "Use"
    check_column.Width = 42
    check_column.CellTemplate = _make_checkbox_template(
        check_property,
        click_handler,
        preview_handler=preview_handler,
    )
    grid_view.Columns.Add(check_column)
    for header, path, width in columns:
        column = GridViewColumn()
        column.Header = header
        column.Width = width
        column.DisplayMemberBinding = Binding(path)
        grid_view.Columns.Add(column)
    view.View = grid_view
    try:
        view.GroupStyle.Add(GroupStyle())
    except Exception:
        pass
    return view


def _make_list_view(columns, header_handler=None):
    view = ListView()
    view.SelectionMode = SelectionMode.Extended
    view.HorizontalContentAlignment = HorizontalAlignment.Stretch
    view.Margin = Thickness(0)
    grid_view = GridView()
    for header, path, width in columns:
        column = GridViewColumn()
        column.Header = header
        column.Width = width
        column.DisplayMemberBinding = Binding(path)
        grid_view.Columns.Add(column)
    view.View = grid_view
    if header_handler is not None:
        try:
            view.AddHandler(
                GridViewColumnHeader.ClickEvent,
                RoutedEventHandler(header_handler),
            )
        except Exception:
            pass
    return view


class SheetVisibilityQCWindow(Window):
    def __init__(self, gateway):
        Window.__init__(self)
        self.gateway = gateway
        self.doc_key = ""
        self.included_sheets = []
        self.ignored_sheets = []
        self.missing_rows = []
        self.covered_rows = []
        self.coverage_rows = []
        self.imported_missing_ids = []
        self.sheet_rows = []
        self.category_rows = []
        self.selected_category_keys = []
        self.status_by_id = {}
        self._checkbox_click_targets = {}
        self._closing = False
        self._close_confirmed = False
        self.config = script.get_config("SheetVisibilityQC")
        self._sort_state = {}
        self._targeted_mode = ""
        self._ui_busy = False
        self._build_ui()
        self.Closing += self.window_closing
        self._set_busy(False)
        self._set_status("Loading sheets...")
        self._raise_load_sheets()

    def _build_ui(self):
        self.Title = TITLE
        self.Width = 1250
        self.Height = 720
        self.MinWidth = 980
        self.MinHeight = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize
        self.Background = Brushes.White
        self.Foreground = Brushes.Black
        try:
            self.Tag = WINDOW_MARKER
            setattr(self, WINDOW_MARKER, True)
        except Exception:
            pass

        root = Grid()
        root.Margin = Thickness(12)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(330)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        header = StackPanel()
        title = TextBlock()
        title.Text = TITLE
        title.FontSize = 20
        title.FontWeight = FontWeights.Bold
        header.Children.Add(title)
        self.document_text = TextBlock()
        self.document_text.Text = "Document: -"
        self.document_text.Margin = Thickness(0, 3, 0, 0)
        header.Children.Add(self.document_text)
        note = TextBlock()
        note.Text = (
            "Collector-based graphics candidate audit. Results may include elements "
            "outside crop boundaries or obscured by other elements."
        )
        note.TextWrapping = TextWrapping.Wrap
        note.Margin = Thickness(0, 4, 0, 8)
        header.Children.Add(note)
        root.Children.Add(header)

        selection_border = Border()
        selection_border.BorderBrush = Brushes.LightGray
        selection_border.BorderThickness = Thickness(1)
        selection_border.Padding = Thickness(8)
        selection_grid = Grid()
        selection_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        selection_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(12)))
        selection_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))

        sheets_panel = StackPanel()
        sheets_panel.Children.Add(_make_header("Sheets to scan"))
        sheet_controls = WrapPanel()
        self.sheet_search = TextBox()
        self.sheet_search.Width = 260
        self.sheet_search.Margin = Thickness(0, 0, 6, 4)
        self.sheet_search.ToolTip = "Search sheets by number, name, or browser group"
        self.sheet_search.TextChanged += self.sheet_search_changed
        sheet_controls.Children.Add(self.sheet_search)
        self.sheet_group_combo = ComboBox()
        self.sheet_group_combo.Width = 175
        self.sheet_group_combo.Margin = Thickness(0, 0, 6, 4)
        self.sheet_group_combo.Items.Add("Browser organization")
        self.sheet_group_combo.Items.Add("Sheet number prefix")
        self.sheet_group_combo.Items.Add("Flat list")
        self.sheet_group_combo.SelectedIndex = 0
        self.sheet_group_combo.SelectionChanged += self.sheet_group_changed
        self.sheet_group_combo.ToolTip = "Group sheets using the current Revit browser organization or a simple list"
        sheet_controls.Children.Add(self.sheet_group_combo)
        self.show_checked_only = CheckBox()
        self.show_checked_only.Content = "Show checked only"
        self.show_checked_only.VerticalAlignment = VerticalAlignment.Center
        self.show_checked_only.Margin = Thickness(0, 0, 0, 4)
        self.show_checked_only.Checked += self.sheet_filter_changed
        self.show_checked_only.Unchecked += self.sheet_filter_changed
        self.show_checked_only.ToolTip = "Hide sheets that are not included in the scan"
        sheet_controls.Children.Add(self.show_checked_only)
        sheets_panel.Children.Add(sheet_controls)

        self.sheet_list = _make_check_list_view(
            "is_included",
            self.sheet_checkbox_clicked,
            (("Sheet", "display", 360), ("Browser group", "browser_group", 180), ("Views", "view_count", 55)),
            preview_handler=self.checkbox_preview_mouse_down,
        )
        self.sheet_list.Height = 238
        self.sheet_list.SelectionChanged += self.sheet_selection_changed
        sheets_panel.Children.Add(self.sheet_list)
        sheet_buttons = WrapPanel()
        self.check_all_sheets_button = _make_button(
            "Check all filtered",
            135,
            "Check every sheet currently visible after search and the checked-only filter.",
        )
        self.check_all_sheets_button.Click += self.include_all_clicked
        sheet_buttons.Children.Add(self.check_all_sheets_button)
        self.uncheck_all_sheets_button = _make_button(
            "Uncheck all filtered",
            145,
            "Uncheck every sheet currently visible after search and the checked-only filter.",
        )
        self.uncheck_all_sheets_button.Click += self.ignore_all_clicked
        sheet_buttons.Children.Add(self.uncheck_all_sheets_button)
        sheets_panel.Children.Add(sheet_buttons)
        Grid.SetColumn(sheets_panel, 0)
        selection_grid.Children.Add(sheets_panel)

        categories_panel = StackPanel()
        categories_panel.Children.Add(_make_header("Categories to scan"))
        self.category_search = TextBox()
        self.category_search.ToolTip = "Search supported categories"
        self.category_search.Margin = Thickness(0, 0, 0, 4)
        self.category_search.TextChanged += self.category_search_changed
        categories_panel.Children.Add(self.category_search)
        self.category_list = _make_check_list_view(
            "is_checked",
            self.category_checkbox_clicked,
            (("Category", "name", 220), ("Group", "group", 185)),
            preview_handler=self.checkbox_preview_mouse_down,
        )
        self.category_list.Height = 238
        self.category_list.SelectionChanged += self.category_selection_changed
        categories_panel.Children.Add(self.category_list)
        category_buttons = WrapPanel()
        self.check_electrical_button = _make_button(
            "Check electrical",
            120,
            "Restore the default electrical category selection. Electrical circuits are intentionally excluded.",
        )
        self.check_electrical_button.Click += self.check_electrical_clicked
        category_buttons.Children.Add(self.check_electrical_button)
        self.check_all_categories_button = _make_button(
            "Check all",
            90,
            "Include every supported mechanical, plumbing, electrical, and MEP-supporting category.",
        )
        self.check_all_categories_button.Click += self.check_all_categories_clicked
        category_buttons.Children.Add(self.check_all_categories_button)
        self.uncheck_all_categories_button = _make_button(
            "Uncheck all",
            105,
            "Clear all category checkboxes. A scan requires at least one category.",
        )
        self.uncheck_all_categories_button.Click += self.uncheck_all_categories_clicked
        category_buttons.Children.Add(self.uncheck_all_categories_button)
        categories_panel.Children.Add(category_buttons)
        Grid.SetColumn(categories_panel, 2)
        selection_grid.Children.Add(categories_panel)

        selection_border.Child = selection_grid
        Grid.SetRow(selection_border, 1)
        root.Children.Add(selection_border)

        actions_border = Border()
        actions_border.BorderBrush = Brushes.LightGray
        actions_border.BorderThickness = Thickness(1)
        actions_border.Padding = Thickness(5, 3, 5, 3)
        actions_border.Margin = Thickness(0, 8, 0, 0)
        actions_panel = WrapPanel()
        actions_panel.Orientation = Orientation.Horizontal
        actions_panel.HorizontalAlignment = HorizontalAlignment.Left
        self.export_button = _make_button(
            "Export JSON",
            105,
            "Export a machine-readable snapshot of the sheet set, category choices, results, and tags. Use Import JSON later to restore results and recheck failures without a full scan.",
        )
        self.export_button.Click += self.export_clicked
        actions_panel.Children.Add(self.export_button)
        self.import_button = _make_button(
            "Import JSON",
            105,
            "Load a previous Sheet Visibility QC JSON snapshot. This does not scan; use Recheck failed to test only its previously failed elements.",
        )
        self.import_button.Click += self.import_clicked
        actions_panel.Children.Add(self.import_button)
        self.report_button = _make_button(
            "Export report",
            120,
            "Export a human-readable HTML report with the document, sheet set, scan rules, tags, and result tables.",
        )
        self.report_button.Click += self.report_clicked
        actions_panel.Children.Add(self.report_button)
        self.recheck_button = _make_button(
            "Recheck failed",
            125,
            "Re-scan only the elements imported from the previous report that were not visible; it does not repeat the full candidate scan.",
        )
        self.recheck_button.Click += self.recheck_imported_clicked
        actions_panel.Children.Add(self.recheck_button)
        self.find_sheets_button = _make_button(
            "Find all sheets",
            125,
            "Find every non-drafting sheet view that passes the selected elements, across all sheets in the model.",
        )
        self.find_sheets_button.Click += self.find_all_sheets_clicked
        actions_panel.Children.Add(self.find_sheets_button)
        self.scan_button = _make_button(
            "Scan selected sheets",
            155,
            "Run the full visibility QC for included sheets and checked categories.",
        )
        self.scan_button.Click += self.scan_clicked
        actions_panel.Children.Add(self.scan_button)
        self.select_button = _make_button(
            "Select selected",
            125,
            "Select the highlighted result elements in the active Revit model.",
        )
        self.select_button.Click += self.select_clicked
        actions_panel.Children.Add(self.select_button)
        self.show_button = _make_button(
            "Show in model",
            120,
            "Try to zoom to and select the highlighted result elements in the active Revit model.",
        )
        self.show_button.Click += self.show_clicked
        actions_panel.Children.Add(self.show_button)
        self.coverage_button = _make_button(
            "Find views for covered",
            170,
            "For selected rows on the Covered tab, find all additional sheet views that pass those elements.",
        )
        self.coverage_button.Click += self.coverage_clicked
        actions_panel.Children.Add(self.coverage_button)
        self.cancel_button = _make_button(
            "Cancel scan",
            105,
            "Cancel queued scan work and prevent the next view collector from starting. The current Revit API call may finish first.",
        )
        self.cancel_button.Click += self.cancel_clicked
        actions_panel.Children.Add(self.cancel_button)
        self.close_button = _make_button(
            "Close",
            80,
            "Close the utility. If a scan is running, you will be asked to confirm cancellation.",
        )
        self.close_button.Click += self.close_clicked
        actions_panel.Children.Add(self.close_button)
        actions_border.Child = actions_panel
        Grid.SetRow(actions_border, 2)
        root.Children.Add(actions_border)

        results_border = Border()
        results_border.BorderBrush = Brushes.LightGray
        results_border.BorderThickness = Thickness(1)
        results_border.Padding = Thickness(8)
        results_grid = Grid()
        results_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        results_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        results_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        result_header = _make_header("Results")
        Grid.SetRow(result_header, 0)
        results_grid.Children.Add(result_header)

        result_filters = WrapPanel()
        result_filters.Margin = Thickness(0, 0, 0, 5)
        filter_label = TextBlock()
        filter_label.Text = "Filter:"
        filter_label.VerticalAlignment = VerticalAlignment.Center
        filter_label.Margin = Thickness(0, 0, 5, 0)
        result_filters.Children.Add(filter_label)
        self.result_filter_field = ComboBox()
        self.result_filter_field.Width = 120
        self.result_filter_field.Margin = Thickness(0, 0, 5, 0)
        for item in ("All fields", "Category", "Family", "Family / Type", "Type", "Level", "Status"):
            self.result_filter_field.Items.Add(item)
        self.result_filter_field.SelectedIndex = 0
        self.result_filter_field.SelectionChanged += self.result_filter_changed
        self.result_filter_field.ToolTip = "Choose which result column to search."
        result_filters.Children.Add(self.result_filter_field)
        self.result_filter_text = TextBox()
        self.result_filter_text.Width = 235
        self.result_filter_text.Margin = Thickness(0, 0, 5, 0)
        self.result_filter_text.ToolTip = "Filter the active results tab by category, family/type, level, status, or all fields."
        self.result_filter_text.TextChanged += self.result_filter_changed
        result_filters.Children.Add(self.result_filter_text)
        self.clear_result_filter_button = _make_button(
            "Clear filter",
            90,
            "Clear the results filter and show every row in the active tab.",
        )
        self.clear_result_filter_button.Click += self.clear_result_filter_clicked
        result_filters.Children.Add(self.clear_result_filter_button)
        tag_label = TextBlock()
        tag_label.Text = "Tag selected:"
        tag_label.VerticalAlignment = VerticalAlignment.Center
        tag_label.Margin = Thickness(12, 0, 5, 0)
        result_filters.Children.Add(tag_label)
        self.tag_combo = ComboBox()
        self.tag_combo.Width = 135
        self.tag_combo.Margin = Thickness(0, 0, 5, 0)
        for item in SCAN_STATUS_TAGS:
            self.tag_combo.Items.Add(item)
        self.tag_combo.SelectedIndex = 0
        self.tag_combo.SelectionChanged += self.result_tag_changed
        self.tag_combo.ToolTip = "Apply a review status to the selected result rows; tags are included in JSON and HTML exports."
        result_filters.Children.Add(self.tag_combo)
        self.tag_button = _make_button(
            "Apply tag",
            85,
            "Apply the selected status tag to highlighted result rows."
        )
        self.tag_button.Click += self.tag_selected_clicked
        result_filters.Children.Add(self.tag_button)
        Grid.SetRow(result_filters, 1)
        results_grid.Children.Add(result_filters)

        self.tabs = TabControl()
        self.missing_tab = TabItem()
        self.missing_tab.Header = "Not visible on selected sheets"
        self.covered_tab = TabItem()
        self.covered_tab.Header = "Covered / passed"
        self.coverage_tab = TabItem()
        self.coverage_tab.Header = "View coverage details"

        self.missing_list = _make_list_view(
             (("Element Id", "id", 82), ("Category", "category", 145),
              ("Family : Type", "family_type", 300), ("Level", "level", 145),
              ("Status", "status", 125)),
            self._result_header_handler("missing"),
        )
        self.covered_list = _make_list_view(
            (("Element Id", "id", 82), ("Category", "category", 145),
              ("Family : Type", "family_type", 300), ("Level", "level", 145),
              ("Status", "status", 125), ("Passed on sheet/view", "passed_sheet", 300)),
            self._result_header_handler("covered"),
        )
        self.coverage_list = _make_list_view(
             (("Element Id", "id", 82), ("Category", "category", 145),
              ("Family : Type", "family_type", 285), ("Level", "level", 130),
              ("Status", "status", 125), ("View count", "view_count", 85),
              ("Views on selected sheets", "views", 500)),
            self._result_header_handler("coverage"),
        )
        self.missing_tab.Content = self.missing_list
        self.covered_tab.Content = self.covered_list
        self.coverage_tab.Content = self.coverage_list
        self.tabs.Items.Add(self.missing_tab)
        self.tabs.Items.Add(self.covered_tab)
        self.tabs.Items.Add(self.coverage_tab)
        self.tabs.SelectionChanged += self.result_tab_changed
        Grid.SetRow(self.tabs, 2)
        results_grid.Children.Add(self.tabs)
        results_border.Child = results_grid
        Grid.SetRow(results_border, 3)
        root.Children.Add(results_border)

        footer = Grid()
        footer.Margin = Thickness(0, 8, 0, 0)
        footer.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        footer.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(220)))
        self.status_text = TextBlock()
        self.status_text.VerticalAlignment = VerticalAlignment.Center
        self.status_text.TextWrapping = TextWrapping.Wrap
        Grid.SetColumn(self.status_text, 0)
        footer.Children.Add(self.status_text)
        self.progress_bar = ProgressBar()
        self.progress_bar.Minimum = 0
        self.progress_bar.Maximum = 1
        self.progress_bar.Value = 0
        self.progress_bar.Height = 16
        self.progress_bar.Margin = Thickness(8, 0, 8, 0)
        self.progress_bar.VerticalAlignment = VerticalAlignment.Center
        self.progress_bar.Visibility = Visibility.Collapsed
        Grid.SetColumn(self.progress_bar, 1)
        footer.Children.Add(self.progress_bar)
        Grid.SetRow(footer, 4)
        root.Children.Add(footer)

        self.Content = root
        self.missing_list.SelectionChanged += self.result_selection_changed
        self.covered_list.SelectionChanged += self.result_selection_changed
        self.coverage_list.SelectionChanged += self.result_selection_changed

    def _set_status(self, text):
        self.status_text.Text = _text(text, "Ready")

    def _set_busy(self, busy):
        self._ui_busy = bool(busy)
        enabled = not bool(busy)
        for name in (
            "export_button",
            "import_button",
            "report_button",
            "scan_button",
            "check_all_sheets_button",
            "uncheck_all_sheets_button",
            "check_electrical_button",
            "check_all_categories_button",
            "uncheck_all_categories_button",
        ):
            control = getattr(self, name, None)
            if control is not None:
                control.IsEnabled = enabled
        if getattr(self, "cancel_button", None) is not None:
            self.cancel_button.IsEnabled = bool(busy)
        self._sync_result_buttons()

    def _set_progress(self, current=0, total=1, visible=False):
        total_value = max(1, int(total or 1))
        current_value = max(0, min(int(current or 0), total_value))
        self.progress_bar.Maximum = total_value
        self.progress_bar.Value = current_value
        self.progress_bar.Visibility = Visibility.Visible if visible else Visibility.Collapsed

    def cancel_clicked(self, sender, args):
        if not self._ui_busy and not self.gateway.active_work():
            return
        self.gateway.cancel(closing=False)
        self._targeted_mode = ""
        self._set_progress(0, 1, visible=False)
        self._set_busy(False)
        self._set_status("Scan cancelled. The current Revit collector call, if any, will be allowed to finish.")

    def window_closing(self, sender, args):
        if self._close_confirmed:
            return
        if self._ui_busy or self.gateway.active_work():
            try:
                args.Cancel = True
            except Exception:
                pass
            forms.alert(
                "A Sheet Visibility QC operation is still running. Closing this window will cancel the queued scan work and stop the next view from being processed.",
                title=TITLE,
            )
            self.gateway.cancel(closing=True)
            self._targeted_mode = ""
            self._closing = True
            self._close_confirmed = True
            self._set_progress(0, 1, visible=False)
            self._set_busy(False)
            self.Close()
            return
        self.gateway.cancel(closing=True)
        self._closing = True
        self._close_confirmed = True

    def _refresh_sheet_lists(self):
        query = _text(getattr(self.sheet_search, "Text", None), "").strip().lower()
        checked_only = bool(getattr(self.show_checked_only, "IsChecked", False))
        rows = []
        for item in list(self.sheet_rows or []):
            searchable = "{} {} {}".format(item.number, item.name, item.browser_group).lower()
            if query and query not in searchable:
                continue
            if checked_only and not item.is_included:
                continue
            mode = _text(getattr(self.sheet_group_combo, "SelectedItem", None), "Flat list")
            if mode == "Browser organization":
                item.group_sort = item.browser_group
            elif mode == "Sheet number prefix":
                item.group_sort = item.number.split("-")[0].strip() or "Ungrouped"
            else:
                item.group_sort = "All sheets"
            rows.append(item)
        rows.sort(key=lambda item: (item.group_sort.lower(), item.number.lower(), item.name.lower()))
        self.sheet_list.ItemsSource = None
        if _text(getattr(self.sheet_group_combo, "SelectedItem", None), "") == "Flat list":
            self.sheet_list.ItemsSource = rows
        else:
            try:
                collection_view = CollectionViewSource.GetDefaultCollectionView(rows)
                collection_view.GroupDescriptions.Clear()
                collection_view.GroupDescriptions.Add(PropertyGroupDescription("group_sort"))
                self.sheet_list.ItemsSource = collection_view
            except Exception:
                self.sheet_list.ItemsSource = rows
        self._sync_sheet_state()
        self._sync_result_buttons()

    def sheet_search_changed(self, sender, args):
        self._refresh_sheet_lists()

    def sheet_filter_changed(self, sender, args):
        self._refresh_sheet_lists()

    def sheet_group_changed(self, sender, args):
        self._refresh_sheet_lists()

    def _sync_sheet_state(self):
        self.included_sheets = [item for item in self.sheet_rows if item.is_included]
        self.ignored_sheets = [item for item in self.sheet_rows if not item.is_included]

    def _set_sheet_rows_state(self, rows, checked):
        for item in list(rows or []):
            item.is_included = bool(checked)
        self._refresh_sheet_lists()

    def _visible_sheet_rows(self):
        query = _text(getattr(self.sheet_search, "Text", None), "").strip().lower()
        checked_only = bool(getattr(self.show_checked_only, "IsChecked", False))
        rows = []
        for item in self.sheet_rows:
            searchable = "{} {} {}".format(item.number, item.name, item.browser_group).lower()
            if query and query not in searchable:
                continue
            if checked_only and not item.is_included:
                continue
            rows.append(item)
        return rows

    def checkbox_preview_mouse_down(self, sender, args):
        row = getattr(sender, "DataContext", None)
        if row is None:
            return
        control = self.sheet_list if row in self.sheet_rows else self.category_list
        try:
            selected = list(control.SelectedItems)
        except Exception:
            selected = []
        self._checkbox_click_targets[id(sender)] = (
            selected if row in selected else [row]
        )

    def sheet_checkbox_clicked(self, sender, args):
        row = getattr(sender, "DataContext", None)
        if row is None:
            return
        desired = bool(getattr(sender, "IsChecked", False))
        targets = self._checkbox_click_targets.pop(id(sender), None)
        if targets is None:
            selected = list(self.sheet_list.SelectedItems)
            targets = selected if row in selected else [row]
        self._set_sheet_rows_state(targets, desired)

    def _refresh_category_list(self):
        query = _text(getattr(self.category_search, "Text", None), "").strip().lower()
        rows = [
            item for item in self.category_rows
            if not query or query in "{} {}".format(item.name, item.group).lower()
        ]
        self.category_list.ItemsSource = None
        self.category_list.ItemsSource = rows
        self._sync_category_state()

    def _sync_category_state(self):
        self.selected_category_keys = [item.key for item in self.category_rows if item.is_checked]

    def category_search_changed(self, sender, args):
        self._refresh_category_list()

    def category_selection_changed(self, sender, args):
        self._sync_result_buttons()

    def category_checkbox_clicked(self, sender, args):
        row = getattr(sender, "DataContext", None)
        if row is None:
            return
        desired = bool(getattr(sender, "IsChecked", False))
        targets = self._checkbox_click_targets.pop(id(sender), None)
        if targets is None:
            selected = list(self.category_list.SelectedItems)
            targets = selected if row in selected else [row]
        for item in targets:
            item.is_checked = desired
        self._refresh_category_list()

    def check_electrical_clicked(self, sender, args):
        for item in self.category_rows:
            item.is_checked = item.group == "Electrical"
        self._refresh_category_list()

    def check_all_categories_clicked(self, sender, args):
        for item in self.category_rows:
            item.is_checked = True
        self._refresh_category_list()

    def uncheck_all_categories_clicked(self, sender, args):
        for item in self.category_rows:
            item.is_checked = False
        self._refresh_category_list()

    def _apply_sheets(self, result):
        self.doc_key = _text(result.get("doc_key"), "")
        self.document_text.Text = "Document: {}".format(_text(result.get("doc_title"), "-"))
        self.sheet_rows = [SheetRow(item) for item in list(result.get("sheets") or [])]
        self._sync_sheet_state()
        self.category_rows = [CategoryRow(item) for item in list(result.get("categories") or [])]
        saved_categories = getattr(self.config, "category_keys", None)
        valid_categories = set(item.key for item in self.category_rows)
        if saved_categories is not None:
            saved_categories = [
                _text(item, "") for item in list(saved_categories)
                if _text(item, "") in valid_categories
            ]
            for item in self.category_rows:
                item.is_checked = item.key in saved_categories
        self._sync_category_state()
        self._refresh_sheet_lists()
        self._refresh_category_list()
        self._set_status(
            "Loaded {} sheets and {} supported categories. Check sheets and categories before scanning.".format(
                len(self.sheet_rows), len(self.category_rows)
            )
        )

    def _sheet_ids(self):
        self._sync_sheet_state()
        return [int(item.element_id) for item in self.included_sheets if item.element_id > 0]

    def _category_keys(self):
        self._sync_category_state()
        return list(self.selected_category_keys or [])

    def _save_category_selection(self):
        self.selected_category_keys = self._category_keys()
        try:
            self.config.category_keys = list(self.selected_category_keys)
            script.save_config()
        except Exception as ex:
            LOGGER.warning("Could not save Sheet Visibility QC category configuration: {}".format(ex))

    def _row_to_dict(self, row):
        return {
            "element_id": int(getattr(row, "element_id", 0) or 0),
            "id": _text(getattr(row, "id", None), ""),
            "category": _text(getattr(row, "category", None), "-"),
            "family": _text(getattr(row, "family", None), "-"),
            "type": _text(getattr(row, "type", None), "-"),
            "family_type": _text(getattr(row, "family_type", None), "-"),
            "level": _text(getattr(row, "level", None), "-"),
            "view_count": int(getattr(row, "view_count", 0) or 0),
            "views": _text(getattr(row, "views", None), "-"),
            "passed_sheet": _text(getattr(row, "passed_sheet", None), "-"),
            "status": _text(getattr(row, "status", None), "Needs review"),
            "display": _text(getattr(row, "display", None), ""),
        }

    def _sheet_to_dict(self, sheet):
        return {
            "element_id": int(getattr(sheet, "element_id", 0) or 0),
            "number": _text(getattr(sheet, "number", None), ""),
            "name": _text(getattr(sheet, "name", None), ""),
            "view_count": int(getattr(sheet, "view_count", 0) or 0),
            "browser_group": _text(getattr(sheet, "browser_group", None), "Ungrouped"),
            "is_included": bool(getattr(sheet, "is_included", True)),
            "display": _text(getattr(sheet, "display", None), ""),
        }

    def _result_export_payload(self):
        doc = _active_doc()
        return {
            "format": "CED.SheetVisibilityQC",
            "version": 3,
            "exported_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "document": {
                "title": _document_title(doc),
                "path": _document_path(doc),
                "doc_key": self.doc_key,
            },
            "sheet_set": {
                "included": [self._sheet_to_dict(item) for item in self.included_sheets],
                "ignored": [self._sheet_to_dict(item) for item in self.ignored_sheets],
            },
            "scan_options": {
                "supported_categories": True,
                "exclude_design_options": True,
                "exclude_drafting_views": True,
                "category_keys": list(self.selected_category_keys or []),
            },
            "results": {
                "missing": [self._row_to_dict(row) for row in self.missing_rows],
                "covered": [self._row_to_dict(row) for row in self.covered_rows],
            },
        }

    def _set_result_rows(self, missing_rows, covered_rows):
        self._sort_state = {}
        self.missing_rows = [self._result_row(row) for row in list(missing_rows or [])]
        self.covered_rows = [self._result_row(row) for row in list(covered_rows or [])]
        self._refresh_result_lists()
        self.missing_tab.Header = "Not visible ({})".format(len(self.missing_rows))
        self.covered_tab.Header = "Covered / passed ({})".format(len(self.covered_rows))

    def _result_row(self, data):
        if isinstance(data, ResultRow):
            row = data
        else:
            row = ResultRow(data)
            if isinstance(data, dict) and "status" not in data:
                prior = self.status_by_id.get(int(row.element_id))
                if prior:
                    row.status = prior
        if int(row.element_id) > 0:
            self.status_by_id[int(row.element_id)] = _text(row.status, "Needs review")
        return row

    def _result_matches(self, row):
        query = _text(getattr(self.result_filter_text, "Text", None), "").strip().lower()
        if not query:
            return True
        field = _text(getattr(self.result_filter_field, "SelectedItem", None), "All fields")
        values = {
            "Category": (getattr(row, "category", ""),),
            "Family / Type": (getattr(row, "family_type", ""),),
            "Family": (getattr(row, "family", ""),),
            "Type": (getattr(row, "type", ""),),
            "Level": (getattr(row, "level", ""),),
            "Status": (getattr(row, "status", ""),),
        }.get(field)
        if values is None:
            values = (
                getattr(row, "id", ""),
                getattr(row, "category", ""),
                getattr(row, "family", ""),
                getattr(row, "type", ""),
                getattr(row, "family_type", ""),
                getattr(row, "level", ""),
                getattr(row, "status", ""),
                getattr(row, "passed_sheet", ""),
                getattr(row, "views", ""),
            )
        return any(query in _text(value, "").lower() for value in values)

    def _refresh_result_lists(self):
        for rows, control in (
            (self.missing_rows, self.missing_list),
            (self.covered_rows, self.covered_list),
            (self.coverage_rows, self.coverage_list),
        ):
            control.Items.Clear()
            for row in list(rows or []):
                if self._result_matches(row):
                    control.Items.Add(row)

    def result_filter_changed(self, sender, args):
        self._refresh_result_lists()
        self._sync_result_buttons()

    def clear_result_filter_clicked(self, sender, args):
        self.result_filter_text.Text = ""

    def result_tag_changed(self, sender, args):
        self._sync_result_buttons()

    def tag_selected_clicked(self, sender, args):
        tag = _text(getattr(self.tag_combo, "SelectedItem", None), "")
        if not tag:
            return
        rows = self._selected_rows()
        if not rows:
            return
        for row in rows:
            row.status = tag
            self.status_by_id[int(row.element_id)] = tag
        self._refresh_result_lists()
        self._set_status("Tagged {} result{} as {}.".format(
            len(rows), "" if len(rows) == 1 else "s", tag
        ))
        self._sync_result_buttons()

    def _restore_imported_sheet_set(self, payload):
        imported_document = payload.get("document") or {}
        current_doc = _active_doc()
        imported_path = _text(imported_document.get("path"), "")
        current_path = _document_path(current_doc)
        imported_title = _text(imported_document.get("title"), "")
        current_title = _document_title(current_doc)
        same_document = bool(
            imported_path and current_path and imported_path.lower() == current_path.lower()
        )
        if not same_document and not imported_path and imported_title and imported_title == current_title:
            same_document = True
        if not same_document:
            return False

        sheet_set = payload.get("sheet_set") or {}
        included_data = list(sheet_set.get("included") or [])
        ignored_data = list(sheet_set.get("ignored") or [])
        current_lookup = {}
        for item in list(self.sheet_rows or []):
            current_lookup[int(item.element_id)] = item
        referenced = set()
        imported_included = []
        imported_ignored = []
        for data, destination in ((included_data, imported_included), (ignored_data, imported_ignored)):
            for item_data in data:
                value = int(item_data.get("element_id") or 0)
                if value <= 0 or value not in current_lookup or value in referenced:
                    continue
                referenced.add(value)
                destination.append(current_lookup[value])
        if not referenced:
            return False
        for value, item in current_lookup.items():
            if value not in referenced:
                imported_included.append(item)
        imported_included.sort(key=lambda item: (item.number.lower(), item.name.lower()))
        imported_ignored.sort(key=lambda item: (item.number.lower(), item.name.lower()))
        for item in imported_included:
            item.is_included = True
        for item in imported_ignored:
            item.is_included = False
        self._sync_sheet_state()
        self._refresh_sheet_lists()
        return True

    def export_clicked(self, sender, args):
        if not self.missing_rows and not self.covered_rows:
            self._set_status("There are no results to export yet.")
            return
        path = forms.save_file(
            file_ext="json",
            title="Export Sheet Visibility QC Results",
            default_name="Sheet_Visibility_QC_Results",
        )
        if not path:
            return
        if not _text(path).lower().endswith(".json"):
            path = "{}.json".format(path)
        try:
            with codecs.open(path, "w", "utf-8") as stream:
                json.dump(self._result_export_payload(), stream, indent=2, sort_keys=True)
            self._set_status("Exported QC results to {}".format(path))
        except Exception as ex:
            LOGGER.exception("Sheet Visibility QC export failed: {}".format(ex))
            forms.alert("Could not export QC results:\n\n{}".format(ex), title=TITLE)

    def _report_html(self):
        payload = self._result_export_payload()
        document = payload.get("document") or {}
        sheet_set = payload.get("sheet_set") or {}
        results = payload.get("results") or {}
        missing = list(results.get("missing") or [])
        covered = list(results.get("covered") or [])
        included = list(sheet_set.get("included") or [])
        ignored = list(sheet_set.get("ignored") or [])
        scan_options = payload.get("scan_options") or {}
        selected_category_keys = set(scan_options.get("category_keys") or [])
        category_labels = [
            "{} / {}".format(item.group, item.name)
            for item in self.category_rows
            if item.key in selected_category_keys
        ]
        category_text = ", ".join(category_labels) or "No category selection recorded"

        def table(headers, rows, empty_text):
            parts = ["<table><thead><tr>"]
            for header in headers:
                parts.append("<th>{}</th>".format(_html_escape(header)))
            parts.append("</tr></thead><tbody>")
            if rows:
                for row in rows:
                    parts.append("<tr>")
                    for value in row:
                        parts.append("<td>{}</td>".format(_html_escape(value)))
                    parts.append("</tr>")
            else:
                parts.append("<tr><td class=\"empty\" colspan=\"{}\">{}</td></tr>".format(
                    len(headers), _html_escape(empty_text)
                ))
            parts.append("</tbody></table>")
            return "".join(parts)

        missing_rows = [
            (
                item.get("id", ""),
                item.get("category", "-"),
                item.get("family_type", "-"),
                item.get("level", "-"),
                item.get("status", "Needs review"),
            )
            for item in missing
        ]
        covered_rows = [
            (
                item.get("id", ""),
                item.get("category", "-"),
                item.get("family_type", "-"),
                item.get("level", "-"),
                item.get("status", "Covered"),
                item.get("passed_sheet", item.get("views", "-")),
            )
            for item in covered
        ]
        sheet_rows = []
        for item in included:
            sheet_rows.append(("Included", item.get("display", "")))
        for item in ignored:
            sheet_rows.append(("Ignored", item.get("display", "")))

        html = [
            "<!DOCTYPE html>",
            "<html><head><meta charset=\"utf-8\"><title>Sheet Visibility QC Report</title>",
            "<style>",
            "body{font-family:Segoe UI,Arial,sans-serif;color:#202124;margin:32px;}",
            "h1{color:#1f4e79;margin-bottom:4px;}h2{color:#1f4e79;border-bottom:1px solid #c9d5e3;padding-bottom:4px;margin-top:28px;}",
            ".meta{color:#555;margin-bottom:18px;} .summary{display:flex;gap:12px;flex-wrap:wrap;margin:18px 0;}",
            ".card{border:1px solid #c9d5e3;background:#f4f8fb;padding:12px 18px;min-width:130px;}",
            ".number{font-size:22px;font-weight:bold;color:#1f4e79;} table{border-collapse:collapse;width:100%;margin-top:8px;}",
            "th{background:#1f4e79;color:white;text-align:left;padding:7px;border:1px solid #b7c4d1;}",
            "td{padding:6px;border:1px solid #d5dce3;vertical-align:top;}tr:nth-child(even){background:#f7f9fb;}",
            ".empty{text-align:center;color:#666;font-style:italic;} .note{background:#fff8df;border-left:4px solid #d6a700;padding:10px;margin:14px 0;}",
            "</style></head><body>",
            "<h1>Sheet Visibility QC Report</h1>",
            "<div class=\"meta\"><strong>Document:</strong> {}<br><strong>Path:</strong> {}<br><strong>Generated:</strong> {}<br><strong>Categories:</strong> {}</div>".format(
                _html_escape(document.get("title", "Untitled")),
                _html_escape(document.get("path", "Not saved")),
                _html_escape(payload.get("exported_at", "")),
                _html_escape(category_text),
            ),
            "<div class=\"summary\">",
            "<div class=\"card\"><div class=\"number\">{}</div>Not visible</div>".format(len(missing)),
            "<div class=\"card\"><div class=\"number\">{}</div>Covered / passed</div>".format(len(covered)),
            "<div class=\"card\"><div class=\"number\">{}</div>Included sheets</div>".format(len(included)),
            "<div class=\"card\"><div class=\"number\">{}</div>Ignored sheets</div>".format(len(ignored)),
            "</div>",
            "<div class=\"note\"><strong>Scan rules:</strong> Checked supported categories; electrical circuits excluded; design-option elements excluded; drafting views excluded. Revit view collectors return graphics candidates, so crop boundaries and occlusion may still require visual review.</div>",
            "<h2>Not visible on the selected sheet set ({})</h2>".format(len(missing)),
            table(
                ("Element Id", "Category", "Family : Type", "Level", "Status"),
                missing_rows,
                "No elements failed the visibility check.",
            ),
            "<h2>Covered / passed ({})</h2>".format(len(covered)),
            table(
                ("Element Id", "Category", "Family : Type", "Level", "Status", "First passing sheet/view"),
                covered_rows,
                "No covered elements were reported.",
            ),
            "<h2>Sheet set</h2>",
            table(("Use", "Sheet"), sheet_rows, "No sheets were included in the report."),
            "</body></html>",
        ]
        return "".join(html)

    def report_clicked(self, sender, args):
        if not self.missing_rows and not self.covered_rows:
            self._set_status("There are no results to report yet.")
            return
        path = forms.save_file(
            file_ext="html",
            title="Export Sheet Visibility QC Report",
            default_name="Sheet_Visibility_QC_Report",
        )
        if not path:
            return
        if not _text(path).lower().endswith(".html"):
            path = "{}.html".format(path)
        try:
            with codecs.open(path, "w", "utf-8") as stream:
                stream.write(self._report_html())
            self._set_status("Exported human-readable report to {}".format(path))
        except Exception as ex:
            LOGGER.exception("Sheet Visibility QC report export failed: {}".format(ex))
            forms.alert("Could not export the QC report:\n\n{}".format(ex), title=TITLE)

    def import_clicked(self, sender, args):
        path = forms.pick_file(file_ext="json", title="Import Sheet Visibility QC Results")
        if not path:
            return
        if isinstance(path, (list, tuple)):
            path = path[0] if path else None
        if not path:
            return
        try:
            with codecs.open(path, "r", "utf-8") as stream:
                payload = json.load(stream)
            if payload.get("format") != "CED.SheetVisibilityQC":
                raise Exception("The selected file is not a Sheet Visibility QC export.")
            results = payload.get("results") or {}
            missing = list(results.get("missing") or [])
            covered = list(results.get("covered") or [])
            if not missing and not covered:
                raise Exception("The export contains no element results.")
            restored = self._restore_imported_sheet_set(payload)
            imported_category_keys = (payload.get("scan_options") or {}).get("category_keys")
            if imported_category_keys is not None and self.category_rows:
                imported_category_keys = set(
                    _text(item, "") for item in list(imported_category_keys or [])
                )
                for item in self.category_rows:
                    item.is_checked = item.key in imported_category_keys
                self._save_category_selection()
                self._refresh_category_list()
            self._set_result_rows(missing, covered)
            self.coverage_rows = []
            self.coverage_list.Items.Clear()
            self.coverage_tab.Header = "View coverage details"
            self.imported_missing_ids = [
                int(row.element_id)
                for row in self.missing_rows
                if int(row.element_id) > 0
            ]
            self.tabs.SelectedIndex = 0
            sheet_message = "sheet set restored" if restored else "current sheet set retained"
            self._set_status(
                "Imported {} failed and {} passed result{}; {}. Click Recheck failed to test only the imported failures.".format(
                    len(self.missing_rows),
                    len(self.covered_rows),
                    "s" if len(self.covered_rows) != 1 else "",
                    sheet_message,
                )
            )
            self._sync_result_buttons()
        except Exception as ex:
            LOGGER.exception("Sheet Visibility QC import failed: {}".format(ex))
            forms.alert("Could not import QC results:\n\n{}".format(ex), title=TITLE)

    def _raise_load_sheets(self):
        self._set_busy(True)
        if not self.gateway.raise_operation("load_sheets", callback=self.operation_complete):
            self._set_busy(False)
            self._set_status("Could not request sheet list.")

    def ignore_selected_clicked(self, sender, args):
        selected = list(self.sheet_list.SelectedItems)
        self._set_sheet_rows_state(selected, False)

    def include_selected_clicked(self, sender, args):
        selected = list(self.sheet_list.SelectedItems)
        self._set_sheet_rows_state(selected, True)

    def include_all_clicked(self, sender, args):
        self._set_sheet_rows_state(self._visible_sheet_rows(), True)

    def ignore_all_clicked(self, sender, args):
        self._set_sheet_rows_state(self._visible_sheet_rows(), False)

    def scan_clicked(self, sender, args):
        sheet_ids = self._sheet_ids()
        if not sheet_ids:
            forms.alert("Include at least one sheet before scanning.", title=TITLE)
            return
        category_keys = self._category_keys()
        if not category_keys:
            forms.alert("Check at least one model category before scanning.", title=TITLE)
            return
        self._save_category_selection()
        self._set_busy(True)
        self._set_progress(0, 1, visible=True)
        self._set_status("Collecting selected model categories from sheet views...")
        if not self.gateway.raise_operation(
            "scan_start",
            payload={"sheet_ids": sheet_ids, "category_keys": category_keys},
            callback=self.operation_complete,
        ):
            self._set_busy(False)
            self._set_progress(0, 1, visible=False)
            self._set_status("A scan is already running.")

    def _schedule_scan_step(self):
        if self._closing or self.gateway.cancelled or self.gateway.cancel_requested:
            return
        try:
            self.Dispatcher.BeginInvoke(
                Action(self._raise_scan_step),
                DispatcherPriority.Background,
            )
        except Exception:
            self._raise_scan_step()

    def _raise_scan_step(self):
        if self._closing or self.gateway.cancelled or self.gateway.cancel_requested:
            return
        if not self.gateway.raise_operation("scan_step", callback=self.operation_complete):
            self._set_busy(False)
            self._set_progress(0, 1, visible=False)
            self._set_status("The scan could not continue because another Revit operation is running.")

    def _schedule_targeted_coverage_step(self):
        if self._closing or self.gateway.cancelled or self.gateway.cancel_requested:
            return
        try:
            self.Dispatcher.BeginInvoke(
                Action(self._raise_targeted_coverage_step),
                DispatcherPriority.Background,
            )
        except Exception:
            self._raise_targeted_coverage_step()

    def _raise_targeted_coverage_step(self):
        if self._closing or self.gateway.cancelled or self.gateway.cancel_requested:
            return
        if not self.gateway.raise_operation(
            "targeted_coverage_step",
            callback=self.operation_complete,
        ):
            self._targeted_mode = ""
            self._set_busy(False)
            self._set_progress(0, 1, visible=False)
            self._set_status(
                "The targeted scan could not continue because another Revit operation is running."
            )

    def _active_result_list(self):
        selected_index = int(self.tabs.SelectedIndex)
        if selected_index == 0:
            return self.missing_list
        if selected_index == 1:
            return self.covered_list
        return self.coverage_list

    def _result_header_handler(self, context):
        def handler(sender, args):
            self._sort_result_list(context, args)
        return handler

    def _column_binding_path(self, args):
        try:
            column = args.Column
            binding = column.DisplayMemberBinding
            return _text(binding.Path.Path, "")
        except Exception:
            try:
                header_paths = {
                    "Element Id": "id",
                    "Category": "category",
                    "Family : Type": "family_type",
                    "Level": "level",
                    "Status": "status",
                    "Passed on sheet/view": "passed_sheet",
                    "View count": "view_count",
                    "Views on selected sheets": "views",
                }
                return header_paths.get(_text(args.Column.Header, ""), "")
            except Exception:
                return ""

    def _sort_result_list(self, context, args):
        try:
            args.Handled = True
        except Exception:
            pass
        sort_key = self._column_binding_path(args)
        source_map = {
            "missing": ("missing_rows", self.missing_list),
            "covered": ("covered_rows", self.covered_list),
            "coverage": ("coverage_rows", self.coverage_list),
        }
        if not sort_key or context not in source_map:
            return

        previous = self._sort_state.get(context)
        descending = bool(previous and previous[0] == sort_key and not previous[1])
        self._sort_state[context] = (sort_key, descending)
        source_name, control = source_map[context]
        rows = list(getattr(self, source_name) or [])

        def row_key(row):
            raw_value = getattr(row, sort_key, None)
            if sort_key in ("id", "view_count"):
                try:
                    return int(raw_value or 0)
                except Exception:
                    return 0
            return _text(raw_value, "").lower()

        rows.sort(key=row_key, reverse=descending)
        setattr(self, source_name, rows)
        self._refresh_result_lists()
        direction = "descending" if descending else "ascending"
        self._set_status("Sorted {} results by {} ({})".format(context, sort_key, direction))

    def _selected_rows(self):
        return list(self._active_result_list().SelectedItems)

    def _selected_ids(self):
        values = []
        seen = set()
        for row in self._selected_rows():
            value = _id_value(getattr(row, "element_id", None), default=0)
            if value > 0 and value not in seen:
                seen.add(value)
                values.append(value)
        return values

    def _sync_result_buttons(self):
        try:
            selected_ids = self._selected_ids()
        except Exception:
            selected_ids = []
        operation_free = (not self._ui_busy) and (not self.gateway.busy())
        enabled = operation_free and bool(selected_ids)
        has_results = bool(self.missing_rows or self.covered_rows)
        for name, value in (
            ("export_button", operation_free and has_results),
            ("report_button", operation_free and has_results),
            ("import_button", operation_free),
            ("recheck_button", operation_free and bool(self.imported_missing_ids)),
            ("find_sheets_button", enabled),
            ("select_button", enabled),
            ("show_button", enabled),
            ("tag_button", enabled),
            ("clear_result_filter_button", operation_free),
        ):
            control = getattr(self, name, None)
            if control is not None:
                control.IsEnabled = value
        tabs = getattr(self, "tabs", None)
        if tabs is not None:
            self.coverage_button.IsEnabled = (
                enabled and int(tabs.SelectedIndex) == 1
            )

    def result_selection_changed(self, sender, args):
        self._sync_result_buttons()

    def sheet_selection_changed(self, sender, args):
        self._sync_result_buttons()

    def result_tab_changed(self, sender, args):
        self._sync_result_buttons()

    def _send_element_action(self, operation):
        element_ids = self._selected_ids()
        if not element_ids:
            return
        self._set_busy(True)
        self._set_status("{} {} element{}...".format(
            "Showing" if operation == "show" else "Selecting",
            len(element_ids),
            "" if len(element_ids) == 1 else "s",
        ))
        if not self.gateway.raise_operation(
            operation,
            payload={"element_ids": element_ids},
            callback=self.operation_complete,
        ):
            self._set_busy(False)
            self._set_status("The Revit operation is already running.")

    def select_clicked(self, sender, args):
        self._send_element_action("select")

    def show_clicked(self, sender, args):
        self._send_element_action("show")

    def coverage_clicked(self, sender, args):
        if int(self.tabs.SelectedIndex) != 1:
            return
        element_ids = self._selected_ids()
        if not element_ids:
            return
        self._start_targeted_coverage(
            "chosen_sheets",
            self._sheet_ids(),
            element_ids,
            all_sheets=False,
            empty_label="No selected view passed this element",
        )

    def _start_targeted_coverage(
        self,
        mode,
        sheet_ids,
        element_ids,
        all_sheets=False,
        empty_label="No sheet view passed this element",
    ):
        if not element_ids:
            return
        category_keys = self._category_keys()
        if not category_keys:
            forms.alert("Check at least one model category before running coverage.", title=TITLE)
            return
        self._save_category_selection()
        self._targeted_mode = _text(mode, "coverage")
        self._set_busy(True)
        self._set_progress(0, 1, visible=True)
        if self._targeted_mode == "recheck":
            status = "Rechecking {} previously failed element{}...".format(
                len(element_ids), "" if len(element_ids) == 1 else "s"
            )
        elif all_sheets:
            status = "Finding all sheet placements for {} selected element{}...".format(
                len(element_ids), "" if len(element_ids) == 1 else "s"
            )
        else:
            status = "Finding selected element views across the chosen sheet set..."
        self._set_status(status)
        if not self.gateway.raise_operation(
            "targeted_coverage_start",
            payload={
                "sheet_ids": list(sheet_ids or []),
                "element_ids": list(element_ids or []),
                "all_sheets": bool(all_sheets),
                "empty_label": empty_label,
                "category_keys": category_keys,
            },
            callback=self.operation_complete,
        ):
            self._targeted_mode = ""
            self._set_busy(False)
            self._set_progress(0, 1, visible=False)
            self._set_status("The Revit operation is already running.")

    def find_all_sheets_clicked(self, sender, args):
        element_ids = self._selected_ids()
        if not element_ids:
            return
        self._start_targeted_coverage(
            "all_sheets",
            [],
            element_ids,
            all_sheets=True,
            empty_label="No sheet view passed this element",
        )

    def recheck_imported_clicked(self, sender, args):
        element_ids = list(self.imported_missing_ids or [])
        if not element_ids:
            return
        sheet_ids = self._sheet_ids()
        if not sheet_ids:
            forms.alert("Include at least one sheet before rechecking.", title=TITLE)
            return
        self._start_targeted_coverage(
            "recheck",
            sheet_ids,
            element_ids,
            all_sheets=False,
            empty_label="Still not visible on the selected sheet set",
        )

    def _apply_primary_scan(self, result):
        self.doc_key = _text(result.get("doc_key"), self.doc_key)
        self.document_text.Text = "Document: {}".format(_text(result.get("doc_title"), "-"))
        self._set_result_rows(
            list(result.get("missing_rows") or []),
            list(result.get("covered_rows") or []),
        )
        self.imported_missing_ids = []
        self.coverage_rows = []
        self.coverage_list.Items.Clear()
        self.coverage_tab.Header = "View coverage details"
        warnings = list(result.get("warnings") or [])
        summary = (
            "Sheets: {} | Views: {} | Categories: {} | Candidates: {} | "
            "Covered: {} | Not visible: {}"
        ).format(
            result.get("sheet_count", 0),
            result.get("view_count", 0),
            result.get("category_count", 0),
            result.get("candidate_count", 0),
            result.get("covered_count", 0),
            result.get("missing_count", 0),
        )
        if warnings:
            summary += " | Warnings: {}".format(len(warnings))
        self._set_status(summary)
        if warnings:
            LOGGER.warning("Sheet Visibility QC warnings:\n{}".format("\n".join(warnings)))

    def _apply_coverage_scan(self, result):
        self._sort_state = {}
        self.coverage_rows = [self._result_row(item) for item in list(result.get("rows") or [])]
        self._refresh_result_lists()
        self.coverage_tab.Header = "View coverage details ({})".format(len(self.coverage_rows))
        self.tabs.SelectedIndex = 2
        warnings = list(result.get("warnings") or [])
        status = "View coverage found for {} selected element{} across {} view{}.".format(
            len(self.coverage_rows),
            "" if len(self.coverage_rows) == 1 else "s",
            result.get("view_count", 0),
            "" if int(result.get("view_count", 0)) == 1 else "s",
        )
        if warnings:
            status += " Warnings: {}.".format(len(warnings))
            LOGGER.warning("Sheet Visibility QC coverage warnings:\n{}".format("\n".join(warnings)))
        self._set_status(status)

    def _apply_recheck_result(self, result):
        rows = [self._result_row(item) for item in list(result.get("rows") or [])]
        unresolved = [row for row in rows if int(row.view_count) <= 0]
        corrected = [row for row in rows if int(row.view_count) > 0]

        covered_by_id = {}
        for row in self.covered_rows:
            covered_by_id[int(row.element_id)] = row
        for row in corrected:
            covered_by_id[int(row.element_id)] = row

        covered_rows = list(covered_by_id.values())
        covered_rows.sort(
            key=lambda row: (
                _text(row.category, "").lower(),
                _text(row.family_type, "").lower(),
                int(row.element_id),
            )
        )
        self._set_result_rows(unresolved, covered_rows)
        self.imported_missing_ids = [int(row.element_id) for row in unresolved]
        self.coverage_rows = rows
        self._refresh_result_lists()
        self.coverage_tab.Header = "Recheck coverage details ({})".format(len(rows))
        self.tabs.SelectedIndex = 0

        warnings = list(result.get("warnings") or [])
        status = "Recheck complete: {} corrected, {} still not visible.".format(
            len(corrected), len(unresolved)
        )
        if warnings:
            status += " Warnings: {}.".format(len(warnings))
            LOGGER.warning("Sheet Visibility QC recheck warnings:\n{}".format("\n".join(warnings)))
        self._set_status(status)

    def operation_complete(self, status, operation, result, error):
        if self._closing or self.gateway.cancelled:
            return
        if status == "cancelled":
            self._targeted_mode = ""
            self._set_progress(0, 1, visible=False)
            self._set_busy(False)
            self._set_status("Scan cancelled.")
            return
        if status != "ok":
            self._targeted_mode = ""
            self._set_progress(0, 1, visible=False)
            self._set_busy(False)
            forms.alert("{} failed:\n\n{}".format(operation, error), title=TITLE)
            self._set_status("Operation failed.")
            return
        if operation == "load_sheets":
            self._apply_sheets(result or {})
            self._set_busy(False)
            return
        if operation == "scan_start":
            data = result or {}
            self._set_progress(
                data.get("progress_current", 0),
                data.get("progress_total", 1),
                visible=True,
            )
            self._set_status(
                "Prepared {} category candidates; scanning {} placed view{}...".format(
                    data.get("candidate_count", 0),
                    data.get("view_count", 0),
                    "" if int(data.get("view_count", 0)) == 1 else "s",
                )
            )
            self._schedule_scan_step()
            return
        if operation == "scan_step":
            data = result or {}
            self._set_progress(
                data.get("current", 0),
                data.get("total", 1),
                visible=True,
            )
            if bool(data.get("done")):
                self._apply_primary_scan(data.get("result") or {})
                self._set_progress(data.get("total", 1), data.get("total", 1), visible=False)
                self._set_busy(False)
            else:
                self._set_status(
                    "Scanning view {}/{}: {}".format(
                        data.get("current", 0),
                        data.get("total", 0) - 1,
                        data.get("view_display", "view"),
                    )
                )
                self._schedule_scan_step()
            return
        if operation == "targeted_coverage_start":
            data = result or {}
            self._set_progress(
                data.get("progress_current", 0),
                data.get("progress_total", 1),
                visible=True,
            )
            self._set_status(
                "Prepared {} selected element{}; scanning {} placed view{}...".format(
                    data.get("target_count", 0),
                    "" if int(data.get("target_count", 0)) == 1 else "s",
                    data.get("view_count", 0),
                    "" if int(data.get("view_count", 0)) == 1 else "s",
                )
            )
            self._schedule_targeted_coverage_step()
            return
        if operation == "targeted_coverage_step":
            data = result or {}
            self._set_progress(
                data.get("current", 0),
                data.get("total", 1),
                visible=True,
            )
            if bool(data.get("done")):
                if self._targeted_mode == "recheck":
                    self._apply_recheck_result(data.get("result") or {})
                else:
                    self._apply_coverage_scan(data.get("result") or {})
                self._targeted_mode = ""
                self._set_progress(data.get("total", 1), data.get("total", 1), visible=False)
                self._set_busy(False)
            else:
                self._set_status(
                    "Targeted scan view {}/{}: {}".format(
                        data.get("current", 0),
                        data.get("total", 0) - 1,
                        data.get("view_display", "view"),
                    )
                )
                self._schedule_targeted_coverage_step()
            return
        if operation == "scan":
            self._apply_primary_scan(result or {})
            self._set_progress(1, 1, visible=False)
            self._set_busy(False)
            return
        if operation == "coverage":
            self._apply_coverage_scan(result or {})
            self._set_progress(0, 1, visible=False)
            self._set_busy(False)
            return
        if operation in ("select", "show"):
            selected = int((result or {}).get("selected") or 0)
            self._set_status("{} {} element{}.".format(
                "Selected" if operation == "select" else "Showed and selected",
                selected,
                "" if selected == 1 else "s",
            ))
            self._set_progress(0, 1, visible=False)
            self._set_busy(False)

    def close_clicked(self, sender, args):
        self.Close()


def _existing_window():
    app = Application.Current
    if app is None:
        return None
    try:
        windows = list(app.Windows)
    except Exception:
        windows = []
    for window in windows:
        try:
            if getattr(window, "Tag", "") == WINDOW_MARKER:
                return window
            if bool(getattr(window, WINDOW_MARKER, False)):
                return window
        except Exception:
            pass
    return None


def _launch():
    existing = _existing_window()
    if existing is not None:
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        return
    doc = _active_doc()
    if doc is None:
        forms.alert("No active Revit document is available.", title=TITLE)
        return
    gateway = SheetVisibilityGateway(logger=LOGGER)
    window = SheetVisibilityQCWindow(gateway)
    globals()["_CED_SHEET_VISIBILITY_QC_WINDOW"] = window
    globals()["_CED_SHEET_VISIBILITY_QC_GATEWAY"] = gateway
    window.Show()
    try:
        window.Activate()
    except Exception:
        pass


_launch()
