# -*- coding: utf-8 -*-
__title__ = "Create Circuits by Device Parameter"
__doc__ = ("Modeless circuit grouping and creation by a selected device "
           "parameter, with active-document tracking and model navigation.")

import os
import sys

import clr

for _asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Collections.Generic import List
from System.Windows import Application, WindowState
from pyrevit import revit, forms, script, DB

from Snippets import revit_helpers

THIS_DIR = os.path.dirname(__file__)
if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

import cg_core
import cg_window
import cg_collect
import cg_apply

logger = script.get_logger()
TITLE = "Create Circuits by Device Parameter"
_WINDOW_MARKER = "_ae_circuit_grouper_window_persistent_v6"
_LEGACY_WINDOW_MARKERS = (
    "_ae_circuit_grouper_window_persistent_v5",
    "_ae_circuit_grouper_window_persistent_v4",
    "_ae_circuit_grouper_window_persistent_v3",
    "_ae_circuit_grouper_window_persistent_v2",
)
_ACTIVE_SCOPE = object()


def _id_value(item):
    return revit_helpers.get_elementid_value(item)


def _id_from(value):
    return revit_helpers.elementid_from_value(value)


def _document_key(doc):
    if doc is None:
        return ""
    try:
        return str(doc.GetHashCode())
    except Exception:
        return "{}|{}".format(
            getattr(doc, "Title", "") or "",
            getattr(doc, "PathName", "") or "",
        )


def _active_uidoc():
    try:
        return __revit__.ActiveUIDocument
    except Exception:
        return getattr(revit, "uidoc", None)


def _active_doc():
    uidoc = _active_uidoc()
    return uidoc.Document if uidoc is not None else getattr(revit, "doc", None)


def _snapshot(doc, stored_scope_id_values=_ACTIVE_SCOPE):
    if doc is None:
        raise Exception("No active document.")

    selected = None
    if stored_scope_id_values is _ACTIVE_SCOPE:
        # Revit selection APIs already return DB.ElementId objects. Keep them
        # native here; the numeric-to-ElementId helper is only for IDs that
        # crossed the modeless UI/external-event payload boundary.
        uidoc = _active_uidoc()
        picked = []
        if uidoc is not None and uidoc.Document == doc:
            try:
                picked = list(uidoc.Selection.GetElementIds())
            except Exception:
                picked = []
        if picked:
            selected = picked
    elif stored_scope_id_values is not None:
        selected = [
            _id_from(value) for value in list(stored_scope_id_values or [])
        ]

    rows_data = cg_collect.collect_devices(doc, selected)
    selected_values = (
        [_id_value(value) for value in selected] if selected is not None else None)
    scope_label = (
        "selection ({} element{})".format(
            len(selected_values), "" if len(selected_values) == 1 else "s")
        if selected_values is not None else "entire model")

    parameter_options = cg_core.common_group_params(rows_data)
    group_param_options = [
        cg_core.SPACE_GROUP_OPTION,
        cg_core.IDENTITY_GROUP_OPTION,
    ] + [p for p in parameter_options if p != "Identity Mark"]
    panel_names, name_to_id, panel_info = cg_collect.collect_panels(doc)

    empty_message = ""
    if not rows_data:
        empty_message = (
            "No circuitable elements were found in the target selection."
            if selected is not None else
            "No circuitable elements were found in the target document.")

    return {
        "document_key": _document_key(doc),
        "document_title": getattr(doc, "Title", "-") or "-",
        "scope_label": scope_label,
        "scope_ids": selected_values,
        "rows_data": rows_data,
        "panel_names": panel_names,
        "name_to_id": name_to_id,
        "panel_info": panel_info,
        "rating_options": cg_core.RATING_OPTIONS,
        "group_param_options": group_param_options,
        "default_group_param": cg_core.default_group_param(group_param_options),
        "name_param_options": parameter_options,
        "default_name_param": cg_core.default_name_param(parameter_options),
        "empty_message": empty_message,
    }


class CircuitGrouperExternalEventGateway(object):
    def __init__(self):
        self._pending = None
        self._handler = _CircuitGrouperExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def is_busy(self):
        try:
            event_pending = bool(self._event.IsPending)
        except Exception:
            event_pending = False
        return self._pending is not None or event_pending

    def _raise(self, op_name, payload=None, callback=None):
        request = {
            "op": str(op_name or ""),
            "payload": dict(payload or {}),
            "callback": callback,
        }
        # Window activation queues a cheap target-document sync. If the same
        # click is also a real user command, let that command replace the sync
        # already waiting in Revit's external-event queue.
        if (self._pending is not None and
                self._pending.get("op") == "sync" and
                request["op"] != "sync"):
            self._pending = request
            return True
        if self.is_busy():
            return False
        self._pending = request
        try:
            self._event.Raise()
            return True
        except Exception:
            self._pending = None
            return False

    def raise_sync(self, document_key, callback):
        return self._raise("sync", {"document_key": document_key}, callback)

    def raise_refresh(self, document_key, scope_ids, callback):
        return self._raise("refresh", {
            "document_key": document_key,
            "scope_ids": list(scope_ids) if scope_ids is not None else None,
        }, callback)

    def raise_navigate(self, document_key, element_ids, show, callback):
        return self._raise("navigate", {
            "document_key": document_key,
            "element_ids": list(element_ids or []),
            "show": bool(show),
        }, callback)

    def raise_run(self, document_key, plans, scope_ids, callback):
        return self._raise("run", {
            "document_key": document_key,
            "plans": list(plans or []),
            "scope_ids": list(scope_ids) if scope_ids is not None else None,
        }, callback)

    def _consume(self):
        pending = self._pending
        self._pending = None
        return pending


class _CircuitGrouperExternalEventHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self._gateway = gateway

    def Execute(self, application):  # noqa: N802
        pending = self._gateway._consume()
        if not pending:
            return
        op_name = pending.get("op")
        payload = pending.get("payload") or {}
        callback = pending.get("callback")
        status, result, error = "ok", None, None
        try:
            uidoc = application.ActiveUIDocument
            doc = uidoc.Document if uidoc is not None else None
            current_key = _document_key(doc)
            if op_name == "sync":
                changed = current_key != str(payload.get("document_key") or "")
                result = {
                    "changed": changed,
                    "snapshot": _snapshot(doc) if changed else None,
                }
            elif op_name == "refresh":
                requested_key = str(payload.get("document_key") or "")
                scope_ids = payload.get("scope_ids")
                result = {
                    "snapshot": _snapshot(
                        doc, scope_ids if current_key == requested_key else _ACTIVE_SCOPE)
                }
            elif op_name == "navigate":
                if current_key != str(payload.get("document_key") or ""):
                    result = {"retarget_snapshot": _snapshot(doc)}
                else:
                    ids = List[DB.ElementId]()
                    for value in payload.get("element_ids") or []:
                        element_id = _id_from(value)
                        if doc.GetElement(element_id) is not None:
                            ids.Add(element_id)
                    if ids.Count == 0:
                        raise Exception("None of the chosen devices exist in the target document.")
                    if bool(payload.get("show")):
                        uidoc.ShowElements(ids)
                    else:
                        uidoc.Selection.SetElementIds(ids)
                    result = {"count": ids.Count, "show": bool(payload.get("show"))}
            elif op_name == "run":
                if current_key != str(payload.get("document_key") or ""):
                    result = {"retarget_snapshot": _snapshot(doc)}
                else:
                    plans = list(payload.get("plans") or [])
                    logger.info("Circuit grouper Run: resolving target panels.")
                    name_to_id = cg_collect.collect_target_panel_ids(
                        doc,
                        [plan.get("panel") for plan in plans],
                    )
                    workflow_group = DB.TransactionGroup(doc, TITLE)
                    workflow_group.Start()
                    try:
                        logger.info("Circuit grouper Run: creating circuits.")
                        report = cg_apply.run(doc, plans, name_to_id, logger)
                        logger.info(
                            "Circuit grouper Run: created %s circuit(s); assigning panels.",
                            int(report.get("created") or 0),
                        )
                        assignment = cg_apply.assign_created_circuits_to_panels(
                            doc, report.get("created_circuit_ids_by_panel", {}),
                            name_to_id, logger)
                        report["panel_assignment"] = assignment
                        workflow_group.Assimilate()
                        # Native ElementIds are only needed by the assignment
                        # service inside this API callback. Do not retain them
                        # in the modeless UI result payload.
                        report.pop("created_circuit_ids_by_panel", None)
                        logger.info("Circuit grouper Run: Revit changes committed.")
                    except Exception:
                        try:
                            workflow_group.RollBack()
                        except Exception:
                            pass
                        raise
                    # Return only the write report. Recollecting every device
                    # and all panel metadata here made Run pay the full Refresh
                    # cost after the transaction had already succeeded.
                    logger.info("Circuit grouper Run: complete.")
                    result = report
            else:
                raise Exception("Unknown operation: {}".format(op_name))
        except Exception as ex:
            status, error = "error", ex
            logger.exception("Circuit grouper external operation failed: %s", ex)
        if callback is not None:
            try:
                callback(status, op_name, result, error)
            except Exception:
                logger.exception("Circuit grouper completion callback failed.")

    def GetName(self):  # noqa: N802
        return "CED Create Circuits by Device Parameter"


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
            marker = str(getattr(window, "Tag", "") or "")
            if marker == _WINDOW_MARKER:
                return window
            if marker in _LEGACY_WINDOW_MARKERS:
                # A prior window/controller can retain an obsolete or inert
                # ExternalEvent after reload. Close it instead of focusing a
                # visually alive window backed by stale code.
                window.Close()
        except Exception:
            pass
    return None


def _focus_existing(window):
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


def main():
    existing = _find_existing_window()
    if existing is not None:
        _focus_existing(existing)
        return

    doc = _active_doc()
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    uidoc = _active_uidoc()
    picked = list(uidoc.Selection.GetElementIds()) if uidoc is not None else []
    if not picked:
        proceed = forms.alert(
            "Nothing is selected, so the tool will scan the ENTIRE model for circuitable elements.\n\n"
            "Run time may increase on a large model. Select devices first for a focused scope.\n\n"
            "Scan the whole model anyway?",
            title=TITLE, yes=True, no=True,
        )
        if not proceed:
            return

    # A non-empty initial selection is still native DB.ElementId data, so let
    # _snapshot read it through the active-selection branch. None explicitly
    # requests the whole-model scope after the user confirms the scan.
    snapshot = _snapshot(doc) if picked else _snapshot(doc, None)
    if not snapshot.get("rows_data"):
        forms.alert(snapshot.get("empty_message"), title=TITLE)
        return

    logger.debug("%s scope=%s, %d circuitable element(s)",
                 TITLE, snapshot.get("scope_label"),
                 len(snapshot.get("rows_data") or []))
    gateway = CircuitGrouperExternalEventGateway()
    cg_window.show_modeless(snapshot, gateway, activate=True)


main()
