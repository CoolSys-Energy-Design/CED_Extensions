# -*- coding: utf-8 -*-
"""ExternalEvent gateway for the modeless Wire Tools window."""

from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType
from System import EventHandler
from System.Collections.Generic import List
from pyrevit import DB, UI, revit, script

from wire_tools_logic import (
    DeviceSelectionFilter,
    NodeSelectionFilter,
    SCHEME_LABELS,
    SELECTION_RULES,
    SCHEME_WIRE_BY_CIRCUIT,
    SCHEME_WIRE_TO_NODE,
    TAG_EXISTING_SKIP,
    active_view_homerun_ids,
    element_id_from,
    element_id_value,
    electrical_connectors,
    circuits_from_elements,
    connector_type_name,
    is_supported_wire_view,
    is_valid_node,
    run_scheme,
    safe_element_name,
    supported_view_text,
    tag_homeruns,
    valid_device_ids,
    valid_homerun_ids,
    system_type_choices,
    wire_tag_type_choices,
    wire_type_choices,
)

# Keep the detailed reports available for debugging without opening an output
# window during normal Wire Tools use.
ENABLE_REPORTING = False

ACTIVE_DOCUMENT_CHANGED_MESSAGE = (
    "Active document changed.\n\nClose and reopen Wire Tools."
)
FAMILY_DOCUMENT_MESSAGE = (
    "Wire Tools does not work in family documents.\n\n"
    "Close and reopen it from a project document."
)


class WireToolsContextError(Exception):
    """Expected modeless-window context invalidation."""

    pass


def _is_family_document(document):
    try:
        return bool(document.IsFamilyDocument)
    except Exception:
        return False


def _document_key(document):
    try:
        return "{}|{}".format(document.PathName, document.Title)
    except Exception:
        return "{}".format(document)


def _view_name(view):
    try:
        return str(view.Name)
    except Exception:
        return "<Unnamed view>"


def _unique_values(values):
    result = []
    seen_values = set()
    for value in list(values or []):
        try:
            integer_value = int(value)
        except Exception:
            continue
        if integer_value not in seen_values:
            seen_values.add(integer_value)
            result.append(integer_value)
    return result


def _element_ids_from_selection(ui_document):
    return _unique_values([
        element_id_value(element_id)
        for element_id in list(ui_document.Selection.GetElementIds())
    ])


def _system_type_status(invalid_devices, requested_system_type):
    if requested_system_type is None:
        return ""
    type_name = connector_type_name(requested_system_type)
    missing_count = len([
        item for item in list(invalid_devices or [])
        if item.get("category") == "system_type_mismatch"
    ])
    ambiguous_count = len([
        item for item in list(invalid_devices or [])
        if item.get("category") == "ambiguous_connector"
    ])
    no_circuit_count = len([
        item for item in list(invalid_devices or [])
        if item.get("category") == "no_circuit"
    ])
    messages = []
    if missing_count:
        noun = "device" if missing_count == 1 else "devices"
        messages.append(
            "{} selected {} do not have a {} connector.".format(
                missing_count,
                noun,
                type_name,
            )
        )
    if ambiguous_count:
        noun = "device" if ambiguous_count == 1 else "devices"
        messages.append(
            "{} selected {} have multiple matching {} connectors.".format(
                ambiguous_count,
                noun,
                type_name,
            )
        )
    if no_circuit_count:
        noun = "device" if no_circuit_count == 1 else "devices"
        messages.append(
            "{} selected {} do not have a matching {} circuit.".format(
                no_circuit_count,
                noun,
                type_name,
            )
        )
    return " ".join(messages)


class WireToolsExternalEventGateway(object):
    def __init__(self, window, document, ui_application=None):
        self.window = window
        self.document_key = _document_key(document)
        try:
            self.active_view_id = element_id_value(document.ActiveView.Id)
        except Exception:
            self.active_view_id = None
        self.ui_application = ui_application
        self.last_active_view_id = self.active_view_id
        self.lifecycle_handlers = {}
        self.document_closing_handler = None
        self.lifecycle_attached = False
        self.device_ids = []
        self.node_id = None
        self.homerun_ids = []
        self.pending = None
        self.invalid_context = False
        self.handler = _WireToolsHandler(self)
        self.event = UI.ExternalEvent.Create(self.handler)

    def attach_lifecycle(self):
        if self.lifecycle_attached or self.ui_application is None:
            return self.lifecycle_attached
        try:
            self.lifecycle_handlers["view-activated"] = revit.events.add_handler(
                "view-activated",
                self._view_activated,
            )
            self.lifecycle_handlers["doc-changed"] = revit.events.add_handler(
                "doc-changed",
                self._document_changed,
            )
            self.document_closing_handler = EventHandler[
                DB.Events.DocumentClosingEventArgs
            ](self._document_closing)
            self.ui_application.Application.DocumentClosing += self.document_closing_handler
            self.lifecycle_attached = True
            return True
        except Exception as error:
            self.detach_lifecycle()
            self.window.report_gateway_error(
                "Could not subscribe to document lifecycle events: {}".format(error)
            )
            return False

    def detach_lifecycle(self):
        for event_name, event_handler in list(self.lifecycle_handlers.items()):
            if event_handler is None:
                continue
            try:
                revit.events.remove_handler(event_name, event_handler)
            except Exception:
                pass
        self.lifecycle_handlers = {}
        if self.document_closing_handler is not None and self.ui_application is not None:
            try:
                self.ui_application.Application.DocumentClosing -= self.document_closing_handler
            except Exception:
                pass
        self.document_closing_handler = None
        self.lifecycle_attached = False

    def _document_closing(self, sender, args):
        del sender
        try:
            closing_document = args.Document
            if _document_key(closing_document) != self.document_key:
                return
            self.detach_lifecycle()
            self.invalidate_context(ACTIVE_DOCUMENT_CHANGED_MESSAGE)
        except Exception as error:
            script.get_logger().warning(
                "Wire Tools document closing handler failed: {}".format(error)
            )

    def _view_activated(self, sender, args):
        del sender
        try:
            active_view = args.CurrentActiveView
            active_uidoc = self.ui_application.ActiveUIDocument
            if active_uidoc is None:
                return
            active_document = active_uidoc.Document
            if _is_family_document(active_document):
                self.invalidate_context(FAMILY_DOCUMENT_MESSAGE)
                return
            if _document_key(active_document) != self.document_key:
                self.invalidate_context(ACTIVE_DOCUMENT_CHANGED_MESSAGE)
                return
            active_view_id = element_id_value(active_view.Id)
            if active_view_id == self.last_active_view_id:
                return
            self.last_active_view_id = active_view_id
            self.window.receive_result(
                "lifecycle",
                "active_view_changed",
                {
                    "refresh_available": is_supported_wire_view(active_view),
                    "active_view_name": _view_name(active_view),
                },
                None,
            )
        except Exception as error:
            script.get_logger().warning(
                "Wire Tools view lifecycle handler failed: {}".format(error)
            )

    def _document_changed(self, sender, args):
        del sender
        try:
            changed_document = args.GetDocument()
            if _document_key(changed_document) != self.document_key:
                return
            message = self.current_context_message()
            if message:
                self.invalidate_context(message)
        except Exception as error:
            script.get_logger().warning(
                "Wire Tools document lifecycle handler failed: {}".format(error)
            )

    def current_context_message(self):
        if self.ui_application is None:
            return None
        try:
            ui_document = self.ui_application.ActiveUIDocument
            document = ui_document.Document if ui_document is not None else None
        except Exception:
            return ACTIVE_DOCUMENT_CHANGED_MESSAGE
        if document is None:
            return ACTIVE_DOCUMENT_CHANGED_MESSAGE
        if _is_family_document(document):
            return FAMILY_DOCUMENT_MESSAGE
        if _document_key(document) != self.document_key:
            return ACTIVE_DOCUMENT_CHANGED_MESSAGE
        return None

    def current_active_view_supported(self):
        """Return whether the current active view can become the target view."""
        if self.invalid_context or self.ui_application is None:
            return False
        if self.current_context_message() is not None:
            return False
        try:
            ui_document = self.ui_application.ActiveUIDocument
            return bool(ui_document is not None
                        and is_supported_wire_view(ui_document.ActiveView))
        except Exception:
            return False

    def check_context(self):
        if self.invalid_context:
            return False
        message = self.current_context_message()
        if message:
            self.invalidate_context(message)
            return False
        return True

    def invalidate_context(self, message):
        if self.invalid_context:
            return
        self.invalid_context = True
        self.pending = None
        self.detach_lifecycle()
        try:
            self.window.set_invalid_context(message)
        except Exception as error:
            script.get_logger().warning(
                "Could not update Wire Tools invalid-context state: {}".format(error)
            )

    def raise_action(self, action_name, payload=None):
        if self.invalid_context:
            return False
        if not self.check_context():
            return False
        try:
            if self.pending is not None or bool(self.event.IsPending):
                return False
        except Exception:
            if self.pending is not None:
                return False
        self.pending = {
            "action": str(action_name),
            "payload": dict(payload or {}),
        }
        try:
            self.event.Raise()
            return True
        except Exception as error:
            self.pending = None
            self.window.report_gateway_error(
                "Could not queue Wire Tools operation: {}".format(error)
            )
            return False

    def consume(self):
        pending = self.pending
        self.pending = None
        return pending


class _WireToolsHandler(UI.IExternalEventHandler):
    def __init__(self, gateway):
        self.gateway = gateway

    def _document_context(self, application):
        ui_document = application.ActiveUIDocument
        document = ui_document.Document if ui_document is not None else None
        if document is None:
            raise WireToolsContextError(ACTIVE_DOCUMENT_CHANGED_MESSAGE)
        if _is_family_document(document):
            raise WireToolsContextError(FAMILY_DOCUMENT_MESSAGE)
        if _document_key(document) != self.gateway.document_key:
            raise WireToolsContextError(ACTIVE_DOCUMENT_CHANGED_MESSAGE)
        return ui_document, document

    def _context(self, application, require_supported=True):
        ui_document, document = self._document_context(application)
        active_view = ui_document.ActiveView
        view = active_view
        if self.gateway.active_view_id is not None:
            try:
                target_view = document.GetElement(
                    element_id_from(self.gateway.active_view_id)
                )
                if target_view is not None:
                    view = target_view
            except Exception:
                pass
        if require_supported and not is_supported_wire_view(view):
            raise ValueError(
                "Wire Tools requires a floor plan or reflected ceiling plan. "
                "Target view: {}.".format(supported_view_text(view))
            )
        return ui_document, document, view

    def _send(self, status, action_name, result=None, error=None):
        try:
            self.gateway.window.receive_result(status, action_name, result, error)
        except Exception as callback_error:
            script.get_logger().exception(
                "Wire Tools UI callback failed: {}".format(callback_error)
            )

    def _sync_result(self, document, view, scheme, requested_system_type=None):
        type_choices = system_type_choices(document, self.gateway.device_ids)
        available_keys = [item.get("id") for item in type_choices]
        if requested_system_type not in available_keys and type_choices:
            requested_system_type = type_choices[0].get("id")
        valid_devices, invalid_devices = valid_device_ids(
            document,
            self.gateway.device_ids,
            scheme,
            requested_system_type=requested_system_type,
        )
        circuit_count = 0
        if scheme == SCHEME_WIRE_BY_CIRCUIT:
            circuit_count = len(circuits_from_elements(
                document,
                valid_devices,
                requested_system_type=requested_system_type,
            ))
        node_element = None
        node_connector_count = 0
        node_name = "-"
        node_type = "-"
        if self.gateway.node_id:
            node_element = document.GetElement(element_id_from(self.gateway.node_id))
            if node_element is not None and is_valid_node(node_element):
                try:
                    node_connector_count = len(electrical_connectors(node_element))
                except Exception:
                    node_connector_count = 0
                node_name = safe_element_name(node_element)
                try:
                    node_type = safe_element_name(
                        document.GetElement(node_element.GetTypeId())
                    )
                except Exception:
                    node_type = "-"
        valid_homeruns, invalid_homeruns = valid_homerun_ids(
            document,
            self.gateway.homerun_ids,
            view,
        )
        del invalid_homeruns
        system_type_status = _system_type_status(
            invalid_devices,
            requested_system_type,
        )
        return {
            "view_supported": is_supported_wire_view(view),
            "view_name": _view_name(view),
            "refresh_available": self.gateway.current_active_view_supported(),
            "wire_types": wire_type_choices(document),
            "tag_types": wire_tag_type_choices(document),
            "device_count": len(valid_devices),
            "invalid_device_count": len(invalid_devices),
            "selected_device_count": len(_unique_values(self.gateway.device_ids)),
            "system_type_choices": type_choices,
            "system_type_key": requested_system_type,
            "system_type_status": system_type_status,
            "circuit_count": circuit_count,
            "node_id": self.gateway.node_id,
            "node_name": node_name,
            "node_type": node_type,
            "node_connector_count": node_connector_count,
            "homerun_count": len(valid_homeruns),
        }

    def _selection_result(
            self,
            document,
            element_ids,
            scheme,
            requested_system_type=None):
        raw_ids = list(element_ids or [])
        node_value = None
        if self.gateway.node_id is not None:
            try:
                node_value = element_id_value(self.gateway.node_id)
            except Exception:
                node_value = None
        device_raw_ids = [
            raw_id for raw_id in raw_ids
            if node_value is None or element_id_value(raw_id) != node_value
        ]
        type_choices = system_type_choices(document, device_raw_ids)
        available_keys = [item.get("id") for item in type_choices]
        if requested_system_type not in available_keys and type_choices:
            requested_system_type = type_choices[0].get("id")
        diagnostics = []
        valid_elements, invalid_elements = valid_device_ids(
            document,
            device_raw_ids,
            scheme,
            diagnostics=diagnostics,
            requested_system_type=requested_system_type,
        )
        self.gateway.device_ids = _unique_values(device_raw_ids)
        circuit_count = 0
        no_circuit_elements = []
        if scheme == SCHEME_WIRE_BY_CIRCUIT:
            circuit_count = len(circuits_from_elements(
                document,
                valid_elements,
                requested_system_type=requested_system_type,
            ))
            no_circuit_elements = [
                item for item in invalid_elements
                if item.get("category") == "no_circuit"
            ]
        return {
            "raw_count": len(device_raw_ids),
            "unique_count": len(_unique_values(device_raw_ids)),
            "device_count": len(valid_elements),
            "invalid_device_count": len(invalid_elements),
            "selected_device_count": len(_unique_values(device_raw_ids)),
            "system_type_choices": type_choices,
            "system_type_key": requested_system_type,
            "system_type_status": _system_type_status(
                invalid_elements,
                requested_system_type,
            ),
            "circuit_count": circuit_count,
            "invalid": invalid_elements,
            "no_circuit": no_circuit_elements,
            "diagnostics": diagnostics,
            "node_id": self.gateway.node_id,
            "node_excluded": node_value is not None and len(device_raw_ids) != len(raw_ids),
        }

    def Execute(self, application):
        pending = self.gateway.consume()
        if not pending:
            return
        if self.gateway.invalid_context:
            return
        action_name = pending.get("action")
        payload = pending.get("payload") or {}
        try:
            if action_name == "sync":
                ui_document, document, view = self._context(
                    application,
                    require_supported=False,
                )
                del ui_document
                scheme = str(payload.get("scheme") or SCHEME_WIRE_BY_CIRCUIT)
                settings = payload.get("settings") or {}
                self._send(
                    "ok",
                    action_name,
                    self._sync_result(
                        document,
                        view,
                        scheme,
                        settings.get("system_type_key"),
                    ),
                )
                return

            if action_name == "refresh_target_view":
                ui_document, document = self._document_context(application)
                view = ui_document.ActiveView
                if not is_supported_wire_view(view):
                    raise ValueError(
                        "Switch to a floor plan or reflected ceiling plan before "
                        "refreshing the target view."
                    )
                self.gateway.active_view_id = element_id_value(view.Id)
                self.gateway.device_ids = []
                self.gateway.node_id = None
                self.gateway.homerun_ids = []
                scheme = str(payload.get("scheme") or SCHEME_WIRE_BY_CIRCUIT)
                settings = payload.get("settings") or {}
                self._send(
                    "ok",
                    action_name,
                    self._sync_result(
                        document,
                        view,
                        scheme,
                        settings.get("system_type_key"),
                    ),
                )
                return

            ui_document, document, view = self._context(application)
            scheme = str(payload.get("scheme") or SCHEME_WIRE_BY_CIRCUIT)

            if action_name == "select_devices":
                selection_filter = DeviceSelectionFilter(
                    scheme,
                    excluded_element_id=self.gateway.node_id,
                )
                preselected_references = List[DB.Reference]()
                for device_id in list(self.gateway.device_ids or []):
                    try:
                        element = document.GetElement(element_id_from(device_id))
                        if element is not None and selection_filter.AllowElement(element):
                            preselected_references.Add(DB.Reference(element))
                    except Exception:
                        continue
                try:
                    picked_references = ui_document.Selection.PickObjects(
                        ObjectType.Element,
                        selection_filter,
                        "Select main-model MEP devices or circuits, then finish.",
                        preselected_references,
                    )
                except OperationCanceledException:
                    self._send("cancelled", action_name, {
                        "device_count": len(self.gateway.device_ids),
                    })
                    return
                picked_ids = [
                    element_id_value(picked_reference.ElementId)
                    for picked_reference in picked_references
                ]
                result = self._selection_result(
                    document,
                    picked_ids,
                    scheme,
                    (payload.get("settings") or {}).get("system_type_key"),
                )
                if ENABLE_REPORTING:
                    _report_selection(document, result, scheme)
                self._send("ok", action_name, result)
                return

            if action_name == "use_current_selection":
                result = self._selection_result(
                    document,
                    _element_ids_from_selection(ui_document),
                    scheme,
                    (payload.get("settings") or {}).get("system_type_key"),
                )
                if ENABLE_REPORTING:
                    _report_selection(document, result, scheme)
                self._send("ok", action_name, result)
                return

            if action_name == "clear_devices":
                self.gateway.device_ids = []
                self._send("ok", action_name, {
                    "device_count": 0,
                    "invalid_device_count": 0,
                    "selected_device_count": 0,
                    "system_type_choices": [],
                    "system_type_key": None,
                    "system_type_status": "",
                })
                return

            if action_name == "select_node":
                selection_filter = NodeSelectionFilter()
                try:
                    picked_reference = ui_document.Selection.PickObject(
                        ObjectType.Element,
                        selection_filter,
                        "Select one electrical node, then finish.",
                    )
                except OperationCanceledException:
                    self._send("cancelled", action_name, {
                        "node_id": self.gateway.node_id,
                    })
                    return
                node_id = element_id_value(picked_reference.ElementId)
                node_element = document.GetElement(element_id_from(node_id))
                if not is_valid_node(node_element):
                    raise ValueError("The selected node has no usable electrical connector.")
                self.gateway.node_id = node_id
                self._send("ok", action_name, {"node_id": node_id})
                return

            if action_name == "clear_node":
                self.gateway.node_id = None
                self._send("ok", action_name, {"node_id": None})
                return

            if action_name == "select_homeruns":
                action_settings = payload.get("settings") or {}
                skip_tagged = (
                    action_settings.get("existing_tag_behavior")
                    == TAG_EXISTING_SKIP
                )
                self.gateway.homerun_ids = _unique_values(
                    active_view_homerun_ids(
                        document,
                        view,
                        skip_tagged=skip_tagged,
                    )
                )
                try:
                    selected_ids = List[DB.ElementId]([
                        element_id_from(wire_value)
                        for wire_value in self.gateway.homerun_ids
                    ])
                    ui_document.Selection.SetElementIds(selected_ids)
                except Exception as error:
                    script.get_logger().warning(
                        "Homerun wires were indexed but could not be selected in Revit: {}".format(
                            error
                        )
                    )
                self._send("ok", action_name, {
                    "homerun_count": len(self.gateway.homerun_ids),
                })
                return

            if action_name == "clear_homeruns":
                self.gateway.homerun_ids = []
                self._send("ok", action_name, {"homerun_count": 0})
                return

            if action_name == "run_scheme":
                valid_elements, invalid_elements = valid_device_ids(
                    document,
                    self.gateway.device_ids,
                    scheme,
                    requested_system_type=(payload.get("settings") or {}).get(
                        "system_type_key"
                    ),
                )
                if not valid_elements:
                    raise ValueError("Select at least one valid electrical device first.")
                node_element = None
                if scheme == SCHEME_WIRE_TO_NODE:
                    if not self.gateway.node_id:
                        raise ValueError("Select a node before running Wire to Node.")
                    node_element = document.GetElement(
                        element_id_from(self.gateway.node_id)
                    )
                    if not is_valid_node(node_element):
                        raise ValueError("The selected node is no longer valid.")
                settings = dict(payload.get("settings") or {})
                result = run_scheme(
                    document,
                    view,
                    scheme,
                    valid_elements,
                    node_element,
                    settings,
                )
                result["invalid_devices"] = invalid_elements
                if ENABLE_REPORTING:
                    _report_operation(document, result)
                self._send("ok", action_name, result)
                return

            if action_name == "tag_homeruns":
                valid_ids, invalid_homeruns = valid_homerun_ids(
                    document,
                    self.gateway.homerun_ids,
                    view,
                )
                if not valid_ids:
                    raise ValueError(
                        "Use Select Homeruns to find open-ended wires in the active view first."
                    )
                settings = dict(payload.get("settings") or {})
                result = tag_homeruns(
                    document,
                    view,
                    valid_ids,
                    settings.get("tag_type_id"),
                    bool(settings.get("add_leaders", True)),
                    settings.get("existing_tag_behavior"),
                )
                result["selected_count"] = len(valid_ids)
                result["invalid"] = invalid_homeruns
                if ENABLE_REPORTING:
                    _report_tagging(document, result)
                self._send("ok", action_name, result)
                return

            raise ValueError("Unknown Wire Tools action: {}".format(action_name))
        except WireToolsContextError as error:
            self.gateway.invalidate_context(str(error))
            return
        except Exception as error:
            script.get_logger().exception(
                "Wire Tools operation failed: {}".format(error)
            )
            self._send("error", action_name, None, error)

    def GetName(self):
        return "CED Wire Tools External Event"


def _report_element(output, document, item, prefix):
    element = item.get("element")
    if element is not None:
        try:
            label = output.linkify(element.Id)
        except Exception:
            label = str(item.get("id", "?"))
        name = safe_element_name(element)
    else:
        label = str(item.get("id", "?"))
        name = "<unknown>"
    output.print_md(
        "- {} {} ({}) - {}".format(
            prefix,
            label,
            name,
            item.get("reason", name),
        )
    )


def _report_selection_detail(output, document, detail):
    element = detail.get("element")
    normalized_id = detail.get("normalized_id")
    if element is not None:
        try:
            label = output.linkify(element.Id)
        except Exception:
            label = str(normalized_id)
        name = detail.get("name", safe_element_name(element))
    else:
        label = "raw {} / normalized {}".format(
            detail.get("raw_id", "?"),
            normalized_id if normalized_id is not None else "<none>",
        )
        name = "<unresolved>"
    outcome = "PASS" if detail.get("accepted") else "FAIL"
    output.print_md(
        "- **{}** {} ({}) - final stage: **{}** - {}".format(
            outcome,
            label,
            name,
            detail.get("final_stage", "unknown"),
            detail.get("reason", "No final reason recorded."),
        )
    )
    for step in detail.get("steps", []):
        step_state = "PASS" if step.get("passed") else "FAIL"
        output.print_md(
            "  - {} **{}**: {}".format(
                step_state,
                step.get("stage", "unknown"),
                step.get("message", ""),
            )
        )
    resolution_steps = detail.get("resolution_steps", [])
    if resolution_steps:
        output.print_md("  - Circuit resolver trace:")
        for step in resolution_steps:
            step_state = "PASS" if step.get("passed") else "FAIL"
            output.print_md(
                "    - {} **{}**: {}".format(
                    step_state,
                    step.get("stage", "unknown"),
                    step.get("message", ""),
                )
            )
    connector_count = detail.get("connector_count", 0)
    if connector_count:
        output.print_md(
            "  - Connector summary: **{}** usable connector(s); types: {}.".format(
                connector_count,
                ", ".join(detail.get("connector_types", []))
                or "none",
            )
        )
    circuit_ids = detail.get("circuit_ids", [])
    if circuit_ids:
        output.print_md(
            "  - Circuit summary: **{}** eligible circuit(s); IDs: {}.".format(
                len(circuit_ids),
                ", ".join([str(value) for value in circuit_ids]),
            )
        )
        output.print_md(
            "    - System types: {}.".format(
                ", ".join(detail.get("circuit_types", [])) or "<unavailable>"
            )
        )


def _report_selection(document, result, scheme):
    output = script.get_output()
    output.print_md("## Wire Tools selection diagnostics")
    output.print_md("- Scheme: **{}**".format(
        SCHEME_LABELS.get(scheme, scheme)
    ))
    output.print_md("### Selection rules used")
    output.print_md("- {}".format(
        SELECTION_RULES.get(scheme, "No selection rule was registered for this scheme.")
    ))
    output.print_md("- Raw ElementIds received: **{}**".format(
        result.get("raw_count", 0)
    ))
    output.print_md("- Unique ElementIds processed: **{}**".format(
        result.get("unique_count", 0)
    ))
    output.print_md("- Accepted selection candidates: **{}**".format(
        result.get("device_count", 0)
    ))
    output.print_md("- Rejected ElementIds: **{}**".format(
        result.get("invalid_device_count", 0)
    ))
    if scheme == SCHEME_WIRE_BY_CIRCUIT:
        output.print_md("- Eligible circuits resolved: **{}**".format(
            result.get("circuit_count", 0)
        ))
        output.print_md("- Candidates without an eligible circuit: **{}**".format(
            len(result.get("no_circuit", []))
        ))
    output.print_md("### Per-element validation trace")
    for detail in result.get("diagnostics", []):
        _report_selection_detail(output, document, detail)
    for item in result.get("invalid", []):
        if item.get("category") != "no_circuit":
            _report_element(output, document, item, "Rejected")
    for item in result.get("no_circuit", []):
        _report_element(output, document, item, "No circuit")
    if not result.get("invalid", []) and not result.get("no_circuit", []):
        output.print_md("- No selection rejections.")
    output.show()


def _report_operation(document, result):
    output = script.get_output()
    output.print_md("## Wire Tools report")
    output.print_md("- Scheme: **{}**".format(result.get("scheme", "-")))
    output.print_md("- Wires created: **{}**".format(result.get("created", 0)))
    output.print_md("- Homeruns created: **{}**".format(
        len(result.get("homeruns", []))
    ))
    output.print_md("- Existing wires deleted: **{}**".format(result.get("deleted", 0)))
    output.print_md("- Skipped: **{}**".format(len(result.get("skipped", []))))
    output.print_md("- Failures: **{}**".format(len(result.get("failures", []))))
    for item in result.get("invalid_devices", []):
        _report_element(output, document, item, "Invalid device")
    for item in result.get("skipped", []):
        _report_element(output, document, item, "Skipped")
    for item in result.get("failures", []):
        _report_element(output, document, item, "Failure")
    output.show()


def _report_tagging(document, result):
    output = script.get_output()
    output.print_md("## Wire Tools homerun-tagging report")
    output.print_md("- Homeruns selected: **{}**".format(result.get("selected_count", 0)))
    output.print_md("- Tags created: **{}**".format(result.get("created", 0)))
    output.print_md("- Existing tags deleted: **{}**".format(result.get("deleted", 0)))
    output.print_md("- Already-tagged wires skipped: **{}**".format(
        len(result.get("skipped", []))
    ))
    output.print_md("- Failures: **{}**".format(len(result.get("failures", []))))
    for item in result.get("invalid", []):
        _report_element(output, document, item, "Invalid homerun")
    for item in result.get("failures", []):
        _report_element(output, document, item, "Failure")
    for item in result.get("skipped", []):
        _report_element(output, document, item, "Skipped")
    output.show()
