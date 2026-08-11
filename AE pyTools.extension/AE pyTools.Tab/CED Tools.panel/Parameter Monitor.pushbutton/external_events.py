# -*- coding: utf-8 -*-
"""Single modeless-UI gateway for all Revit API and project-store actions."""

from __future__ import print_function

import copy
import io
import traceback

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from pyrevit import forms, script

import element_linker_sync_service
import models
import navigation_service
import parameter_service
import parent_move_service
import relationship_service
import set_io
import source_service
import storage_service
import tracking_service

TITLE = "Parameter Monitor"


class _Choice(forms.TemplateListItem):
    def __init__(self, item, label, checked=False):
        self._label = str(label or "")
        forms.TemplateListItem.__init__(self, item, checked=checked)

    @property
    def name(self):
        return self._label


def _unwrap(item):
    if item is None:
        return None
    return getattr(item, "item", item)


def _same_document(left, right):
    if left is None or right is None:
        return False
    try:
        return bool(left.Equals(right))
    except Exception:
        return left is right


def _source_choices(document, source_hint=None):
    sources = source_service.list_sources(document, include_unavailable=False)
    hint = source_hint or {}
    preferred = []
    remainder = []
    for source in sources:
        same_type = source.get("source_type") == hint.get("source_type")
        same_name = (
            str(source.get("display_name") or "").strip().lower()
            == str(hint.get("display_name") or "").strip().lower()
        )
        choice = _Choice(source, source.get("display_name") or "Source Model")
        if same_type and (same_name or not hint.get("display_name")):
            preferred.append(choice)
        else:
            remainder.append(choice)
    return preferred + remainder


def _choose_source(document, source_hint=None, title="Select Source Model"):
    choices = _source_choices(document, source_hint=source_hint)
    if not choices:
        forms.alert("No readable host or linked model source is available.", title=TITLE)
        return None
    return _unwrap(forms.SelectFromList.show(
        choices,
        title=title,
        button_name="Use Source",
        multiselect=False,
    ))


def _choose_category(document, source):
    resolved = source_service.resolve_source(document, source)
    if not resolved.get("available"):
        forms.alert(resolved.get("message") or "Source model is unavailable.", title=TITLE)
        return None, None
    categories = source_service.list_categories(resolved.get("source_document"))
    choices = [_Choice(item, item.get("name") or "Category") for item in categories]
    selected = _unwrap(forms.SelectFromList.show(
        choices,
        title="Select Model Category",
        button_name="Use Category",
        multiselect=False,
    ))
    return selected, resolved


def _parameter_label(descriptor):
    scope = "Type" if descriptor.get("scope") == "type" else "Instance"
    return "{} [{}]  {} / {}".format(
        descriptor.get("name") or "Unnamed Parameter",
        scope,
        int(descriptor.get("available_count", 0) or 0),
        int(descriptor.get("element_count", 0) or 0),
    )


def _choose_parameters(source_document, category, selected_keys=None, elements=None):
    if elements is None:
        elements = source_service.collect_elements(source_document, category)
    descriptors = parameter_service.discover_parameters(elements, source_document)
    selected_keys = set(selected_keys or [])
    choices = [
        _Choice(
            descriptor,
            _parameter_label(descriptor),
            checked=descriptor.get("key") in selected_keys,
        )
        for descriptor in descriptors
    ]
    selected = forms.SelectFromList.show(
        choices,
        title="Select Tracked Parameters",
        button_name="Track Parameters",
        multiselect=True,
    )
    if selected is None:
        return None
    return [_unwrap(item) for item in list(selected or [])]


def _prompt_set_name(default_name):
    value = forms.ask_for_string(
        default=str(default_name or "Tracking Set"),
        prompt="Tracking Set name:",
        title=TITLE,
    )
    if value is None:
        return None
    value = str(value).strip()
    return value or str(default_name or "Tracking Set")


def _create_set_interactive(document, store, logger):
    source = _choose_source(document)
    if source is None:
        return None
    category, resolved = _choose_category(document, source)
    if category is None:
        return None
    descriptors = _choose_parameters(resolved.get("source_document"), category)
    if descriptors is None:
        return None
    name = _prompt_set_name(category.get("name") or "Tracking Set")
    if name is None:
        return None
    track_location = bool(forms.alert(
        "Track LocationPoint position and rotation for new elements in this set?\n\n"
        "This can be changed per element later.",
        title=TITLE,
        yes=True,
        no=True,
    ))
    defaults = {
        "track_new_elements": track_location,
        "translation_tolerance": 0.001,
        "angular_tolerance": 0.0017453292519943296,
    }
    updated_store, tracking_set = tracking_service.create_tracking_set(
        document,
        store,
        name,
        source,
        category,
        descriptors,
        location_defaults=defaults,
        logger=logger,
    )
    return updated_store, "Created {} with {} baseline element(s).".format(
        tracking_set.get("name"), len(tracking_set.get("elements") or {})
    )


def _edit_set_interactive(document, store, set_id, logger):
    tracking_set = models.find_set(store, set_id)
    if tracking_set is None:
        raise ValueError("Tracking Set no longer exists.")
    resolved = source_service.resolve_source(document, tracking_set.get("source") or {})
    if not resolved.get("available"):
        raise ValueError(resolved.get("message") or "Source model is unavailable.")
    selected_keys = [item.get("key") for item in tracking_set.get("tracked_properties") or []]
    explicit_elements = None
    if str(tracking_set.get("membership") or "") == models.MEMBERSHIP_EXPLICIT:
        explicit_elements = source_service.collect_set_elements(
            resolved.get("source_document"), tracking_set
        )
    descriptors = _choose_parameters(
        resolved.get("source_document"),
        tracking_set.get("category") or {},
        selected_keys=selected_keys,
        elements=explicit_elements,
    )
    if descriptors is None:
        return None
    name = _prompt_set_name(tracking_set.get("name"))
    if name is None:
        return None
    current_default = bool(
        (tracking_set.get("location_defaults") or {}).get("track_new_elements", False)
    )
    track_new_elements = bool(forms.alert(
        "Track location by default for future Added elements?\n\n"
        "Current default: {}. Existing per-element choices are not changed.".format(
            "On" if current_default else "Off"
        ),
        title=TITLE,
        yes=True,
        no=True,
    ))
    updated_store, updated_set = tracking_service.edit_tracking_set(
        document,
        store,
        set_id,
        name,
        descriptors,
        active=tracking_set.get("active", True),
        track_new_elements=track_new_elements,
        logger=logger,
    )
    return updated_store, "Updated {}. Newly added properties were accepted immediately.".format(
        updated_set.get("name")
    )


def _export_definitions(store, preferred_set_id=None):
    tracking_sets = list(store.get("tracking_sets") or [])
    if not tracking_sets:
        forms.alert("There are no Tracking Sets to export.", title=TITLE)
        return None
    choices = []
    for tracking_set in tracking_sets:
        choices.append(_Choice(
            tracking_set,
            tracking_set.get("name") or "Tracking Set",
            checked=(tracking_set.get("set_id") == preferred_set_id),
        ))
    selected = forms.SelectFromList.show(
        choices,
        title="Export Tracking Set Definitions",
        button_name="Export Selected",
        multiselect=True,
    )
    selected_sets = [_unwrap(item) for item in list(selected or [])]
    if not selected_sets:
        return None
    path = forms.save_file(
        file_ext="json",
        default_name="Parameter_Monitor_Tracking_Sets.json",
        title="Export Tracking Set Definitions",
    )
    if not path:
        return None
    set_io.dump_file(path, selected_sets)
    return "Exported {} Tracking Set definition(s) to {}.".format(len(selected_sets), path)


def _map_imported_descriptors(source_document, category, imported_descriptors):
    elements = source_service.collect_elements(source_document, category)
    available = parameter_service.discover_parameters(elements, source_document)
    mapped = []
    unresolved = []
    for imported in list(imported_descriptors or []):
        match = None
        for candidate in available:
            if parameter_service.descriptor_matches(imported, candidate):
                match = candidate
                break
        if match is None:
            unresolved.append(imported)
            mapped.append(copy.deepcopy(imported))
        else:
            mapped.append(copy.deepcopy(match))
    return mapped, unresolved


def _import_definitions(document, store, logger):
    path = forms.pick_file(file_ext="json", title="Import Tracking Set Definitions")
    if not path:
        return None
    definitions = set_io.load_file(path)
    if not definitions:
        forms.alert("The file contains no Tracking Set definitions.", title=TITLE)
        return None
    result = copy.deepcopy(store)
    imported_count = 0
    skipped = []
    for definition in definitions:
        source = _choose_source(
            document,
            source_hint=definition.get("source_hint") or {},
            title="Map Source for '{}'".format(definition.get("name") or "Tracking Set"),
        )
        if source is None:
            skipped.append(definition.get("name") or "Tracking Set")
            continue
        resolved = source_service.resolve_source(document, source)
        category = source_service.resolve_category(
            resolved.get("source_document"),
            definition.get("category") or {},
        )
        if category is None:
            forms.alert(
                "Category '{}' could not be resolved in {}. This definition will be skipped.".format(
                    (definition.get("category") or {}).get("name") or "Unknown",
                    source.get("display_name") or "source model",
                ),
                title=TITLE,
            )
            skipped.append(definition.get("name") or "Tracking Set")
            continue
        descriptors, unresolved = _map_imported_descriptors(
            resolved.get("source_document"),
            category,
            definition.get("tracked_properties") or [],
        )
        if unresolved:
            names = ", ".join([item.get("name") or item.get("key") for item in unresolved])
            proceed = forms.alert(
                "These parameters could not be resolved and will be tracked as Missing:\n\n{}\n\n"
                "Import this Tracking Set anyway?".format(names),
                title=TITLE,
                yes=True,
                no=True,
            )
            if not proceed:
                skipped.append(definition.get("name") or "Tracking Set")
                continue
        result, created = tracking_service.create_tracking_set(
            document,
            result,
            definition.get("name"),
            source,
            category,
            descriptors,
            location_defaults=definition.get("location_defaults") or None,
            logger=logger,
        )
        created["active"] = bool(definition.get("active", True))
        result = tracking_service._replace_set(result, created)
        imported_count += 1
    message = "Imported {} Tracking Set definition(s) with fresh baselines.".format(imported_count)
    if skipped:
        message += " Skipped: {}.".format(", ".join(skipped))
    return result, message


def _csv_cell(value):
    text = str(value if value is not None else "")
    if any(marker in text for marker in (",", "\"", "\n", "\r")):
        return "\"{}\"".format(text.replace("\"", "\"\""))
    return text


def _require_resolvable_set(store, set_id):
    tracking_set = models.find_set(store, set_id)
    if tracking_set is None:
        raise ValueError("Select a Tracking Set first.")
    if tracking_set.get("status") in (
        models.SET_SOURCE_UNAVAILABLE,
        models.SET_CHECK_FAILED,
    ):
        raise ValueError(
            "Run a successful Scan Selected Set before resolving this Tracking Set."
        )
    return tracking_set


def _metadata_report_value(metadata, field):
    metadata = metadata or {}
    direct = metadata.get(field)
    if direct is not None and str(direct) != "":
        return direct
    parts = [item.strip() for item in str(metadata.get("family_type") or "").split(" : ", 1)]
    if field == "family_name" and parts:
        return parts[0]
    if field == "type_name" and len(parts) > 1:
        return parts[1]
    if field == "type_name" and parts:
        return parts[0]
    return ""


def _export_report(store, set_id):
    tracking_set = models.find_set(store, set_id)
    if tracking_set is None:
        raise ValueError("Select a Tracking Set first.")
    default_name = "{}_Parameter_Monitor_Report.csv".format(
        str(tracking_set.get("name") or "Tracking_Set").replace(" ", "_")
    )
    path = forms.save_file(file_ext="csv", default_name=default_name, title="Export Review Report")
    if not path:
        return None
    rows = [[
        "Tracking Set", "Element State", "Persistent Identity", "Element", "ElementId",
        "Family : Type", "Level", "Property", "Accepted", "Current", "Value State",
    ]]
    descriptors = dict([
        (item.get("key"), item) for item in tracking_set.get("tracked_properties") or []
    ])
    for persistent_id, record in sorted((tracking_set.get("elements") or {}).items()):
        metadata = record.get("metadata") or {}
        keys = list(record.get("changed_property_keys") or [])
        if record.get("state") in (models.ELEMENT_ADDED, models.ELEMENT_REMOVED) and not keys:
            keys = [""]
        if not keys and int(record.get("missing_count", 0) or 0) > 0:
            keys = [
                key for key, value in (record.get("current_properties") or {}).items()
                if (value or {}).get("state") == models.VALUE_MISSING
            ]
        for key in keys:
            if key == models.LOCATION_PROPERTY_KEY:
                prop_name = "Location"
                accepted = record.get("accepted_location")
                current = record.get("current_location")
                value_state = (current or {}).get("state")
            elif key in (models.FAMILY_PROPERTY_KEY, models.TYPE_PROPERTY_KEY):
                field = "family_name" if key == models.FAMILY_PROPERTY_KEY else "type_name"
                prop_name = "Family" if key == models.FAMILY_PROPERTY_KEY else "Type"
                accepted_metadata = record.get("accepted_metadata") or record.get("metadata") or {}
                current_metadata = record.get("current_metadata") or record.get("metadata") or {}
                accepted = _metadata_report_value(accepted_metadata, field)
                current = _metadata_report_value(current_metadata, field)
                value_state = "changed"
            elif key:
                prop_name = (descriptors.get(key) or {}).get("name") or key
                accepted_value = (record.get("accepted_properties") or {}).get(key) or {}
                current_value = (record.get("current_properties") or {}).get(key) or {}
                accepted = accepted_value.get("display")
                current = current_value.get("display")
                value_state = current_value.get("state")
            else:
                prop_name = ""
                accepted = ""
                current = ""
                value_state = ""
            rows.append([
                tracking_set.get("name"), record.get("state"), persistent_id,
                metadata.get("friendly_name"), metadata.get("element_id"),
                metadata.get("family_type"), metadata.get("level"), prop_name,
                accepted, current, value_state,
            ])
    for persistent_id in tracking_set.get("untracked_ids") or []:
        rows.append([
            tracking_set.get("name"), "untracked", persistent_id, "", "", "", "", "", "", "", "",
        ])
    with io.open(path, "w", encoding="utf-8") as stream:
        for row in rows:
            stream.write(",".join([_csv_cell(cell) for cell in row]) + "\n")
    return "Exported review report to {}.".format(path)


class ParameterMonitorExternalEventGateway(object):
    def __init__(self, document, logger=None):
        self.document = document
        self.logger = logger or script.get_logger()
        self._pending = None
        self._handler = _ParameterMonitorExternalEventHandler(self)
        self._event = ExternalEvent.Create(self._handler)

    def is_busy(self):
        if self._pending is not None:
            return True
        try:
            return bool(self._event.IsPending)
        except Exception:
            return False

    def raise_action(self, operation, payload=None, callback=None):
        if self.is_busy():
            return False
        self._pending = {
            "operation": str(operation or ""),
            "payload": copy.deepcopy(payload or {}),
            "callback": callback,
        }
        try:
            self._event.Raise()
            return True
        except Exception:
            self._pending = None
            raise

    def _consume(self):
        pending = self._pending
        self._pending = None
        return pending


class _ParameterMonitorExternalEventHandler(IExternalEventHandler):
    def __init__(self, gateway):
        self.gateway = gateway

    def Execute(self, application):  # noqa: N802
        pending = self.gateway._consume()
        if not pending:
            return
        operation = pending.get("operation")
        payload = pending.get("payload") or {}
        callback = pending.get("callback")
        status = "ok"
        result = None
        error = None
        try:
            uidocument = application.ActiveUIDocument
            document = uidocument.Document if uidocument is not None else None
            if not _same_document(document, self.gateway.document):
                if operation == "refresh_store" and document is not None:
                    # Reload Project Data is the recovery path after a
                    # document switch: re-target the gateway to the active
                    # document and load its monitor data.
                    self.gateway.document = document
                else:
                    raise ValueError(
                        "The active document changed. Click 'Reload Project "
                        "Data' to re-target Parameter Monitor to the active "
                        "project."
                    )
            result = self._execute_operation(document, uidocument, operation, payload)
            if result is None:
                status = "cancelled"
        except Exception as ex:
            status = "error"
            error = ex
            try:
                error._parameter_monitor_traceback = traceback.format_exc()
            except Exception:
                pass
            try:
                self.gateway.logger.exception(
                    "Parameter Monitor external operation failed: %s", operation
                )
            except Exception:
                pass
        if callback is not None:
            try:
                callback(status, operation, result, error)
            except Exception:
                pass

    def _save_result(self, document, store, message, transaction_name):
        saved = storage_service.save(
            document,
            store,
            transaction_name=transaction_name,
            logger=self.gateway.logger,
        )
        return {"store": saved, "message": message}

    def _execute_operation(self, document, uidocument, operation, payload):
        store = storage_service.load(document, logger=self.gateway.logger)
        set_id = payload.get("set_id")
        persistent_id = payload.get("persistent_id")
        persistent_ids = [
            str(item or "") for item in list(payload.get("persistent_ids") or []) if item
        ]
        if not persistent_ids and persistent_id:
            persistent_ids = [str(persistent_id)]

        if operation == "refresh_store":
            return {"store": store, "message": "Project monitor data refreshed."}
        if operation == "add_set":
            created = _create_set_interactive(document, store, self.gateway.logger)
            if created is None:
                return None
            updated, message = created
            return self._save_result(document, updated, message, "Parameter Monitor - Add Set")
        if operation == "edit_set":
            edited = _edit_set_interactive(document, store, set_id, self.gateway.logger)
            if edited is None:
                return None
            updated, message = edited
            return self._save_result(document, updated, message, "Parameter Monitor - Edit Set")
        if operation == "delete_set":
            tracking_set = models.find_set(store, set_id)
            if tracking_set is None:
                raise ValueError("Select a Tracking Set first.")
            if not forms.alert(
                "Delete '{}' and all of its Parameter Monitor baseline data?".format(
                    tracking_set.get("name") or "Tracking Set"
                ),
                title=TITLE,
                yes=True,
                no=True,
            ):
                return None
            updated = tracking_service.delete_tracking_set(store, set_id)
            return self._save_result(document, updated, "Tracking Set deleted.", "Parameter Monitor - Delete Set")
        if operation == "toggle_active":
            updated, tracking_set = tracking_service.toggle_set_active(store, set_id)
            message = "{} is now {}.".format(
                tracking_set.get("name"),
                "active" if tracking_set.get("active") else "inactive",
            )
            return self._save_result(document, updated, message, "Parameter Monitor - Toggle Set")
        if operation == "check_set":
            updated, tracking_set = tracking_service.scan_tracking_set(
                document, store, set_id, logger=self.gateway.logger
            )
            return self._save_result(
                document,
                updated,
                tracking_set.get("status_message") or "Check complete.",
                "Parameter Monitor - Check Set",
            )
        if operation == "check_all":
            updated, scanned_ids = tracking_service.scan_all_active(
                document, store, logger=self.gateway.logger
            )
            return self._save_result(
                document,
                updated,
                "Checked {} active Tracking Set(s).".format(len(scanned_ids)),
                "Parameter Monitor - Check All",
            )
        if operation == "resolve_property":
            _require_resolvable_set(store, set_id)
            updated, _tracking_set = tracking_service.resolve_property(
                store, set_id, persistent_id, payload.get("property_key")
            )
            return self._save_result(document, updated, "Property accepted.", "Parameter Monitor - Resolve Property")
        if operation == "resolve_element":
            _require_resolvable_set(store, set_id)
            updated, _tracking_set = tracking_service.resolve_element(store, set_id, persistent_id)
            return self._save_result(document, updated, "Element changes accepted.", "Parameter Monitor - Resolve Element")
        if operation == "resolve_set":
            tracking_set = _require_resolvable_set(store, set_id)
            if not forms.alert(
                "Accept all current parameter/location changes and Added elements in '{}'?\n\n"
                "Removed records will remain until explicitly removed.".format(tracking_set.get("name")),
                title=TITLE,
                yes=True,
                no=True,
            ):
                return None
            updated, _tracking_set = tracking_service.resolve_set(store, set_id)
            return self._save_result(document, updated, "Tracking Set changes accepted.", "Parameter Monitor - Resolve Set")
        if operation in ("untrack_element", "untrack_elements"):
            count = len(persistent_ids)
            if count <= 0:
                raise ValueError("Select one or more tracked elements first.")
            if not forms.alert(
                "Untrack {} selected element(s)? Their accepted data and device relationships "
                "will be discarded; only their identities are retained.".format(count),
                title=TITLE,
                yes=True,
                no=True,
            ):
                return None
            updated, _tracking_set = tracking_service.untrack_elements(
                store,
                set_id,
                persistent_ids,
            )
            return self._save_result(
                document,
                updated,
                "{} element(s) untracked.".format(count),
                "Parameter Monitor - Untrack Elements",
            )
        if operation == "restore_element":
            updated, _tracking_set = tracking_service.restore_element(document, store, set_id, persistent_id)
            return self._save_result(document, updated, "Element restored with a fresh baseline.", "Parameter Monitor - Restore Element")
        if operation == "remove_record":
            if not forms.alert(
                "Permanently forget this removed record? The identity will not be added to Untracked.",
                title=TITLE,
                yes=True,
                no=True,
            ):
                return None
            updated, _tracking_set = tracking_service.remove_record(store, set_id, persistent_id)
            return self._save_result(document, updated, "Removed record deleted.", "Parameter Monitor - Remove Record")
        if operation == "remove_all_removed":
            if not forms.alert(
                "Permanently forget every removed record in this Tracking Set?",
                title=TITLE,
                yes=True,
                no=True,
            ):
                return None
            updated, _tracking_set = tracking_service.remove_all_removed(store, set_id)
            return self._save_result(document, updated, "All removed records deleted.", "Parameter Monitor - Remove Removed Records")
        if operation == "toggle_location":
            if not persistent_ids:
                raise ValueError("Select one or more tracked elements first.")
            updated, _tracking_set = tracking_service.set_elements_location_tracking(
                document,
                store,
                set_id,
                persistent_ids,
                bool(payload.get("enabled")),
            )
            return self._save_result(
                document,
                updated,
                "Location tracking updated for {} element(s).".format(len(persistent_ids)),
                "Parameter Monitor - Location Tracking",
            )
        if operation in ("location_all_on", "location_all_off"):
            enabled = operation == "location_all_on"
            verb = "enable" if enabled else "disable"
            if not forms.alert(
                "{} location tracking for every available element in this Tracking Set?\n\n"
                "Enabling creates a fresh accepted location. Disabling removes stored location data.".format(
                    verb.title()
                ),
                title=TITLE,
                yes=True,
                no=True,
            ):
                return None
            updated, _tracking_set = tracking_service.set_all_element_location_tracking(
                document, store, set_id, enabled
            )
            return self._save_result(
                document,
                updated,
                "Location tracking {}d for available elements.".format(verb),
                "Parameter Monitor - Bulk Location Tracking",
            )
        if operation == "add_device_child":
            device = relationship_service.pick_device(uidocument)
            if device is None:
                return None
            updated, _tracking_set, record = tracking_service.add_manual_child(
                document, store, set_id, persistent_id, device
            )
            label = ((record.get("metadata") or {}).get("friendly_name")
                     or "Device")
            return self._save_result(
                document,
                updated,
                "{} added as a Manual linked child.".format(label),
                "Parameter Monitor - Add Device Child",
            )
        if operation == "unlink_child":
            child_key = str(payload.get("child_persistent_id") or "")
            if not child_key:
                raise ValueError("Select a linked child first.")
            if not forms.alert(
                "Unlink this Manual child? It will no longer be monitored.",
                title=TITLE,
                yes=True,
                no=True,
            ):
                return None
            updated, _tracking_set = tracking_service.remove_manual_child(
                store, set_id, child_key
            )
            return self._save_result(
                document,
                updated,
                "Linked child removed from the monitor.",
                "Parameter Monitor - Unlink Child",
            )
        if operation in ("select_element", "show_element"):
            tracking_set = models.find_set(store, set_id)
            records = [
                (tracking_set.get("elements") or {}).get(key)
                for key in persistent_ids
            ] if tracking_set else []
            records = [record for record in records if record is not None]
            if tracking_set is None or not records:
                raise ValueError("Select one or more available monitored elements first.")
            success = navigation_service.select_tracked_many(
                uidocument,
                tracking_set,
                records,
            ) if operation == "select_element" else navigation_service.show_tracked_many(
                uidocument,
                tracking_set,
                records,
            )
            if not success:
                raise ValueError("The selected elements cannot be navigated to in the current source state.")
            return {"message": "Model navigation updated for {} element(s).".format(len(records))}
        if operation in ("select_device", "show_device"):
            tracking_set = models.find_set(store, set_id)
            record = (tracking_set.get("elements") or {}).get(persistent_id) if tracking_set else None
            relationship = (record or {}).get("relationship") or {}
            unique_id = relationship.get("device_unique_id")
            success = navigation_service.select_host_unique_id(uidocument, unique_id) \
                if operation == "select_device" else navigation_service.show_host_unique_id(uidocument, unique_id)
            if not success:
                raise ValueError("The linked host device is unavailable.")
            return {"store": store, "message": "Device navigation updated."}
        if operation == "select_circuit":
            tracking_set = models.find_set(store, set_id)
            record = (tracking_set.get("elements") or {}).get(persistent_id) if tracking_set else None
            circuits = list(((record or {}).get("relationship_context") or {}).get("circuits") or [])
            if not circuits:
                raise ValueError("The linked device has no current circuit.")
            circuit = circuits[0]
            if len(circuits) > 1:
                choices = [_Choice(
                    item,
                    "{} - {}".format(item.get("panel_name") or "Panel", item.get("circuit_number") or item.get("circuit_name")),
                ) for item in circuits]
                circuit = _unwrap(forms.SelectFromList.show(
                    choices,
                    title="Select Current Circuit",
                    button_name="Select Circuit",
                    multiselect=False,
                ))
                if circuit is None:
                    return None
            if not navigation_service.select_circuit_id(uidocument, circuit.get("circuit_id")):
                raise ValueError("The selected circuit is unavailable.")
            return {"store": store, "message": "Circuit selected in model."}
        if operation == "export_definitions":
            message = _export_definitions(store, preferred_set_id=set_id)
            return {"store": store, "message": message} if message else None
        if operation == "import_definitions":
            imported = _import_definitions(document, store, self.gateway.logger)
            if imported is None:
                return None
            updated, message = imported
            return self._save_result(document, updated, message, "Parameter Monitor - Import Sets")
        if operation == "export_report":
            message = _export_report(store, set_id)
            return {"store": store, "message": message} if message else None
        if operation == "sync_element_linker":
            synced = element_linker_sync_service.run_sync(
                document, uidocument, store, logger=self.gateway.logger
            )
            if synced is None:
                return None
            updated, sync_set_id, message = synced
            result = self._save_result(
                document, updated, message, "Parameter Monitor - Element Linker Sync"
            )
            result["sync_set_id"] = sync_set_id
            return result
        if operation == "move_with_parent":
            if not persistent_ids:
                raise ValueError("Select one or more tracked child elements first.")
            updated, message = parent_move_service.move_children_with_parent(
                document, store, set_id, persistent_ids, logger=self.gateway.logger
            )
            return self._save_result(
                document, updated, message, "Parameter Monitor - Move with Parent"
            )
        raise ValueError("Unknown Parameter Monitor operation: {}".format(operation))

    def GetName(self):  # noqa: N802
        return "CED Parameter Monitor External Event"
