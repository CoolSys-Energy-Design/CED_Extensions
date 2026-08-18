# -*- coding: utf-8 -*-
"""ExternalEvent gateway and Revit-side operations for Tag by Example."""

import math

from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType
from System import EventHandler
from System.Collections.Generic import List
from pyrevit import revit, DB, UI, script

from Snippets import revit_helpers
from Snippets._rotateutils import apply_orientation_rules
from Snippets.tag_geometry import copy_geometry_for_target
from Snippets.tag_geometry import geometry_from_world_points
from Snippets.tag_host_transform import get_host_placement_frame
from tag_api_compat import create_independent_tag
from tag_api_compat import get_leader_snapshot
from tag_api_compat import get_rotation_angle
from tag_api_compat import get_tag_orientation
from tag_api_compat import get_tag_type_id
from tag_api_compat import get_tagged_local_ids
from tag_api_compat import get_tagged_references
from tag_api_compat import id_value
from tag_api_compat import set_leader_elbow
from tag_api_compat import set_leader_end
from tag_api_compat import set_leader_end_condition
from tag_api_compat import valid_id
from tag_host_adapters import ExampleTagSelectionFilter
from tag_host_adapters import TargetSelectionFilter
from tag_host_adapters import category_id
from tag_host_adapters import family_name
from tag_host_adapters import get_tag_owner_view
from tag_host_adapters import host_description
from tag_host_adapters import is_multi_category_tag
from tag_host_adapters import is_nested_instance
from tag_host_adapters import is_phase_one_host
from tag_host_adapters import is_supported_tag_view
from tag_host_adapters import is_visible_candidate
from tag_host_adapters import matches_target_mode
from tag_host_adapters import supported_view_description
from tag_host_adapters import target_reference
from tag_host_adapters import type_name
from tag_host_adapters import validate_example_tag

TARGET_MODE_LABELS = {
    "type": "All visible elements of the same type",
    "family": "All visible elements of the same family",
    "category": "All visible elements of the same category",
    "manual": "User selection",
}


EXISTING_BEHAVIOR_LABELS = {
    "replace_all": "Replace all tags",
    "replace_matching": "Replace matching reference tag types only",
    "skip_matching": "Skip matching reference tag types",
}


class TagByExampleUserError(Exception):
    """An expected user-action problem that belongs in the UI, not the log."""


def _element_id(value):
    return revit_helpers.elementid_from_value(value)


def _text_name(element, fallback):
    try:
        return str(DB.Element.Name.__get__(element))
    except Exception:
        return fallback


def _orientation_or_default(orientation):
    if orientation is not None:
        return orientation
    return DB.TagOrientation.Horizontal


def _document_key(document):
    try:
        return "{}|{}".format(document.PathName, document.Title)
    except Exception:
        return "{}".format(document)


def _safe_view_name(view):
    try:
        return str(view.Name)
    except Exception:
        return "<Unnamed view>"


def _unique_integer_values(values):
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


def _resolve_tag_data(document, owner_view, tag_id):
    tag = document.GetElement(_element_id(tag_id))
    if tag is None:
        raise ValueError("A reference tag no longer exists.")
    reference, host, frame = validate_example_tag(document, owner_view, tag)
    tag_owner_view = get_tag_owner_view(document, tag)
    head_position = tag.TagHeadPosition
    leader = get_leader_snapshot(tag, reference)
    geometry = geometry_from_world_points(
        head_position,
        leader.get("elbow"),
        leader.get("end"),
        frame,
    )
    geometry.has_leader = leader.get("has_leader", False)
    geometry.leader_end_condition = leader.get("end_condition")
    geometry.orientation = get_tag_orientation(tag)
    geometry.rotation_angle = get_rotation_angle(tag)
    geometry.source_rotation = frame.rotation
    geometry.tag_type_id = get_tag_type_id(tag)
    if not valid_id(geometry.tag_type_id):
        raise ValueError("A reference tag has no resolvable tag type.")

    tag_type = document.GetElement(geometry.tag_type_id)
    tag_type_name = _text_name(tag_type, "")
    if not tag_type_name:
        tag_type_name = revit_helpers.get_family_symbol_name(
            tag_type,
            doc=document,
            fallback="<Unnamed tag type>",
        )
    description = host_description(host, document)
    return {
        "tag": tag,
        "tag_id": id_value(tag.Id),
        "reference": reference,
        "host": host,
        "host_id": id_value(host.Id),
        "host_frame": frame,
        "geometry": geometry,
        "tag_type_id": geometry.tag_type_id,
        "tag_type_name": tag_type_name,
        "owner_view_id": id_value(tag_owner_view.Id),
        "owner_view_name": _safe_view_name(tag_owner_view),
        "orientation": geometry.orientation,
        "orientation_text": str(geometry.orientation),
        "leader": leader,
        "is_multi_category": is_multi_category_tag(document, tag),
        "host_description": description,
        "document_key": _document_key(document),
    }


def _resolve_reference_set(document, gateway):
    owner_view = document.GetElement(_element_id(gateway.owner_view_id))
    if owner_view is None:
        raise ValueError("The reference tag owner view is no longer valid.")
    reference_ids = _unique_integer_values(gateway.example_tag_ids)
    if not reference_ids:
        raise ValueError("Pick at least one reference tag first.")

    examples = []
    host_value = None
    for reference_id in reference_ids:
        example_data = _resolve_tag_data(document, owner_view, reference_id)
        if host_value is None:
            host_value = example_data["host_id"]
        elif host_value != example_data["host_id"]:
            raise ValueError("All reference tags must be hosted by the same element.")
        examples.append(example_data)
    return examples


def _candidate_ids(document, view, example_data, mode, include_nested, manual_ids):
    candidates = []
    skipped = []
    seen_ids = set()
    example_host = example_data[0]["host"]
    allow_any_family = any([
        bool(example_item.get("is_multi_category", False))
        for example_item in example_data
    ])

    if mode == "manual":
        source_elements = []
        for value in list(manual_ids or []):
            source_elements.append(document.GetElement(_element_id(value)))
    else:
        collector = DB.FilteredElementCollector(document, view.Id)
        source_elements = list(collector.WhereElementIsNotElementType())

    for element in source_elements:
        if element is None:
            skipped.append({"id": 0, "reason": "Element no longer exists."})
            continue
        element_value = id_value(element.Id)
        if element_value in seen_ids:
            continue
        seen_ids.add(element_value)
        allowed, reason = is_visible_candidate(
            document, view, element, example_host, include_nested
        )
        if not allowed:
            if mode == "manual" or "Nested" in reason:
                skipped.append({"id": element_value, "element": element, "reason": reason})
            continue
        if not is_phase_one_host(element):
            if mode == "manual":
                skipped.append({
                    "id": element_value,
                    "element": element,
                    "reason": "Unsupported host class.",
                })
            continue
        if mode != "manual" and not matches_target_mode(
                element, example_host, mode, document):
            continue
        if (mode == "manual" and not allow_any_family
                and category_id(element) != category_id(example_host)):
            skipped.append({
                "id": element_value,
                "element": element,
                "reason": "Target category is incompatible with the reference host.",
            })
            continue
        candidates.append(element)
    return candidates, skipped


def _build_proposals(document, view, examples, targets, options):
    proposals = []
    skipped = []
    preserve_rotation = bool(options.get("preserve_rotation", True))

    for target in targets:
        target_value = id_value(target.Id)
        try:
            reference = target_reference(target)
            target_frame = get_host_placement_frame(target, view)
            reference_items = []
            for example_data in examples:
                points = copy_geometry_for_target(
                    example_data["geometry"],
                    example_data["host_frame"],
                    target_frame,
                    preserve_offset=True,
                    adjust_rotation=preserve_rotation,
                )
                reference_items.append({
                    "example": example_data,
                    "reference": reference,
                    "points": points,
                    "type_value": id_value(example_data["tag_type_id"]),
                })
            proposals.append({
                "target": target,
                "target_id": target_value,
                "frame": target_frame,
                "reference_items": reference_items,
                "nested": is_nested_instance(target),
                "family": family_name(target, document),
                "type": type_name(target, document),
            })
        except Exception as error:
            skipped.append({
                "id": target_value,
                "element": target,
                "reason": str(error),
            })
    return proposals, skipped


def _collect_tag_index(document, view):
    """Index supported current-view tags once by every confidently local host."""
    index = {}
    collector = DB.FilteredElementCollector(document, view.Id).OfClass(DB.IndependentTag)
    for tag in collector:
        local_ids = get_tagged_local_ids(tag)
        references = get_tagged_references(tag)
        multi_reference = len(local_ids) != 1 or len(references) > 1
        tag_record = {
            "tag": tag,
            "tag_id": id_value(tag.Id),
            "multi_reference": multi_reference,
            "type_id": id_value(get_tag_type_id(tag)),
        }
        for host_id in local_ids:
            host_value = id_value(host_id)
            index.setdefault(host_value, []).append(tag_record)
    return index


def _add_failure(failures, failure_keys, target_id, element, reason):
    failure_key = "{}|{}".format(target_id, reason)
    if failure_key in failure_keys:
        return
    failure_keys.add(failure_key)
    failures.append({"id": target_id, "element": element, "reason": reason})


def _prepare_existing_tag_actions(document, tag_index, proposals, examples, behavior):
    """Determine deletion and skip decisions before any tag is created."""
    delete_records = {}
    blocked_pairs = set()
    blocked_targets = set()
    skipped_pairs = set()
    failures = []
    failure_keys = set()
    for proposal in proposals:
        target_id = proposal["target_id"]
        records = tag_index.get(target_id, [])
        if behavior == "replace_all":
            unsafe_records = [record for record in records if record["multi_reference"]]
            if unsafe_records:
                blocked_targets.add(target_id)
                for record in unsafe_records:
                    _add_failure(
                        failures,
                        failure_keys,
                        target_id,
                        proposal["target"],
                        "Existing tag references several hosts and was not deleted; "
                        "the target was not retagged.",
                    )
                continue
            for record in records:
                delete_records[record["tag_id"]] = {
                    "record": record,
                    "target_id": target_id,
                    "element": proposal["target"],
                }
            continue

        for item in proposal["reference_items"]:
            type_value = item["type_value"]
            matching_records = [
                record for record in records
                if record["type_id"] == type_value
            ]
            pair_key = (target_id, type_value)
            if not matching_records:
                continue
            if behavior == "skip_matching":
                skipped_pairs.add(pair_key)
                continue

            unsafe_records = [record for record in matching_records
                              if record["multi_reference"]]
            if unsafe_records:
                blocked_pairs.add(pair_key)
                _add_failure(
                    failures,
                    failure_keys,
                    target_id,
                    proposal["target"],
                    "Existing matching tag references several hosts and was not "
                    "deleted; this reference type was not recreated.",
                )
                continue
            for record in matching_records:
                delete_records[record["tag_id"]] = {
                    "record": record,
                    "target_id": target_id,
                    "element": proposal["target"],
                    "type_value": type_value,
                }

    deleted = 0
    if delete_records:
        transaction = DB.Transaction(document, "Delete existing tags")
        transaction.Start()
        try:
            for delete_id in delete_records:
                deletion = delete_records[delete_id]
                record = deletion["record"]
                subtransaction = DB.SubTransaction(document)
                subtransaction.Start()
                try:
                    document.Delete(record["tag"].Id)
                    subtransaction.Commit()
                    deleted += 1
                except Exception as error:
                    subtransaction.RollBack()
                    target_id = deletion["target_id"]
                    if behavior == "replace_all":
                        blocked_targets.add(target_id)
                        reason = "Existing tag deletion failed; the target was not retagged: {}".format(error)
                    else:
                        pair_key = (target_id, deletion.get("type_value", record["type_id"]))
                        blocked_pairs.add(pair_key)
                        reason = "Existing matching tag deletion failed; this reference type was not recreated: {}".format(error)
                    _add_failure(
                        failures,
                        failure_keys,
                        target_id,
                        deletion["element"],
                        reason,
                    )
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise

    return {
        "deleted": deleted,
        "failures": failures,
        "blocked_pairs": blocked_pairs,
        "blocked_targets": blocked_targets,
        "skipped_pairs": skipped_pairs,
    }


def _apply_tag_properties(document, created_tag, item, options):
    warnings = []
    example_data = item["example"]
    geometry = example_data["geometry"]
    points = item["points"]
    reference = item["reference"]

    target_angle = points.get("rotation_angle")
    if target_angle is None:
        target_angle = 0.0
    try:
        if options.get("use_model_orientation", True):
            created_tag.TagOrientation = DB.TagOrientation.AnyModelDirection
            created_tag.RotationAngle = target_angle
        else:
            apply_orientation_rules(created_tag, target_angle, tolerance=math.radians(3))
    except Exception as error:
        warnings.append("Tag orientation could not be applied: {}".format(error))

    try:
        created_tag.TagHeadPosition = points.get("head")
    except Exception as error:
        raise RuntimeError("Tag head position could not be applied: {}".format(error))

    if options.get("copy_leader", True) and geometry.has_leader:
        if not set_leader_end_condition(created_tag, geometry.leader_end_condition):
            warnings.append("Leader end condition is unsupported for this tag.")
        if geometry.elbow_local is not None:
            if not set_leader_elbow(created_tag, reference, points.get("elbow")):
                warnings.append("Leader elbow is unsupported for this tag.")
        if geometry.end_local is not None:
            if not set_leader_end(created_tag, reference, points.get("end")):
                warnings.append("Leader end is unsupported for this tag.")
    return warnings


def _create_tags(document, view, proposals, options, actions):
    created = 0
    duplicates = 0
    existing_skips = 0
    failures = []
    warnings = []
    attempted_pairs = set()
    transaction = DB.Transaction(document, "Create tags")
    transaction.Start()
    try:
        for proposal in proposals:
            target_id = proposal["target_id"]
            for item in proposal["reference_items"]:
                pair_key = (target_id, item["type_value"])
                if pair_key in actions["skipped_pairs"]:
                    existing_skips += 1
                    continue
                if (target_id in actions["blocked_targets"]
                        or pair_key in actions["blocked_pairs"]):
                    continue
                if pair_key in attempted_pairs:
                    duplicates += 1
                    continue
                attempted_pairs.add(pair_key)

                subtransaction = DB.SubTransaction(document)
                subtransaction.Start()
                try:
                    example_data = item["example"]
                    orientation = _orientation_or_default(
                        example_data["geometry"].orientation
                    )
                    has_leader = bool(
                        example_data["geometry"].has_leader
                        and options.get("copy_leader", True)
                    )
                    created_tag = create_independent_tag(
                        document,
                        view.Id,
                        example_data["tag_type_id"],
                        item["reference"],
                        has_leader,
                        orientation,
                        item["points"].get("head"),
                    )
                    document.Regenerate()
                    target_warnings = _apply_tag_properties(
                        document, created_tag, item, options
                    )
                    warnings.extend([(target_id, warning)
                                     for warning in target_warnings])
                    subtransaction.Commit()
                    created += 1
                except Exception as error:
                    subtransaction.RollBack()
                    failures.append({
                        "id": target_id,
                        "element": proposal["target"],
                        "reason": "Tag creation failed: {}".format(error),
                    })
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    return created, duplicates, existing_skips, failures, warnings


def _selection_id_list(values):
    result = List[DB.ElementId]()
    for value in _unique_integer_values(values):
        result.Add(_element_id(value))
    return result


def _selection_reference_list(document, values):
    result = List[DB.Reference]()
    for value in _unique_integer_values(values):
        element = document.GetElement(_element_id(value))
        if element is None:
            continue
        try:
            result.Add(DB.Reference(element))
        except Exception:
            continue
    return result


class TagByExampleExternalEventGateway(object):
    def __init__(self, window, document, owner_view, initial_tag_ids=None,
                 ui_application=None, show_output_report=False):
        self.window = window
        self.document_key = _document_key(document)
        self.owner_view_id = id_value(owner_view.Id)
        self.active_view_id = id_value(owner_view.Id)
        self.example_tag_ids = _unique_integer_values(initial_tag_ids or [])
        self.example_tag_id = self.example_tag_ids[0] if self.example_tag_ids else None
        self.manual_target_ids = []
        self.pending = None
        self.ui_application = ui_application
        self.show_output_report = bool(show_output_report)
        self.lifecycle_handlers = {}
        self.document_closing_handler = None
        self.lifecycle_attached = False
        self.handler = _TagByExampleHandler(self)
        self.event = UI.ExternalEvent.Create(self.handler)

    def attach_lifecycle(self):
        if self.lifecycle_attached or self.ui_application is None:
            return self.lifecycle_attached
        try:
            self.lifecycle_handlers["view-activated"] = revit.events.add_handler(
                "view-activated", self._view_activated
            )
            self.lifecycle_handlers["doc-changed"] = revit.events.add_handler(
                "doc-changed", self._document_changed
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
                "Could not subscribe to Revit lifecycle events: {}".format(error)
            )
            return False

    def detach_lifecycle(self):
        if not self.lifecycle_attached and not self.lifecycle_handlers:
            return
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

    def _view_activated(self, sender, args):
        del sender
        try:
            active_view = args.CurrentActiveView
            active_uidoc = self.ui_application.ActiveUIDocument
            if active_uidoc is None:
                return
            active_document = active_uidoc.Document
            active_document_key = _document_key(active_document)
            if active_document_key != self.document_key:
                self.clear_examples()
                self.detach_lifecycle()
                self.window.receive_result("lifecycle", "document_changed", {
                    "view_supported": False,
                    "view_name": _safe_view_name(active_view),
                }, None)
                return
            active_view_id = id_value(active_view.Id)
            if active_view_id == self.active_view_id:
                return
            self.active_view_id = active_view_id
            self.manual_target_ids = []
            self.window.receive_result("lifecycle", "view_activated", {
                "view_supported": is_supported_tag_view(active_view),
                "view_name": _safe_view_name(active_view),
            }, None)
        except Exception as error:
            script.get_logger().warning(
                "Tag by Example view lifecycle handler failed: {}".format(error)
            )

    def _document_changed(self, sender, args):
        del sender
        try:
            changed_document = args.GetDocument()
            if _document_key(changed_document) != self.document_key:
                return
            reference_ids = set(self.example_tag_ids)
            deleted_ids = args.GetDeletedElementIds()
            for deleted_id in deleted_ids:
                if id_value(deleted_id) in reference_ids:
                    self.clear_examples()
                    self.window.receive_result(
                        "lifecycle",
                        "reference_deleted",
                        {"message": "A reference tag was deleted."},
                        None,
                    )
                    return
        except Exception as error:
            script.get_logger().warning(
                "Tag by Example document lifecycle handler failed: {}".format(error)
            )

    def _document_closing(self, sender, args):
        del sender
        try:
            closing_document = args.Document
            if _document_key(closing_document) != self.document_key:
                return
            self.clear_examples()
            self.detach_lifecycle()
            self.window.receive_result(
                "lifecycle",
                "document_closed",
                {"message": "The reference document was closed."},
                None,
            )
        except Exception as error:
            script.get_logger().warning(
                "Tag by Example document closing handler failed: {}".format(error)
            )

    def set_example_ids(self, tag_ids, owner_view_id):
        self.example_tag_ids = _unique_integer_values(tag_ids)
        self.example_tag_id = self.example_tag_ids[0] if self.example_tag_ids else None
        self.owner_view_id = id_value(owner_view_id)
        self.manual_target_ids = []

    def clear_examples(self):
        self.example_tag_ids = []
        self.example_tag_id = None
        self.manual_target_ids = []

    def raise_action(self, action_name, payload=None):
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
                "Could not queue Revit operation: {}".format(error)
            )
            return False

    def consume(self):
        pending = self.pending
        self.pending = None
        return pending


class _TagByExampleHandler(UI.IExternalEventHandler):
    def __init__(self, gateway):
        self.gateway = gateway

    def _context(self, application, require_supported=True):
        uidoc = application.ActiveUIDocument
        document = uidoc.Document if uidoc is not None else None
        if document is None:
            raise ValueError("No active Revit document is available.")
        if _document_key(document) != self.gateway.document_key:
            raise ValueError("The active document changed; the references are no longer valid.")
        view = uidoc.ActiveView
        view_changed = id_value(view.Id) != self.gateway.active_view_id
        if view_changed:
            self.gateway.active_view_id = id_value(view.Id)
            self.gateway.manual_target_ids = []
        if require_supported and not is_supported_tag_view(view):
            raise ValueError(
                "Tag by Example is available only in floor plan, reflected ceiling "
                "plan, or drafting views. Current view: {}.".format(
                    supported_view_description(view)
                )
            )
        return uidoc, document, view, view_changed

    def _examples(self, document):
        return _resolve_reference_set(document, self.gateway)

    def _result(self, status, action_name, result=None, error=None):
        if bool(getattr(self.gateway.window, "is_closed", False)):
            return
        try:
            self.gateway.window.receive_result(status, action_name, result, error)
        except Exception as callback_error:
            script.get_logger().exception(
                "Tag by Example UI callback failed: {}".format(callback_error)
            )

    def Execute(self, application):
        pending = self.gateway.consume()
        if not pending:
            return
        action_name = pending.get("action")
        payload = pending.get("payload") or {}

        if action_name == "sync":
            try:
                uidoc, document, view, view_changed = self._context(
                    application, require_supported=False
                )
                del uidoc
            except Exception as error:
                self._result("unavailable", action_name, None, error)
                return
            if not is_supported_tag_view(view):
                self._result("ok", action_name, {
                    "has_example": bool(self.gateway.example_tag_ids),
                    "view_supported": False,
                    "view_changed": view_changed,
                    "view_name": _safe_view_name(view),
                })
                return
            if not self.gateway.example_tag_ids:
                self._result("ok", action_name, {
                    "has_example": False,
                    "view_supported": True,
                    "view_changed": view_changed,
                })
                return
            try:
                examples = self._examples(document)
            except Exception as error:
                self.gateway.clear_examples()
                self._result("ok", action_name, {
                    "has_example": False,
                    "view_supported": True,
                    "view_changed": view_changed,
                    "example_invalid": True,
                    "message": str(error),
                })
                return
            self._result("ok", action_name, {
                "has_example": True,
                "view_supported": True,
                "view_changed": view_changed,
                "snapshot": self._snapshot(examples),
            })
            return

        try:
            allow_unsupported = action_name == "pick_example"
            uidoc, document, view, unused_changed = self._context(
                application, require_supported=not allow_unsupported
            )
            del unused_changed
            if action_name == "pick_example":
                if not is_supported_tag_view(view):
                    raise ValueError(
                        "Pick a reference in a floor plan, reflected ceiling plan, "
                        "or drafting view."
                    )
                selection_filter = ExampleTagSelectionFilter(document, view)
                try:
                    picked_references = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        selection_filter,
                        "Select one or more tags on the same host, then finish.",
                    )
                except OperationCanceledException:
                    self._result("cancelled", action_name, None, None)
                    return
                picked_ids = _unique_integer_values([
                    id_value(picked_reference.ElementId)
                    for picked_reference in picked_references
                ])
                if not picked_ids:
                    raise TagByExampleUserError(
                        "No reference tags were selected. Pick at least one reference tag to continue."
                    )
                old_ids = list(self.gateway.example_tag_ids)
                old_owner_view = self.gateway.owner_view_id
                selected_owner_ids = []
                for picked_id in picked_ids:
                    picked_tag = document.GetElement(_element_id(picked_id))
                    try:
                        selected_owner_ids.append(
                            id_value(get_tag_owner_view(document, picked_tag).Id)
                        )
                    except Exception:
                        selected_owner_ids = []
                        break
                selected_owner_ids = _unique_integer_values(selected_owner_ids)
                reference_owner_id = view.Id
                if len(selected_owner_ids) == 1:
                    reference_owner_id = _element_id(selected_owner_ids[0])
                self.gateway.set_example_ids(picked_ids, reference_owner_id)
                try:
                    examples = self._examples(document)
                except Exception:
                    self.gateway.example_tag_ids = old_ids
                    self.gateway.example_tag_id = old_ids[0] if old_ids else None
                    self.gateway.owner_view_id = old_owner_view
                    raise
                self._result("ok", action_name, {
                    "snapshot": self._snapshot(examples),
                    "target_count": 0,
                })
                return

            examples = self._examples(document)

            if action_name == "pick_targets":
                selection_mode = str(payload.get("selection_mode") or "new")
                if selection_mode == "new":
                    self.gateway.manual_target_ids = []
                include_nested = bool(payload.get("options", {}).get(
                    "include_nested", False
                ))
                target_filter = TargetSelectionFilter(
                    document,
                    view,
                    examples[0]["host"],
                    include_nested,
                    allow_any_family=any([
                        bool(example_item.get("is_multi_category", False))
                        for example_item in examples
                    ]),
                )
                preselected = _selection_reference_list(
                    document,
                    self.gateway.manual_target_ids if selection_mode == "edit" else [],
                )
                try:
                    if selection_mode == "edit" and preselected.Count > 0:
                        try:
                            picked_references = uidoc.Selection.PickObjects(
                                ObjectType.Element,
                                target_filter,
                                "Edit targets: add or remove compatible elements, then finish.",
                                preselected,
                            )
                        except TypeError:
                            picked_references = uidoc.Selection.PickObjects(
                                ObjectType.Element,
                                target_filter,
                                "Edit targets: select the final compatible target set, then finish.",
                            )
                    else:
                        picked_references = uidoc.Selection.PickObjects(
                            ObjectType.Element,
                            target_filter,
                            "Select compatible targets, then finish.",
                        )
                except OperationCanceledException:
                    self._result("cancelled", action_name, {
                        "count": len(self.gateway.manual_target_ids),
                        "target_ids": list(self.gateway.manual_target_ids),
                    }, None)
                    return
                picked_ids = _unique_integer_values([
                    id_value(picked_reference.ElementId)
                    for picked_reference in picked_references
                ])
                self.gateway.manual_target_ids = picked_ids
                self._result("ok", action_name, {
                    "count": len(picked_ids),
                    "target_ids": list(picked_ids),
                    "selection_mode": selection_mode,
                })
                return

            if action_name == "preview_selection":
                uidoc.Selection.SetElementIds(
                    _selection_id_list(self.gateway.manual_target_ids)
                )
                self._result("ok", action_name, {
                    "count": len(self.gateway.manual_target_ids),
                })
                return

            mode = str(payload.get("mode") or "type")
            options = dict(payload.get("options") or {})
            manual_ids = payload.get("target_ids")
            if manual_ids is None:
                manual_ids = self.gateway.manual_target_ids
            targets, skipped = _candidate_ids(
                document,
                view,
                examples,
                mode,
                bool(options.get("include_nested", False)),
                manual_ids,
            )
            if action_name == "refresh_targets":
                self._result("ok", action_name, {
                    "count": len(targets),
                    "skipped_count": len(skipped),
                })
                return

            proposals, proposal_skips = _build_proposals(
                document, view, examples, targets, options
            )
            skipped.extend(proposal_skips)
            if action_name == "create":
                tag_index = _collect_tag_index(document, view)
                behavior = str(options.get("existing_behavior") or "skip_matching")
                if behavior not in EXISTING_BEHAVIOR_LABELS:
                    behavior = "skip_matching"
                group = DB.TransactionGroup(document, "Tag by Example")
                group.Start()
                try:
                    actions = _prepare_existing_tag_actions(
                        document, tag_index, proposals, examples, behavior
                    )
                    created, duplicates, existing_skips, creation_failures, warnings = _create_tags(
                        document, view, proposals, options, actions
                    )
                    failures = list(actions["failures"])
                    failures.extend(creation_failures)
                    group.Assimilate()
                except Exception:
                    if group.GetStatus() == DB.TransactionStatus.Started:
                        group.RollBack()
                    raise
                result = {
                    "snapshot": self._snapshot(examples),
                    "candidate_count": len(targets),
                    "created": created,
                    "skipped": skipped,
                    "deleted": actions["deleted"],
                    "duplicates": duplicates,
                    "existing_skips": existing_skips,
                    "failures": failures,
                    "warnings": warnings,
                    "mode": TARGET_MODE_LABELS.get(mode, mode),
                    "existing_behavior": EXISTING_BEHAVIOR_LABELS[behavior],
                }
                if self.gateway.show_output_report:
                    try:
                        self._report(document, result)
                    except Exception as report_error:
                        script.get_logger().warning(
                            "Tag by Example completed, but its output report could not be "
                            "displayed: {}".format(report_error)
                        )
                self._result("ok", action_name, result)
                return

            raise ValueError("Unknown Tag by Example operation: {}".format(action_name))
        except Exception as error:
            if isinstance(error, TagByExampleUserError):
                self._result("user_error", action_name, None, error)
            else:
                script.get_logger().exception(
                    "Tag by Example operation failed: {}".format(error)
                )
                self._result("error", action_name, None, error)

    def _snapshot(self, examples):
        first_example = examples[0]
        description = first_example["host_description"]
        type_names = []
        for example_data in examples:
            if example_data["tag_type_name"] not in type_names:
                type_names.append(example_data["tag_type_name"])
        leader_values = []
        for example_data in examples:
            leader_text = "leader" if example_data["geometry"].has_leader else "no leader"
            leader_values.append(leader_text)
        return {
            "tag_id": first_example["tag_id"],
            "reference_count": len(examples),
            "tag_type": ", ".join(type_names),
            "host_category": description["category"],
            "host_family": description["family"],
            "host_type": description["type"],
            "owner_view": first_example["owner_view_name"],
            "has_leader": any([example_data["geometry"].has_leader
                                for example_data in examples]),
            "leader_summary": ", ".join(leader_values),
            "orientation": ", ".join([
                example_data["orientation_text"] for example_data in examples
            ]),
        }

    def _report(self, document, result):
        if not self.gateway.show_output_report:
            return False
        output = script.get_output()
        try:
            if output.window is None or bool(output.is_closed_by_user):
                return False
        except Exception:
            return False
        output.print_md("## Tag by Example report")
        output.print_md("- Reference tag types: **{}**".format(
            result["snapshot"]["tag_type"]
        ))
        output.print_md("- Reference tags: **{}**".format(
            result["snapshot"]["reference_count"]
        ))
        output.print_md("- Target mode: **{}**".format(result["mode"]))
        output.print_md("- Existing-tag behavior: **{}**".format(
            result["existing_behavior"]
        ))
        output.print_md("- Candidate targets: **{}**".format(result["candidate_count"]))
        output.print_md("- Successfully tagged: **{}**".format(result["created"]))
        output.print_md("- Existing tags deleted: **{}**".format(result["deleted"]))
        output.print_md("- Matching tags skipped: **{}**".format(
            result["existing_skips"]
        ))
        output.print_md("- Duplicate reference types avoided: **{}**".format(
            result["duplicates"]
        ))
        output.print_md("- Skipped targets: **{}**".format(len(result["skipped"])))
        output.print_md("- Failures: **{}**".format(len(result["failures"])))
        for item in result["skipped"]:
            element = item.get("element")
            if element is not None:
                label = output.linkify(element.Id)
                family_text = family_name(element, document)
                type_text = type_name(element, document)
            else:
                label = str(item.get("id", "?"))
                family_text = "<unknown>"
                type_text = "<unknown>"
            output.print_md("- {} — {} / {}: {}".format(
                label, family_text, type_text, item.get("reason", "Unknown reason")
            ))
        for item in result["failures"]:
            element = item.get("element")
            if element is not None:
                label = output.linkify(element.Id)
                family_text = family_name(element, document)
                type_text = type_name(element, document)
            else:
                label = str(item.get("id", "?"))
                family_text = "<unknown>"
                type_text = "<unknown>"
            output.print_md("- {} — {} / {}: {}".format(
                label, family_text, type_text, item.get("reason", "Unknown failure")
            ))
        for target_id, warning in result.get("warnings", []):
            output.print_md("- {} — warning: {}".format(
                output.linkify(_element_id(target_id)), warning
            ))
        try:
            output.show()
        except Exception as error:
            script.get_logger().debug(
                "Tag by Example output window was unavailable after completion: {}"
                .format(error)
            )
            return False
        return True

    def GetName(self):
        return "CED Tag by Example External Event"
