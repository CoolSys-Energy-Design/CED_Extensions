# -*- coding: utf-8 -*-
"""Phase 1 manual baseline, scan, resolve, and monitor-data operations."""

from __future__ import print_function

import copy
import time

import comparison_engine
import location_service
import models
import parameter_service
import relationship_service
import source_service
import text_service
from Snippets import revit_helpers


def _text(value, fallback=""):
    return text_service.to_text(value, fallback=fallback)


def _revit_parameter_text(element, name, source_document):
    """Read a Revit parameter as Unicode without Snippets' ``str()`` helper."""
    parameter = revit_helpers.get_parameter(
        element,
        name,
        include_type=False,
        # These are exact built-in/company parameter names. Avoid enumerating
        # every parameter on every element when either label is absent.
        case_insensitive=False,
        doc=source_document,
    )
    value = revit_helpers.get_parameter_value(parameter, default=None)
    return _text(value, "")


def _cached_type_element(element, source_document, type_cache):
    type_cache = type_cache if type_cache is not None else {}
    type_id = None
    try:
        type_id = element.GetTypeId()
    except Exception:
        pass
    type_key = revit_helpers.get_elementid_value(type_id)
    if type_key in type_cache:
        return type_cache[type_key]
    type_element = revit_helpers.get_type_element(element, doc=source_document)
    type_cache[type_key] = type_element
    return type_element


def _element_metadata(
    element,
    source_document,
    type_cache=None,
    include_workset=False,
    workset_cache=None,
):
    element_id = revit_helpers.get_elementid_value(getattr(element, "Id", None))
    mark = _revit_parameter_text(element, "Mark", source_document)
    if not mark:
        mark = _revit_parameter_text(element, "Equipment ID", source_document)
    name = _text(getattr(element, "Name", None), "")
    type_element = _cached_type_element(element, source_document, type_cache)
    type_name = _text(getattr(type_element, "Name", None), name)
    family_name = ""
    try:
        symbol = element.Symbol
        family_name = _text(symbol.Family.Name, "")
    except Exception:
        try:
            family_name = _text(type_element.FamilyName, "")
        except Exception:
            family_name = ""
    family_type = " : ".join([item for item in (family_name, type_name) if item])
    if not family_type:
        family_type = type_name or name or "-"
    level_name = ""
    try:
        level = source_document.GetElement(element.LevelId)
        level_name = _text(getattr(level, "Name", None), "")
    except Exception:
        level_name = ""
    category_name = ""
    try:
        category_name = _text(element.Category.Name, "")
    except Exception:
        pass
    metadata = {
        "element_id": element_id,
        "unique_id": _text(getattr(element, "UniqueId", None), ""),
        "friendly_name": _text(mark, "") or name or "Element {}".format(element_id),
        "mark": _text(mark, ""),
        "name": name,
        "family_type": family_type,
        "family_name": family_name or "-",
        "type_name": type_name or "-",
        "level": level_name,
        "category": category_name,
    }
    # Persist the source workset for link availability checks.  This field is
    # informational for host sets and essential for accurately protecting a
    # linked set from false removals when only some link worksets are closed.
    workset = {}
    if include_workset:
        workset_cache = workset_cache if workset_cache is not None else {}
        workset_id = revit_helpers.get_elementid_value(
            getattr(element, "WorksetId", None)
        )
        if workset_id not in workset_cache:
            workset_cache[workset_id] = source_service.element_workset_details(
                source_document, element
            )
        workset = workset_cache.get(workset_id) or {}
    if workset:
        metadata["workset_id"] = workset.get("id")
        metadata["workset_name"] = workset.get("name")
    return metadata


def _snapshot_element(
    host_document,
    source_document,
    source_descriptor,
    element,
    descriptors,
    type_cache,
    need_location,
    relationship=None,
    persistent_id_override=None,
    location_transform=None,
    workset_cache=None,
):
    persistent_id = persistent_id_override or source_service.persistent_id(
        source_descriptor, element
    )
    relationship_context = None
    if relationship:
        _device, relationship_context = relationship_service.resolve_relationship(
            host_document, relationship
        )
    return {
        "persistent_id": persistent_id,
        "source_element_unique_id": _text(getattr(element, "UniqueId", None), ""),
        "metadata": _element_metadata(
            element,
            source_document,
            type_cache=type_cache,
            include_workset=(source_descriptor or {}).get("source_type")
            == models.SOURCE_LINK,
            workset_cache=workset_cache,
        ),
        "properties": parameter_service.read_properties(
            element, source_document, descriptors, type_cache=type_cache
        ),
        "location": location_service.read_location(
            element, transform=location_transform
        ) if need_location else None,
        "relationship": copy.deepcopy(relationship),
        "relationship_context": relationship_context,
    }


def _collect_current_map(
    host_document,
    tracking_set,
    resolved_source,
    include_untracked=False,
    force_location=False,
):
    source_document = resolved_source.get("source_document")
    pairs = source_service.collect_set_member_pairs(source_document, tracking_set)
    descriptors = list(tracking_set.get("tracked_properties") or [])
    existing = tracking_set.get("elements") or {}
    untracked = set(tracking_set.get("untracked_ids") or [])
    default_location = bool(
        (tracking_set.get("location_defaults") or {}).get("track_new_elements", False)
    )
    current_map = {}
    type_caches = {}
    workset_caches = {}
    for persistent_id, element, element_document, transform in pairs:
        if persistent_id in untracked and not include_untracked:
            continue
        previous = existing.get(persistent_id) or {}
        need_location = bool(
            force_location or previous.get("track_location", default_location)
        )
        document_key = id(element_document)
        type_cache = type_caches.setdefault(document_key, {})
        workset_cache = workset_caches.setdefault(document_key, {})
        snapshot = _snapshot_element(
            host_document,
            element_document,
            tracking_set.get("source") or {},
            element,
            descriptors,
            type_cache,
            need_location,
            relationship=previous.get("relationship"),
            persistent_id_override=persistent_id,
            location_transform=transform,
            workset_cache=workset_cache,
        )
        current_map[persistent_id] = snapshot
    return current_map


def _replace_set(store, updated_set):
    result = copy.deepcopy(store)
    target = _text(updated_set.get("set_id") or "")
    replaced = False
    for index, tracking_set in enumerate(list(result.get("tracking_sets") or [])):
        if _text(tracking_set.get("set_id") or "") == target:
            result["tracking_sets"][index] = updated_set
            replaced = True
            break
    if not replaced:
        result.setdefault("tracking_sets", []).append(updated_set)
    result["updated_at"] = models.utc_now_text()
    return result


def _require_set(store, set_id):
    tracking_set = models.find_set(store, set_id)
    if tracking_set is None:
        raise ValueError("Tracking Set no longer exists.")
    return tracking_set


def create_tracking_set(host_document, store, name, source, category, descriptors, location_defaults=None, logger=None):
    tracking_set = models.new_tracking_set(
        name,
        source,
        category,
        descriptors,
        location_defaults=location_defaults,
    )
    resolved = source_service.evaluate_source_for_scan(
        host_document, tracking_set, require_complete=True
    )
    if not resolved.get("available"):
        raise ValueError(resolved.get("message") or "Source model is unavailable.")
    started = time.time()
    current_map = _collect_current_map(host_document, tracking_set, resolved)
    track_location = bool(
        (tracking_set.get("location_defaults") or {}).get("track_new_elements", False)
    )
    for persistent_id, snapshot in list(current_map.items()):
        tracking_set["elements"][persistent_id] = models.new_tracked_element(
            snapshot,
            baseline=True,
            track_location=track_location,
        )
    now = models.utc_now_text()
    tracking_set["accepted_source_state"] = copy.deepcopy(resolved.get("source_state") or {})
    tracking_set["current_source_state"] = copy.deepcopy(resolved.get("source_state") or {})
    tracking_set["last_check"] = now
    tracking_set["status"] = models.SET_CLEAN
    tracking_set["status_message"] = "Baseline created for {} element(s).".format(len(current_map))
    tracking_set["updated_at"] = now
    if logger is not None:
        try:
            logger.info(
                "Parameter Monitor baseline set=%s elements=%s elapsed=%.3fs",
                tracking_set.get("set_id"),
                len(current_map),
                time.time() - started,
            )
        except Exception:
            pass
    return _replace_set(store, tracking_set), tracking_set


def _failed_set(tracking_set, message):
    result = copy.deepcopy(tracking_set)
    result["status"] = models.SET_CHECK_FAILED
    result["status_message"] = _text(message or "Check failed.")
    result["last_check"] = models.utc_now_text()
    result["updated_at"] = result["last_check"]
    return result


def scan_tracking_set(host_document, store, set_id, logger=None):
    tracking_set = _require_set(store, set_id)
    resolved = source_service.evaluate_source_for_scan(host_document, tracking_set)
    now = models.utc_now_text()
    if not resolved.get("available"):
        unavailable = copy.deepcopy(tracking_set)
        unavailable["status"] = (
            resolved.get("set_status") or models.SET_SOURCE_UNAVAILABLE
        )
        unavailable["status_message"] = resolved.get("message") or "Source model unavailable."
        unavailable["last_check"] = now
        unavailable["updated_at"] = now
        if logger is not None:
            try:
                logger.warning(
                    "Parameter Monitor scan skipped for set %s: %s",
                    set_id,
                    unavailable["status_message"],
                )
            except Exception:
                pass
        return _replace_set(store, unavailable), unavailable
    started = time.time()
    try:
        current_map = _collect_current_map(host_document, tracking_set, resolved)
        scanned = comparison_engine.apply_scan(
            tracking_set,
            current_map,
            checked_at=now,
            source_state=resolved.get("source_state") or {},
        )
        if logger is not None:
            try:
                logger.info(
                    "Parameter Monitor scan set=%s elements=%s elapsed=%.3fs",
                    set_id,
                    len(current_map),
                    time.time() - started,
                )
            except Exception:
                pass
        return _replace_set(store, scanned), scanned
    except Exception as ex:
        if logger is not None:
            try:
                logger.exception("Parameter Monitor scan failed for set %s", set_id)
            except Exception:
                pass
        failed = _failed_set(
            tracking_set, text_service.diagnostic_text(ex, u"Check failed.")
        )
        return _replace_set(store, failed), failed


def scan_all_active(host_document, store, logger=None):
    result = copy.deepcopy(store)
    scanned_ids = []
    for tracking_set in list(result.get("tracking_sets") or []):
        if not bool(tracking_set.get("active", True)):
            continue
        result, _updated = scan_tracking_set(
            host_document,
            result,
            tracking_set.get("set_id"),
            logger=logger,
        )
        scanned_ids.append(tracking_set.get("set_id"))
    return result, scanned_ids


def edit_tracking_set(
    host_document,
    store,
    set_id,
    name,
    descriptors,
    active=None,
    track_new_elements=None,
    logger=None,
):
    tracking_set = copy.deepcopy(_require_set(store, set_id))
    resolved = source_service.evaluate_source_for_scan(host_document, tracking_set)
    if not resolved.get("available"):
        raise ValueError(resolved.get("message") or "Source model is unavailable.")
    temporary = copy.deepcopy(tracking_set)
    temporary["tracked_properties"] = copy.deepcopy(list(descriptors or []))
    current_map = _collect_current_map(host_document, temporary, resolved)
    updated = comparison_engine.update_tracked_properties(
        tracking_set,
        descriptors,
        current_map=current_map,
    )
    updated["name"] = text_service.to_text(
        name or updated.get("name") or "Tracking Set"
    )
    if active is not None:
        updated["active"] = bool(active)
    if track_new_elements is not None:
        updated.setdefault("location_defaults", {})["track_new_elements"] = bool(
            track_new_elements
        )
    updated["status_message"] = "Tracking Set definition updated."
    return _replace_set(store, updated), updated


def delete_tracking_set(store, set_id):
    result = copy.deepcopy(store)
    target = _text(set_id or "")
    result["tracking_sets"] = [
        item for item in list(result.get("tracking_sets") or [])
        if _text(item.get("set_id") or "") != target
    ]
    result["updated_at"] = models.utc_now_text()
    return result


def toggle_set_active(store, set_id):
    tracking_set = copy.deepcopy(_require_set(store, set_id))
    tracking_set["active"] = not bool(tracking_set.get("active", True))
    tracking_set["updated_at"] = models.utc_now_text()
    return _replace_set(store, tracking_set), tracking_set


def resolve_property(store, set_id, persistent_id, property_key):
    tracking_set = _require_set(store, set_id)
    updated = comparison_engine.resolve_property(tracking_set, persistent_id, property_key)
    return _replace_set(store, updated), updated


def resolve_element(store, set_id, persistent_id):
    tracking_set = _require_set(store, set_id)
    updated = comparison_engine.resolve_element(tracking_set, persistent_id)
    return _replace_set(store, updated), updated


def resolve_set(store, set_id):
    tracking_set = _require_set(store, set_id)
    updated = comparison_engine.resolve_set(tracking_set)
    return _replace_set(store, updated), updated


def untrack_element(store, set_id, persistent_id):
    return untrack_elements(store, set_id, [persistent_id])


def untrack_elements(store, set_id, persistent_ids):
    tracking_set = _require_set(store, set_id)
    updated = comparison_engine.untrack_elements(tracking_set, persistent_ids)
    return _replace_set(store, updated), updated


def _current_snapshot_for_key(host_document, tracking_set, persistent_id, force_location=False):
    resolved = source_service.resolve_source(host_document, tracking_set.get("source") or {})
    if not resolved.get("available"):
        raise ValueError(resolved.get("message") or "Source model is unavailable.")
    source_document = resolved.get("source_document")
    descriptors = list(tracking_set.get("tracked_properties") or [])
    type_cache = {}
    for key, element, element_document, transform in source_service.collect_set_member_pairs(
        source_document, tracking_set
    ):
        if key != _text(persistent_id or ""):
            continue
        previous = (tracking_set.get("elements") or {}).get(key) or {}
        return _snapshot_element(
            host_document,
            element_document,
            tracking_set.get("source") or {},
            element,
            descriptors,
            type_cache,
            force_location or bool(previous.get("track_location", False)),
            relationship=previous.get("relationship"),
            persistent_id_override=key,
            location_transform=transform,
        )
    return None


def restore_element(host_document, store, set_id, persistent_id):
    tracking_set = _require_set(store, set_id)
    snapshot = _current_snapshot_for_key(host_document, tracking_set, persistent_id, force_location=True)
    if snapshot is None:
        raise ValueError(
            "The untracked element is not currently available. Tracking was not restored."
        )
    updated = comparison_engine.restore_element(tracking_set, persistent_id, snapshot)
    return _replace_set(store, updated), updated


def set_all_element_location_tracking(host_document, store, set_id, enabled):
    tracking_set = copy.deepcopy(_require_set(store, set_id))
    current_map = {}
    resolved = None
    if enabled:
        resolved = source_service.resolve_source(
            host_document, tracking_set.get("source") or {}
        )
        if not resolved.get("available"):
            raise ValueError(resolved.get("message") or "Source model is unavailable.")
        current_map = _collect_current_map(
            host_document,
            tracking_set,
            resolved,
            force_location=True,
        )
    else:
        resolved = source_service.resolve_source(
            host_document, tracking_set.get("source") or {}
        )
    tracking_set.setdefault("location_defaults", {})["track_new_elements"] = bool(enabled)
    for persistent_id, record in list((tracking_set.get("elements") or {}).items()):
        if record.get("state") == models.ELEMENT_REMOVED:
            continue
        record["track_location"] = bool(enabled)
        if enabled:
            snapshot = current_map.get(persistent_id) or {}
            current_location = copy.deepcopy(snapshot.get("location"))
            record["current_location"] = current_location
            record["accepted_location"] = copy.deepcopy(current_location)
        else:
            record["current_location"] = None
            record["accepted_location"] = None
        comparison_engine.recompute_record(record, tracking_set)
    if resolved is not None and resolved.get("available"):
        source_state = copy.deepcopy(resolved.get("source_state") or {})
        tracking_set["accepted_source_state"] = copy.deepcopy(source_state)
        tracking_set["current_source_state"] = source_state
        tracking_set["source_conditions"] = []
    elif not enabled:
        tracking_set["accepted_source_state"] = copy.deepcopy(
            tracking_set.get("current_source_state") or {}
        )
        tracking_set["source_conditions"] = []
    comparison_engine.refresh_set_status(tracking_set)
    tracking_set["updated_at"] = models.utc_now_text()
    return _replace_set(store, tracking_set), tracking_set


def remove_record(store, set_id, persistent_id):
    tracking_set = _require_set(store, set_id)
    updated = comparison_engine.remove_record(tracking_set, persistent_id)
    return _replace_set(store, updated), updated


def remove_all_removed(store, set_id):
    tracking_set = _require_set(store, set_id)
    updated = comparison_engine.remove_all_removed(tracking_set)
    return _replace_set(store, updated), updated


def set_element_location_tracking(host_document, store, set_id, persistent_id, enabled):
    return set_elements_location_tracking(
        host_document,
        store,
        set_id,
        [persistent_id],
        enabled,
    )


def _current_locations_for_keys(host_document, tracking_set, persistent_ids):
    targets = set([_text(item or "") for item in list(persistent_ids or []) if item])
    resolved = source_service.resolve_source(host_document, tracking_set.get("source") or {})
    if not resolved.get("available"):
        raise ValueError(resolved.get("message") or "Source model is unavailable.")
    source_document = resolved.get("source_document")
    locations = {}
    for key, element, _element_document, transform in source_service.collect_set_member_pairs(
        source_document, tracking_set
    ):
        if key not in targets:
            continue
        locations[key] = location_service.read_location(element, transform=transform)
        if len(locations) >= len(targets):
            break
    return locations, resolved


def set_elements_location_tracking(
    host_document,
    store,
    set_id,
    persistent_ids,
    enabled,
):
    tracking_set = _require_set(store, set_id)
    keys = []
    seen = set()
    for item in list(persistent_ids or []):
        key = _text(item or "")
        if key and key not in seen:
            keys.append(key)
            seen.add(key)
    eligible_keys = [
        key for key in keys
        if (tracking_set.get("elements") or {}).get(key) is not None
        and (tracking_set.get("elements") or {}).get(key).get("state") != models.ELEMENT_REMOVED
    ]
    if not eligible_keys:
        unchanged = copy.deepcopy(tracking_set)
        return _replace_set(store, unchanged), unchanged

    had_location_tracking = any([
        bool(record.get("track_location", False))
        for record in list((tracking_set.get("elements") or {}).values())
        if record.get("state") != models.ELEMENT_REMOVED
    ])
    current_locations = {}
    resolved = None
    if enabled:
        current_locations, resolved = _current_locations_for_keys(
            host_document,
            tracking_set,
            eligible_keys,
        )
        unavailable = [key for key in eligible_keys if key not in current_locations]
        if unavailable:
            raise ValueError(
                "{} selected element(s) are unavailable for location tracking.".format(
                    len(unavailable)
                )
            )

    updated = copy.deepcopy(tracking_set)
    for key in eligible_keys:
        record = updated["elements"][key]
        record["track_location"] = bool(enabled)
        if enabled:
            location = copy.deepcopy(current_locations.get(key))
            record["current_location"] = location
            record["accepted_location"] = copy.deepcopy(location)
        else:
            record["current_location"] = None
            record["accepted_location"] = None
        comparison_engine.recompute_record(record, updated)

    has_location_tracking = any([
        bool(record.get("track_location", False))
        for record in list((updated.get("elements") or {}).values())
        if record.get("state") != models.ELEMENT_REMOVED
    ])
    if had_location_tracking and not has_location_tracking:
        updated["accepted_source_state"] = copy.deepcopy(
            updated.get("current_source_state") or {}
        )
        updated["source_conditions"] = []
        comparison_engine.refresh_set_status(updated)
    elif had_location_tracking != has_location_tracking:
        resolved = resolved or source_service.resolve_source(
            host_document,
            updated.get("source") or {},
        )
        if resolved.get("available"):
            source_state = copy.deepcopy(resolved.get("source_state") or {})
            updated["accepted_source_state"] = copy.deepcopy(source_state)
            updated["current_source_state"] = source_state
            updated["source_conditions"] = []
            comparison_engine.refresh_set_status(updated)
    else:
        comparison_engine.refresh_set_status(updated)
    updated["updated_at"] = models.utc_now_text()
    return _replace_set(store, updated), updated


def add_manual_child(host_document, store, set_id, parent_persistent_id, device):
    """Register a picked device as a Manual-origin linked child of a
    monitored element. The child carries its own device relationship so
    circuit context resolves on every scan."""
    tracking_set = copy.deepcopy(_require_set(store, set_id))
    if (tracking_set.get("source") or {}).get("source_type") != models.SOURCE_HOST:
        raise ValueError(
            "Add Device is only supported on host-source Tracking Sets."
        )
    parent_key = _text(parent_persistent_id or "")
    parent_record = (tracking_set.get("elements") or {}).get(parent_key)
    if parent_record is None or parent_record.get("state") == models.ELEMENT_REMOVED:
        raise ValueError("Select an available monitored element first.")
    unique_id = _text(getattr(device, "UniqueId", "") or "")
    if not unique_id:
        raise ValueError("The picked device has no stable identity.")
    child_key = "host:{}".format(unique_id)
    if child_key == parent_key:
        raise ValueError("An element cannot be its own linked child.")
    relationship = relationship_service.relationship_from_device(device)
    snapshot = _snapshot_element(
        host_document,
        host_document,
        tracking_set.get("source") or {},
        device,
        list(tracking_set.get("tracked_properties") or []),
        {},
        True,
        relationship=relationship,
        persistent_id_override=child_key,
    )
    existing = (tracking_set.get("elements") or {}).get(child_key)
    if existing is not None and existing.get("state") != models.ELEMENT_REMOVED:
        record = existing
        record["parent_persistent_id"] = parent_key
        record["relationship"] = copy.deepcopy(relationship)
        record["relationship_context"] = copy.deepcopy(
            snapshot.get("relationship_context")
        )
        if not record.get("track_location"):
            location = copy.deepcopy(snapshot.get("location"))
            record["track_location"] = True
            record["current_location"] = location
            record["accepted_location"] = copy.deepcopy(location)
    else:
        record = models.new_tracked_element(
            snapshot, baseline=True, track_location=True
        )
        record["parent_persistent_id"] = parent_key
        tracking_set["elements"][child_key] = record
    tracking_set["untracked_ids"] = [
        item for item in list(tracking_set.get("untracked_ids") or [])
        if item != child_key
    ]
    comparison_engine.recompute_record(record, tracking_set)
    comparison_engine.refresh_set_status(tracking_set)
    tracking_set["updated_at"] = models.utc_now_text()
    return _replace_set(store, tracking_set), tracking_set, record


def remove_manual_child(store, set_id, child_persistent_id):
    """Unlink a Manual-origin child: its record leaves the monitor entirely
    (parentless elements are not tracked). Profile children are managed by
    Sync Element Linker and cannot be unlinked here."""
    tracking_set = copy.deepcopy(_require_set(store, set_id))
    key = _text(child_persistent_id or "")
    record = (tracking_set.get("elements") or {}).get(key)
    if record is None:
        raise ValueError("The linked child no longer exists in the monitor.")
    if not _text(record.get("parent_persistent_id") or ""):
        raise ValueError("The selected element is not a linked child.")
    if record.get("linker_meta") is not None:
        raise ValueError(
            "Profile children are managed by Sync Element Linker; re-run the "
            "sync to update them."
        )
    del tracking_set["elements"][key]
    comparison_engine.refresh_set_status(tracking_set)
    tracking_set["updated_at"] = models.utc_now_text()
    return _replace_set(store, tracking_set), tracking_set


def set_summary(tracking_set):
    return comparison_engine.summarize_set(tracking_set)
