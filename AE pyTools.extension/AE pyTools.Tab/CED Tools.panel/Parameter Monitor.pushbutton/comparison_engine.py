# -*- coding: utf-8 -*-
"""Pure comparison/state-transition engine for Parameter Monitor."""

from __future__ import print_function

import copy
import math

import models
import text_service


def _safe_float(value, default=None):
    try:
        return float(value)
    except Exception:
        return default


def normalized_values_equal(accepted, current, double_tolerance=1.0e-9):
    accepted = accepted or models.missing_value()
    current = current or models.missing_value()
    if accepted.get("state") != current.get("state"):
        return False
    if accepted.get("storage_type") != current.get("storage_type"):
        return False
    if accepted.get("state") != models.VALUE_VALID:
        return accepted.get("raw") == current.get("raw")
    if accepted.get("storage_type") == "double":
        left = _safe_float(accepted.get("raw"))
        right = _safe_float(current.get("raw"))
        if left is None or right is None:
            return left == right
        return abs(left - right) <= float(double_tolerance or 0.0)
    return accepted.get("raw") == current.get("raw")


def _angular_delta(left, right):
    delta = abs(float(left or 0.0) - float(right or 0.0))
    full = math.pi * 2.0
    delta = delta % full
    return min(delta, full - delta)


def locations_equal(accepted, current, translation_tolerance, angular_tolerance):
    if accepted is None and current is None:
        return True
    if accepted is None or current is None:
        return False
    if accepted.get("state") != current.get("state"):
        return False
    if accepted.get("state") != models.VALUE_VALID:
        return True
    dx = _safe_float(accepted.get("x"), 0.0) - _safe_float(current.get("x"), 0.0)
    dy = _safe_float(accepted.get("y"), 0.0) - _safe_float(current.get("y"), 0.0)
    dz = _safe_float(accepted.get("z"), 0.0) - _safe_float(current.get("z"), 0.0)
    distance = math.sqrt((dx * dx) + (dy * dy) + (dz * dz))
    if distance > float(translation_tolerance or 0.0):
        return False
    return _angular_delta(accepted.get("rotation"), current.get("rotation")) <= float(
        angular_tolerance or 0.0
    )


def fingerprints_equal(left, right, tolerance=1.0e-9):
    left_values = list((left or {}).get("matrix") or [])
    right_values = list((right or {}).get("matrix") or [])
    if not left_values and not right_values:
        return True
    if len(left_values) != len(right_values):
        return False
    for left_value, right_value in zip(left_values, right_values):
        if abs(float(left_value) - float(right_value)) > float(tolerance or 0.0):
            return False
    return True


def metadata_value(metadata, key):
    """Read split family/type values while supporting older combined metadata."""
    metadata = metadata or {}
    direct = metadata.get(key)
    if direct is not None and text_service.to_text(direct) != "":
        return text_service.to_text(direct)
    combined = text_service.to_text(metadata.get("family_type") or "")
    parts = [item.strip() for item in combined.split(" : ", 1)]
    if key == "family_name" and parts:
        return parts[0]
    if key == "type_name" and len(parts) > 1:
        return parts[1]
    if key == "type_name" and len(parts) == 1:
        return parts[0]
    return ""


def recompute_record(record, tracking_set):
    if record.get("state") == models.ELEMENT_REMOVED:
        record["changed_property_keys"] = []
        record["change_count"] = 0
        record["missing_count"] = 0
        record["read_error_count"] = 0
        return record

    descriptors = list((tracking_set or {}).get("tracked_properties") or [])
    accepted = record.get("accepted_properties") or {}
    current = record.get("current_properties") or {}
    changed = []
    missing_count = 0
    read_error_count = 0
    for descriptor in descriptors:
        key = models.property_key(descriptor)
        current_value = current.get(key) or models.missing_value()
        if current_value.get("state") == models.VALUE_MISSING:
            missing_count += 1
        if current_value.get("state") in (models.VALUE_READ_ERROR, models.VALUE_UNSUPPORTED):
            read_error_count += 1
        if not normalized_values_equal(accepted.get(key), current_value):
            changed.append(key)

    if record.get("state") != models.ELEMENT_ADDED:
        accepted_metadata = record.get("accepted_metadata") or record.get("metadata") or {}
        current_metadata = record.get("current_metadata") or record.get("metadata") or {}
        if metadata_value(accepted_metadata, "family_name") != metadata_value(current_metadata, "family_name"):
            changed.append(models.FAMILY_PROPERTY_KEY)
        if metadata_value(accepted_metadata, "type_name") != metadata_value(current_metadata, "type_name"):
            changed.append(models.TYPE_PROPERTY_KEY)

    if bool(record.get("track_location")):
        defaults = (tracking_set or {}).get("location_defaults") or {}
        if not locations_equal(
            record.get("accepted_location"),
            record.get("current_location"),
            defaults.get("translation_tolerance", 0.001),
            defaults.get("angular_tolerance", 0.0017453292519943296),
        ):
            changed.append(models.LOCATION_PROPERTY_KEY)

    record["changed_property_keys"] = changed
    record["change_count"] = len(changed)
    record["missing_count"] = missing_count
    record["read_error_count"] = read_error_count
    return record


def summarize_set(tracking_set):
    untracked_ids = set((tracking_set or {}).get("untracked_ids") or [])
    summary = {
        "tracked": 0,
        "changed": 0,
        "added": 0,
        "removed": 0,
        "unchanged": 0,
        "missing_elements": 0,
        "untracked": len(list((tracking_set or {}).get("untracked_ids") or [])),
        "unresolved": 0,
    }
    for persistent_id, record in list(
        ((tracking_set or {}).get("elements") or {}).items()
    ):
        if persistent_id in untracked_ids:
            continue
        state = record.get("state")
        if state == models.ELEMENT_ADDED:
            summary["added"] += 1
            summary["unresolved"] += 1
        elif state == models.ELEMENT_REMOVED:
            summary["removed"] += 1
            summary["unresolved"] += 1
        else:
            summary["tracked"] += 1
            if int(record.get("change_count", 0) or 0) > 0:
                summary["changed"] += 1
                summary["unresolved"] += int(record.get("change_count", 0) or 0)
            else:
                summary["unchanged"] += 1
        if int(record.get("missing_count", 0) or 0) > 0:
            summary["missing_elements"] += 1
    if list((tracking_set or {}).get("source_conditions") or []):
        summary["unresolved"] += len(tracking_set.get("source_conditions") or [])
    return summary


def _set_status_from_summary(tracking_set):
    summary = summarize_set(tracking_set)
    if summary["unresolved"] > 0:
        tracking_set["status"] = models.SET_DIRTY
        tracking_set["status_message"] = "{} unresolved change(s).".format(summary["unresolved"])
    else:
        tracking_set["status"] = models.SET_CLEAN
        tracking_set["status_message"] = "No unresolved changes."
    return summary


def refresh_set_status(tracking_set):
    """Public status/count recomputation after non-scan monitor-data edits."""
    return _set_status_from_summary(tracking_set)


def apply_scan(tracking_set, current_map, checked_at, source_state=None):
    """Apply a successful scan without mutating Accepted values."""
    result = copy.deepcopy(tracking_set)
    current_map = current_map or {}
    elements = result.setdefault("elements", {})
    untracked = set(result.get("untracked_ids") or [])
    defaults = result.get("location_defaults") or {}
    track_new = bool(defaults.get("track_new_elements", False))

    for persistent_id in list(elements.keys()):
        if persistent_id in untracked:
            # Keep the last known record as an identity tombstone. Scans still
            # skip the element, but the Untracked view retains its ID, family,
            # type, level, and other useful context until tracking is restored.
            continue
        record = elements[persistent_id]
        snapshot = current_map.get(persistent_id)
        if snapshot is None:
            if record.get("state") == models.ELEMENT_ADDED:
                elements.pop(persistent_id, None)
                continue
            record["state"] = models.ELEMENT_REMOVED
            record["current_properties"] = {}
            record["current_location"] = None
            record["relationship_context"] = None
            recompute_record(record, result)
            continue
        current_metadata = copy.deepcopy(snapshot.get("metadata") or {})
        record["metadata"] = copy.deepcopy(current_metadata)
        record["current_metadata"] = current_metadata
        record["source_element_unique_id"] = text_service.to_text(
            snapshot.get("source_element_unique_id") or ""
        )
        record["current_properties"] = copy.deepcopy(snapshot.get("properties") or {})
        if bool(record.get("track_location")):
            record["current_location"] = copy.deepcopy(snapshot.get("location"))
        else:
            record["current_location"] = None
        record["relationship_context"] = copy.deepcopy(snapshot.get("relationship_context"))
        if record.get("state") != models.ELEMENT_ADDED:
            record["state"] = models.ELEMENT_TRACKED
        recompute_record(record, result)

    for persistent_id, snapshot in list(current_map.items()):
        if persistent_id in elements or persistent_id in untracked:
            continue
        record = models.new_tracked_element(
            snapshot,
            baseline=False,
            track_location=track_new,
        )
        recompute_record(record, result)
        elements[persistent_id] = record

    source_state = copy.deepcopy(source_state or {})
    result["current_source_state"] = source_state
    accepted_source = result.get("accepted_source_state") or {}
    result["source_conditions"] = []
    location_enabled = any([
        bool(record.get("track_location", False))
        for record in list((result.get("elements") or {}).values())
        if record.get("state") != models.ELEMENT_REMOVED
    ])
    if not location_enabled:
        result["accepted_source_state"] = copy.deepcopy(source_state)
    elif accepted_source and source_state and not fingerprints_equal(accepted_source, source_state):
        result["source_conditions"].append({
            "kind": "link_position_changed",
            "message": "Linked model position or rotation changed.",
        })
    result["last_check"] = checked_at
    result["updated_at"] = checked_at
    _set_status_from_summary(result)
    return result


def resolve_property(tracking_set, persistent_id, property_key):
    result = copy.deepcopy(tracking_set)
    record = (result.get("elements") or {}).get(
        text_service.to_text(persistent_id or "")
    )
    if record is None or record.get("state") == models.ELEMENT_REMOVED:
        return result
    key = text_service.to_text(property_key or "")
    if key == models.LOCATION_PROPERTY_KEY:
        if bool(record.get("track_location")):
            record["accepted_location"] = copy.deepcopy(record.get("current_location"))
    elif key in (models.FAMILY_PROPERTY_KEY, models.TYPE_PROPERTY_KEY):
        field = "family_name" if key == models.FAMILY_PROPERTY_KEY else "type_name"
        current_metadata = record.get("current_metadata") or record.get("metadata") or {}
        accepted_metadata = record.setdefault("accepted_metadata", {})
        accepted_metadata[field] = metadata_value(current_metadata, field)
    else:
        current = record.get("current_properties") or {}
        if key in current:
            record.setdefault("accepted_properties", {})[key] = copy.deepcopy(current[key])
    recompute_record(record, result)
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def resolve_element(tracking_set, persistent_id):
    result = copy.deepcopy(tracking_set)
    record = (result.get("elements") or {}).get(
        text_service.to_text(persistent_id or "")
    )
    if record is None or record.get("state") == models.ELEMENT_REMOVED:
        return result
    record["accepted_properties"] = copy.deepcopy(record.get("current_properties") or {})
    record["accepted_metadata"] = copy.deepcopy(
        record.get("current_metadata") or record.get("metadata") or {}
    )
    if bool(record.get("track_location")):
        record["accepted_location"] = copy.deepcopy(record.get("current_location"))
    record["state"] = models.ELEMENT_TRACKED
    recompute_record(record, result)
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def resolve_set(tracking_set):
    result = copy.deepcopy(tracking_set)
    untracked = set(result.get("untracked_ids") or [])
    for persistent_id in list((result.get("elements") or {}).keys()):
        if persistent_id in untracked:
            continue
        record = result["elements"][persistent_id]
        if record.get("state") == models.ELEMENT_REMOVED:
            continue
        record["accepted_properties"] = copy.deepcopy(record.get("current_properties") or {})
        record["accepted_metadata"] = copy.deepcopy(
            record.get("current_metadata") or record.get("metadata") or {}
        )
        if bool(record.get("track_location")):
            record["accepted_location"] = copy.deepcopy(record.get("current_location"))
        record["state"] = models.ELEMENT_TRACKED
        recompute_record(record, result)
    result["accepted_source_state"] = copy.deepcopy(result.get("current_source_state") or {})
    result["source_conditions"] = []
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def untrack_element(tracking_set, persistent_id):
    return untrack_elements(tracking_set, [persistent_id])


def untrack_elements(tracking_set, persistent_ids):
    """Untrack many records with one copy/status pass."""
    result = copy.deepcopy(tracking_set)
    keys = set([
        text_service.to_text(item or "") for item in list(persistent_ids or []) if item
    ])
    elements = result.get("elements") or {}
    untracked = set(result.get("untracked_ids") or [])
    for key in keys:
        record = elements.get(key)
        if record is None or record.get("state") == models.ELEMENT_REMOVED:
            continue
        untracked.add(key)
    result["untracked_ids"] = sorted(list(untracked))
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def restore_element(tracking_set, persistent_id, snapshot):
    result = copy.deepcopy(tracking_set)
    key = text_service.to_text(persistent_id or "")
    result["untracked_ids"] = [
        item for item in list(result.get("untracked_ids") or []) if item != key
    ]
    if snapshot is not None:
        track_location = bool((result.get("location_defaults") or {}).get("track_new_elements", False))
        result.setdefault("elements", {})[key] = models.new_tracked_element(
            snapshot, baseline=True, track_location=track_location
        )
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def remove_record(tracking_set, persistent_id):
    result = copy.deepcopy(tracking_set)
    key = text_service.to_text(persistent_id or "")
    record = (result.get("elements") or {}).get(key)
    if record is not None and record.get("state") == models.ELEMENT_REMOVED:
        result["elements"].pop(key, None)
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def remove_all_removed(tracking_set):
    result = copy.deepcopy(tracking_set)
    for key in list((result.get("elements") or {}).keys()):
        if result["elements"][key].get("state") == models.ELEMENT_REMOVED:
            result["elements"].pop(key, None)
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def set_location_tracking(tracking_set, persistent_id, enabled, current_location=None):
    result = copy.deepcopy(tracking_set)
    record = (result.get("elements") or {}).get(
        text_service.to_text(persistent_id or "")
    )
    if record is None or record.get("state") == models.ELEMENT_REMOVED:
        return result
    record["track_location"] = bool(enabled)
    if enabled:
        location = copy.deepcopy(current_location if current_location is not None else record.get("current_location"))
        record["current_location"] = location
        record["accepted_location"] = copy.deepcopy(location)
    else:
        record["current_location"] = None
        record["accepted_location"] = None
    recompute_record(record, result)
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result


def update_tracked_properties(tracking_set, descriptors, current_map=None):
    result = copy.deepcopy(tracking_set)
    old_keys = set([models.property_key(item) for item in result.get("tracked_properties") or []])
    new_descriptors = copy.deepcopy(list(descriptors or []))
    new_keys = set([models.property_key(item) for item in new_descriptors])
    added_keys = new_keys - old_keys
    removed_keys = old_keys - new_keys
    result["tracked_properties"] = new_descriptors
    current_map = current_map or {}
    untracked = set(result.get("untracked_ids") or [])

    for persistent_id, record in list((result.get("elements") or {}).items()):
        if persistent_id in untracked:
            continue
        accepted = record.setdefault("accepted_properties", {})
        current = record.setdefault("current_properties", {})
        for key in removed_keys:
            accepted.pop(key, None)
            current.pop(key, None)
        snapshot = current_map.get(persistent_id) or {}
        snapshot_values = snapshot.get("properties") or {}
        for key in added_keys:
            value = copy.deepcopy(snapshot_values.get(key) or models.missing_value())
            accepted[key] = copy.deepcopy(value)
            current[key] = value
        recompute_record(record, result)
    _set_status_from_summary(result)
    result["updated_at"] = models.utc_now_text()
    return result
