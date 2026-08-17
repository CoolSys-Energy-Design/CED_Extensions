# -*- coding: utf-8 -*-
"""Lightweight WPF row models built from the persisted JSON-safe state."""

from __future__ import print_function

import comparison_engine
import models
import text_service
import tracking_service

FILTER_ALL = "all"
FILTER_CHANGES = "changes"
FILTER_REMOVED = "removed"
FILTER_MISSING = "missing"
FILTER_UNTRACKED = "untracked"


class FilterOption(object):
    def __init__(self, key, label):
        self.key = key
        self.label = label


FILTER_OPTIONS = [
    FilterOption(FILTER_ALL, "All Elements"),
    FilterOption(FILTER_CHANGES, "Changes Only"),
    FilterOption(FILTER_REMOVED, "Removed Only"),
    FilterOption(FILTER_MISSING, "Missing Parameters"),
    FilterOption(FILTER_UNTRACKED, "Untracked"),
]


STATUS_LABELS = {
    models.SET_NEVER_CHECKED: "Never Checked",
    models.SET_CLEAN: "Clean",
    models.SET_DIRTY: "Changes Detected",
    models.SET_CHECKING: "Checking",
    models.SET_SOURCE_UNAVAILABLE: "Source Unavailable",
    models.SET_LINK_UNAVAILABLE: "Link Unavailable",
    models.SET_LINK_WORKSETS_UNAVAILABLE: "Link Worksets Unavailable",
    models.SET_CHECK_FAILED: "Check Failed",
}


def _display(value, fallback="-"):
    if value is None or text_service.to_text(value) == "":
        return fallback
    return text_service.to_text(value)


def _value_display(value):
    value = value or {}
    state = value.get("state")
    if state == models.VALUE_MISSING:
        return "Missing Parameter"
    if state == models.VALUE_BLANK:
        return "(blank)"
    if state == models.VALUE_UNSUPPORTED:
        return value.get("display") or "Unsupported"
    if state == models.VALUE_READ_ERROR:
        return value.get("display") or "Read Error"
    return _display(value.get("display"), "(no value)")


def _location_display(location):
    location = location or {}
    if location.get("state") != models.VALUE_VALID:
        return location.get("message") or "Location Not Supported"
    return "X {:.3f}, Y {:.3f}, Z {:.3f}, R {:.2f} deg".format(
        float(location.get("x", 0.0)),
        float(location.get("y", 0.0)),
        float(location.get("z", 0.0)),
        float(location.get("rotation", 0.0)) * 57.29577951308232,
    )


def _metadata_value(metadata, key):
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
    return "-"


class TrackingSetRow(object):
    def __init__(self, tracking_set):
        self.data = tracking_set
        self.set_id = text_service.to_text(tracking_set.get("set_id") or "")
        self.name = text_service.to_text(tracking_set.get("name") or "Tracking Set")
        self.source = text_service.to_text((tracking_set.get("source") or {}).get("display_name") or "Host Model")
        self.category = text_service.to_text((tracking_set.get("category") or {}).get("name") or "Category")
        self.active = bool(tracking_set.get("active", True))
        self.active_text = "ACTIVE" if self.active else "DEACTIVATED"
        self.status = text_service.to_text(tracking_set.get("status") or models.SET_NEVER_CHECKED)
        self.status_text = STATUS_LABELS.get(self.status, self.status.replace("_", " ").title())
        self.status_message = text_service.to_text(tracking_set.get("status_message") or "")
        self.source_condition_text = "; ".join([
            text_service.to_text(item.get("message") or item.get("kind") or "Source condition")
            for item in list(tracking_set.get("source_conditions") or [])
        ])
        self.last_check = _display(tracking_set.get("last_check"), "Never")
        self.property_count = len(tracking_set.get("tracked_properties") or [])
        # The cards and main-grid summary describe the records that can appear
        # in the main grid. Element Linker children live in the inspector and
        # must not inflate these element counts. The set-level unresolved flag
        # still uses the complete persisted summary so child changes are not
        # silently ignored.
        self.full_summary = tracking_service.set_summary(tracking_set)
        self.summary = main_grid_summary(tracking_set)
        self.element_count = sum([
            int(self.summary.get(key, 0) or 0)
            for key in ("changed", "added", "removed", "unchanged")
        ])
        self.unresolved_count = int(self.full_summary.get("unresolved", 0) or 0)
        self.has_changes = self.unresolved_count > 0
        source_type = text_service.to_text((tracking_set.get("source") or {}).get("source_type") or "")
        self.model_text = "Model: {}".format(
            "Link" if source_type == models.SOURCE_LINK else "Host"
        )
        self.changed_text = "{} changed".format(self.summary.get("changed", 0))
        self.added_text = "{} added".format(self.summary.get("added", 0))
        self.removed_text = "{} removed".format(self.summary.get("removed", 0))
        self.element_count_text = "{} element{}".format(
            self.element_count, "" if self.element_count == 1 else "s"
        )
        self.subtitle = "Category: {} | Param Sets: {} | Model: {} | {} elements".format(
            self.category, self.property_count, self.source, self.element_count
        )
        self.counts_text = "{} changed  |  {} added  |  {} removed".format(
            self.summary.get("changed", 0),
            self.summary.get("added", 0),
            self.summary.get("removed", 0),
        )


class ElementRow(object):
    def __init__(self, persistent_id, record=None, untracked=False):
        self.persistent_id = text_service.to_text(persistent_id or "")
        self.record = record
        self.is_untracked = bool(untracked)
        metadata = (
            (record or {}).get("current_metadata")
            or (record or {}).get("metadata")
            or {}
        )
        state = (record or {}).get("state")
        changes = int((record or {}).get("change_count", 0) or 0)
        missing = int((record or {}).get("missing_count", 0) or 0)
        if untracked:
            self.status = "Untracked"
        elif state == models.ELEMENT_REMOVED:
            self.status = "Removed"
        elif state == models.ELEMENT_ADDED:
            self.status = "Added"
        elif changes > 0:
            self.status = "Changed"
        else:
            self.status = "Unchanged"
        self.element = _display(metadata.get("friendly_name"), "Untracked Element" if untracked else "Element")
        self.family = _display(_metadata_value(metadata, "family_name"))
        self.type = _display(_metadata_value(metadata, "type_name"))
        self.family_type = "{} : {}".format(self.family, self.type)
        self.element_id = _display(metadata.get("element_id"))
        self.level = _display(metadata.get("level"))
        self.change_count = changes
        self.missing_count = missing
        self.missing_text = str(missing) if missing else "-"
        self.missing_state = "Missing" if missing else "None"
        changed_keys = set((record or {}).get("changed_property_keys") or [])
        special_keys = set([
            models.LOCATION_PROPERTY_KEY,
            models.FAMILY_PROPERTY_KEY,
            models.TYPE_PROPERTY_KEY,
        ])
        parameter_change_keys = [key for key in changed_keys if key not in special_keys]
        self.parameter_change_count = len(parameter_change_keys)
        if not changed_keys and changes:
            self.parameter_change_count = changes
        self.parameter_change_text = str(self.parameter_change_count) if self.parameter_change_count else "-"
        self.parameter_change_state = "Changed" if self.parameter_change_count else "Unchanged"
        self.location_change_count = 1 if models.LOCATION_PROPERTY_KEY in changed_keys else 0
        self.location_change_text = "Changed" if self.location_change_count else "-"
        self.location_change_state = "Changed" if self.location_change_count else "Unchanged"
        self.family_change_text = "Changed" if models.FAMILY_PROPERTY_KEY in changed_keys else "-"
        self.family_change_state = "Changed" if models.FAMILY_PROPERTY_KEY in changed_keys else "Unchanged"
        self.type_change_text = "Changed" if models.TYPE_PROPERTY_KEY in changed_keys else "-"
        self.type_change_state = "Changed" if models.TYPE_PROPERTY_KEY in changed_keys else "Unchanged"
        self.location_text = "On" if bool((record or {}).get("track_location", False)) else "Off"
        context = (record or {}).get("relationship_context") or {}
        self.circuit = _display(context.get("status_text"))
        self.can_navigate = (not untracked) and state != models.ELEMENT_REMOVED
        self.can_resolve = (not untracked) and state != models.ELEMENT_REMOVED and (
            state == models.ELEMENT_ADDED or changes > 0
        )
        self.can_untrack = (not untracked) and state != models.ELEMENT_REMOVED
        self.can_restore = untracked
        self.can_remove_record = state == models.ELEMENT_REMOVED
        self.has_relationship = bool((record or {}).get("relationship"))
        self.parent_persistent_id = text_service.to_text((record or {}).get("parent_persistent_id") or "")


class PropertyRow(object):
    def __init__(self, key, name, accepted, current, changed, value_state, scope=""):
        self.key = text_service.to_text(key or "")
        self.name = text_service.to_text(name or "Property")
        self.scope = text_service.to_text(scope or "")
        self.accepted = text_service.to_text(accepted or "-")
        self.current = text_service.to_text(current or "-")
        self.changed = bool(changed)
        self.state = "Changed" if changed else text_service.to_text(value_state or "Unchanged").replace("_", " ").title()
        self.can_resolve = bool(changed)


class LinkedChildRow(object):
    """One row of the LINKED CHILDREN panel list."""

    def __init__(self, persistent_id, record):
        record = record or {}
        metadata = (
            record.get("current_metadata")
            or record.get("metadata")
            or {}
        )
        self.persistent_id = text_service.to_text(persistent_id or "")
        self.family = _display(_metadata_value(metadata, "family_name"))
        self.type = _display(_metadata_value(metadata, "type_name"))
        # Profile = registered by the Element Linker sync (profile-placed);
        # Manual = a monitored element linked by hand.
        self.origin = "Profile" if record.get("linker_meta") else "Manual"
        self.state = text_service.to_text(record.get("state") or models.ELEMENT_TRACKED)


def linked_children_info(tracking_set, record):
    """Children of a selected (parent) record, plus follow-move state.

    ``parent_moved`` is True when the selected record's own accepted vs
    current locations differ beyond the set's tolerances — the condition
    under which its children can follow. ``movable_child_ids`` are the
    non-removed children to pass to the move operation.
    """
    info = {
        "children": [],
        "count": 0,
        "parent_moved": False,
        "movable_child_ids": [],
    }
    if tracking_set is None or record is None:
        return info
    parent_key = text_service.to_text(record.get("persistent_id") or "")
    if not parent_key:
        return info
    children = []
    for persistent_id, candidate in sorted(
        (tracking_set.get("elements") or {}).items()
    ):
        if text_service.to_text((candidate or {}).get("parent_persistent_id") or "") != parent_key:
            continue
        children.append((persistent_id, candidate))
    if not children:
        return info
    info["children"] = [
        LinkedChildRow(persistent_id, candidate)
        for persistent_id, candidate in children
    ]
    info["count"] = len(children)
    if bool(record.get("track_location", False)) and record.get(
        "state"
    ) != models.ELEMENT_REMOVED:
        defaults = (tracking_set.get("location_defaults") or {})
        info["parent_moved"] = not comparison_engine.locations_equal(
            record.get("accepted_location"),
            record.get("current_location"),
            defaults.get("translation_tolerance", 0.001),
            defaults.get("angular_tolerance", 0.0017453292519943296),
        )
    if info["parent_moved"]:
        # Linked-model children are monitor-only: Revit cannot edit link
        # documents, so only host children can follow the parent.
        info["movable_child_ids"] = [
            persistent_id
            for persistent_id, candidate in children
            if (candidate or {}).get("state") != models.ELEMENT_REMOVED
            and not text_service.to_text(persistent_id or "").startswith("link:")
        ]
    return info


def tracking_set_rows(store):
    return [TrackingSetRow(item) for item in list((store or {}).get("tracking_sets") or [])]


def element_rows(tracking_set, filter_key=FILTER_ALL, search_text=""):
    if tracking_set is None:
        return []
    rows = []
    if filter_key == FILTER_UNTRACKED:
        for persistent_id in list(tracking_set.get("untracked_ids") or []):
            rows.append(ElementRow(persistent_id, record=None, untracked=True))
        return rows
    for persistent_id, record in list((tracking_set.get("elements") or {}).items()):
        # Element Linker children are filed under their parent's LINKED
        # CHILDREN panel instead of cluttering the main grid.
        if text_service.to_text((record or {}).get("parent_persistent_id") or ""):
            continue
        row = ElementRow(persistent_id, record=record)
        if tracking_set.get("status") in (
            models.SET_SOURCE_UNAVAILABLE,
            models.SET_LINK_UNAVAILABLE,
            models.SET_LINK_WORKSETS_UNAVAILABLE,
            models.SET_CHECK_FAILED,
        ):
            row.can_navigate = False
            row.can_resolve = False
        if filter_key == FILTER_CHANGES and row.status not in ("Changed", "Added"):
            continue
        if filter_key == FILTER_REMOVED and row.status != "Removed":
            continue
        if filter_key == FILTER_MISSING and row.missing_count <= 0:
            continue
        rows.append(row)
    search = text_service.to_text(search_text or "").strip().lower()
    if search:
        rows = [row for row in rows if search in " ".join([
            row.element.lower(), row.family.lower(), row.type.lower(), row.element_id.lower(),
            row.level.lower(), row.persistent_id.lower(), row.circuit.lower(),
        ])]
    priority = {"Changed": 0, "Added": 1, "Removed": 2, "Unchanged": 3}
    rows.sort(key=lambda row: (priority.get(row.status, 9), row.element.lower(), row.persistent_id))
    return rows


def summarize_element_rows(rows):
    """Count the statuses represented by a main-grid row collection."""
    summary = {"changed": 0, "added": 0, "removed": 0, "unchanged": 0}
    for row in list(rows or []):
        key = text_service.to_text(getattr(row, "status", "") or "").lower()
        if key in summary:
            summary[key] += 1
    return summary


def main_grid_summary(tracking_set):
    """Return unfiltered counts for root records shown by the main grid."""
    if tracking_set is None:
        return summarize_element_rows([])
    rows = []
    for persistent_id, record in list((tracking_set.get("elements") or {}).items()):
        if text_service.to_text((record or {}).get("parent_persistent_id") or ""):
            continue
        rows.append(ElementRow(persistent_id, record=record))
    return summarize_element_rows(rows)


def property_rows(tracking_set, element_row):
    if tracking_set is None or element_row is None or element_row.record is None:
        return []
    record = element_row.record
    changed_keys = set(record.get("changed_property_keys") or [])
    accepted = record.get("accepted_properties") or {}
    current = record.get("current_properties") or {}
    accepted_metadata = record.get("accepted_metadata") or record.get("metadata") or {}
    current_metadata = record.get("current_metadata") or record.get("metadata") or {}
    rows = [
        PropertyRow(
            models.FAMILY_PROPERTY_KEY,
            "Family",
            _metadata_value(accepted_metadata, "family_name"),
            _metadata_value(current_metadata, "family_name"),
            models.FAMILY_PROPERTY_KEY in changed_keys,
            "valid",
            scope="Identity",
        ),
        PropertyRow(
            models.TYPE_PROPERTY_KEY,
            "Type",
            _metadata_value(accepted_metadata, "type_name"),
            _metadata_value(current_metadata, "type_name"),
            models.TYPE_PROPERTY_KEY in changed_keys,
            "valid",
            scope="Identity",
        ),
    ]
    for descriptor in list(tracking_set.get("tracked_properties") or []):
        key = text_service.to_text(descriptor.get("key") or "")
        accepted_value = accepted.get(key) or {}
        current_value = current.get(key) or {}
        rows.append(PropertyRow(
            key,
            descriptor.get("name") or key,
            _value_display(accepted_value),
            _value_display(current_value),
            key in changed_keys,
            current_value.get("state") or "unchanged",
            scope=(descriptor.get("scope") or "").title(),
        ))
    if bool(record.get("track_location")):
        rows.append(PropertyRow(
            models.LOCATION_PROPERTY_KEY,
            "Location / Rotation",
            _location_display(record.get("accepted_location")),
            _location_display(record.get("current_location")),
            models.LOCATION_PROPERTY_KEY in changed_keys,
            (record.get("current_location") or {}).get("state") or "unsupported",
            scope="Special",
        ))
    return rows
