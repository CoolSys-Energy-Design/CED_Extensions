# -*- coding: utf-8 -*-
"""Plain, JSON-safe models for Parameter Monitor.

The Revit-facing services deliberately exchange dictionaries with this module.
Keeping the persisted shape free of Revit/.NET objects makes migrations and
comparison tests runnable outside Revit and compatible with IronPython 2.7.
"""

from __future__ import print_function

import copy
import datetime
import uuid

PAYLOAD_SCHEMA_VERSION = 1
DEFINITION_SCHEMA_VERSION = 1
TOOL_VERSION = "1.0.0"

SOURCE_HOST = "host"
SOURCE_LINK = "link"

SET_NEVER_CHECKED = "never_checked"
SET_CLEAN = "clean"
SET_DIRTY = "changes_detected"
SET_CHECKING = "checking"
SET_SOURCE_UNAVAILABLE = "source_unavailable"
SET_CHECK_FAILED = "check_failed"

ELEMENT_TRACKED = "tracked"
ELEMENT_ADDED = "added"
ELEMENT_REMOVED = "removed"

VALUE_VALID = "valid"
VALUE_MISSING = "missing"
VALUE_BLANK = "blank"
VALUE_UNSUPPORTED = "unsupported"
VALUE_READ_ERROR = "read_error"

LOCATION_PROPERTY_KEY = "__location__"
FAMILY_PROPERTY_KEY = "__family__"
TYPE_PROPERTY_KEY = "__type__"


class PayloadMigrationError(ValueError):
    pass


def utc_now_text():
    return datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")


def new_id():
    return str(uuid.uuid4())


def copy_json_value(value):
    return copy.deepcopy(value)


def normalized_value(state, storage_type=None, raw=None, display=None, message=None):
    value = {
        "state": str(state or VALUE_UNSUPPORTED),
        "storage_type": str(storage_type or "none"),
        "raw": raw,
        "display": display,
    }
    if message:
        value["message"] = str(message)
    return value


def missing_value(display="Missing Parameter"):
    return normalized_value(VALUE_MISSING, "none", None, display)


def new_project_store(project_identity=None):
    now = utc_now_text()
    return {
        "payload_schema_version": PAYLOAD_SCHEMA_VERSION,
        "tool_version": TOOL_VERSION,
        "project_identity": copy_json_value(project_identity or {}),
        "tracking_sets": [],
        "created_at": now,
        "updated_at": now,
    }


def new_tracking_set(name, source, category, properties, location_defaults=None):
    now = utc_now_text()
    return {
        "set_id": new_id(),
        "name": str(name or category.get("name") or "Tracking Set"),
        "source": copy_json_value(source or {}),
        "category": copy_json_value(category or {}),
        "active": True,
        "scan_policy": "manual",
        "tracked_properties": copy_json_value(list(properties or [])),
        "location_defaults": copy_json_value(location_defaults or {
            "track_new_elements": False,
            "translation_tolerance": 0.001,
            "angular_tolerance": 0.0017453292519943296,
        }),
        "last_check": None,
        "status": SET_NEVER_CHECKED,
        "status_message": "Baseline has not been created.",
        "accepted_source_state": {},
        "current_source_state": {},
        "source_conditions": [],
        "elements": {},
        "untracked_ids": [],
        "created_at": now,
        "updated_at": now,
    }


def new_tracked_element(snapshot, baseline=True, track_location=False):
    snapshot = snapshot or {}
    current_metadata = copy_json_value(snapshot.get("metadata") or {})
    current_properties = copy_json_value(snapshot.get("properties") or {})
    current_location = copy_json_value(snapshot.get("location"))
    state = ELEMENT_TRACKED if baseline else ELEMENT_ADDED
    return {
        "persistent_id": str(snapshot.get("persistent_id") or ""),
        "source_element_unique_id": str(snapshot.get("source_element_unique_id") or ""),
        # metadata remains a compatibility alias for the current display state.
        "metadata": copy_json_value(current_metadata),
        "accepted_metadata": copy_json_value(current_metadata) if baseline else {},
        "current_metadata": current_metadata,
        "accepted_properties": copy_json_value(current_properties) if baseline else {},
        "current_properties": current_properties,
        "track_location": bool(track_location),
        "accepted_location": copy_json_value(current_location) if baseline and track_location else None,
        "current_location": current_location if track_location else None,
        "relationship": copy_json_value(snapshot.get("relationship")),
        "relationship_context": copy_json_value(snapshot.get("relationship_context")),
        "state": state,
        "changed_property_keys": [],
        "change_count": 0,
        "missing_count": 0,
        "read_error_count": 0,
    }


def find_set(store, set_id):
    target = str(set_id or "")
    for tracking_set in list((store or {}).get("tracking_sets") or []):
        if str(tracking_set.get("set_id") or "") == target:
            return tracking_set
    return None


def property_key(descriptor):
    return str((descriptor or {}).get("key") or "")


def _ensure_tracking_set_shape(tracking_set):
    tracking_set.setdefault("set_id", new_id())
    tracking_set.setdefault("name", "Tracking Set")
    tracking_set.setdefault("source", {})
    tracking_set.setdefault("category", {})
    tracking_set.setdefault("active", True)
    tracking_set.setdefault("scan_policy", "manual")
    tracking_set.setdefault("tracked_properties", [])
    tracking_set.setdefault("location_defaults", {})
    defaults = tracking_set["location_defaults"]
    defaults.setdefault("track_new_elements", False)
    defaults.setdefault("translation_tolerance", 0.001)
    defaults.setdefault("angular_tolerance", 0.0017453292519943296)
    tracking_set.setdefault("last_check", None)
    tracking_set.setdefault("status", SET_NEVER_CHECKED)
    tracking_set.setdefault("status_message", "")
    tracking_set.setdefault("accepted_source_state", {})
    tracking_set.setdefault("current_source_state", {})
    tracking_set.setdefault("source_conditions", [])
    tracking_set.setdefault("elements", {})
    tracking_set.setdefault("untracked_ids", [])
    tracking_set.setdefault("created_at", utc_now_text())
    tracking_set.setdefault("updated_at", utc_now_text())
    for persistent_id, record in list(tracking_set["elements"].items()):
        record.setdefault("persistent_id", str(persistent_id))
        record.setdefault("source_element_unique_id", "")
        record.setdefault("metadata", {})
        # Existing payloads only stored the latest metadata. Treat that value as
        # the initial accepted/current metadata so future family/type changes can
        # be detected without inventing a historical change.
        record.setdefault("accepted_metadata", copy_json_value(record.get("metadata") or {}))
        record.setdefault("current_metadata", copy_json_value(record.get("metadata") or {}))
        record.setdefault("accepted_properties", {})
        record.setdefault("current_properties", {})
        record.setdefault("track_location", False)
        record.setdefault("accepted_location", None)
        record.setdefault("current_location", None)
        record.setdefault("relationship", None)
        record.setdefault("relationship_context", None)
        record.setdefault("state", ELEMENT_TRACKED)
        record.setdefault("changed_property_keys", [])
        record.setdefault("change_count", 0)
        record.setdefault("missing_count", 0)
        record.setdefault("read_error_count", 0)
    return tracking_set


def migrate_project_store(raw_store, project_identity=None):
    if raw_store is None:
        return new_project_store(project_identity=project_identity)
    if not isinstance(raw_store, dict):
        raise PayloadMigrationError("Parameter Monitor payload must be a JSON object.")
    store = copy_json_value(raw_store)
    version = int(store.get("payload_schema_version", 0) or 0)
    if version > PAYLOAD_SCHEMA_VERSION:
        raise PayloadMigrationError(
            "Payload schema {} is newer than supported schema {}.".format(
                version, PAYLOAD_SCHEMA_VERSION
            )
        )
    if version == 0:
        store["payload_schema_version"] = 1
        version = 1
    if version != PAYLOAD_SCHEMA_VERSION:
        raise PayloadMigrationError("Unsupported payload schema {}.".format(version))
    store.setdefault("tool_version", TOOL_VERSION)
    store.setdefault("project_identity", copy_json_value(project_identity or {}))
    store.setdefault("tracking_sets", [])
    store.setdefault("created_at", utc_now_text())
    store.setdefault("updated_at", utc_now_text())
    for tracking_set in list(store.get("tracking_sets") or []):
        _ensure_tracking_set_shape(tracking_set)
    return store
