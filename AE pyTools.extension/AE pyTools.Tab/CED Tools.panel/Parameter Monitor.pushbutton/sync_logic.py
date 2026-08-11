# -*- coding: utf-8 -*-
"""Pure decision + membership logic for the Element Linker sync.

No Revit imports: everything operates on plain dicts so the tests under
``tests/`` cover grouping, tie-breaking, and idempotent membership updates
outside Revit.
"""

from __future__ import print_function

import copy

import comparison_engine
import models

PARENT_DIRECTIVE_KEY = "parent_parameter"
SIBLING_DIRECTIVE_KEY = "sibling_parameter"


def _is_directive(value):
    return isinstance(value, dict) and (
        PARENT_DIRECTIVE_KEY in value or SIBLING_DIRECTIVE_KEY in value
    )


def _led_has_directives(led):
    parameters = (led or {}).get("parameters")
    if not isinstance(parameters, dict):
        return False
    for value in parameters.values():
        if _is_directive(value):
            return True
    return False


def led_directive_index(profile_data):
    """Map every LED id in the active profile payload to a has-directives flag.

    Walks ``equipment_definitions`` and ``space_profiles`` (both share the
    profile -> linked_sets -> linked_element_definitions shape).
    """
    index = {}
    profile_data = profile_data if isinstance(profile_data, dict) else {}
    for root_key in ("equipment_definitions", "space_profiles"):
        for profile in list(profile_data.get(root_key) or []):
            if not isinstance(profile, dict):
                continue
            for linked_set in list(profile.get("linked_sets") or []):
                if not isinstance(linked_set, dict):
                    continue
                for led in list(linked_set.get("linked_element_definitions") or []):
                    if not isinstance(led, dict):
                        continue
                    led_id = str(led.get("id") or "")
                    if not led_id:
                        continue
                    index[led_id] = bool(
                        index.get(led_id, False) or _led_has_directives(led)
                    )
    return index


def group_children(children):
    """Group child-info dicts by their parent identity.

    Prefers ``parent_persistent_id`` (unambiguous across host and linked
    documents) and falls back to ``parent_unique_id``. Parentless children
    are excluded — they are tracked standalone.
    """
    groups = {}
    for child in list(children or []):
        parent_key = str(
            (child or {}).get("parent_persistent_id")
            or (child or {}).get("parent_unique_id")
            or ""
        )
        if not parent_key:
            continue
        groups.setdefault(parent_key, []).append(child)
    return groups


def find_sync_sets(store):
    """Every tracking set owned by the Element Linker sync (any category)."""
    return [
        tracking_set
        for tracking_set in list((store or {}).get("tracking_sets") or [])
        if str(tracking_set.get("origin") or "") == models.SET_ORIGIN_ELEMENT_LINKER
    ]


def sync_set_name(category_name):
    return "Element Linker - {}".format(category_name)


def _category_of(snapshot):
    category = str(((snapshot or {}).get("metadata") or {}).get("category") or "")
    return category or "Other"


def _find_category_set(store, category_name):
    target = str(category_name or "").strip().lower()
    for tracking_set in find_sync_sets(store):
        name = str((tracking_set.get("category") or {}).get("name") or "")
        if name.strip().lower() == target:
            return tracking_set
    return None


def _new_category_set(source_descriptor, category_name):
    return models.new_tracking_set(
        sync_set_name(category_name),
        source_descriptor,
        {"id": None, "name": str(category_name)},
        [],
        location_defaults={
            "track_new_elements": True,
            "translation_tolerance": 0.001,
            "angular_tolerance": 0.0017453292519943296,
        },
        membership=models.MEMBERSHIP_EXPLICIT,
        origin=models.SET_ORIGIN_ELEMENT_LINKER,
    )


def _apply_entry(tracking_set, entry, snapshots, now, report, set_chosen):
    """Upsert one entry into a sync set, preserving accepted baselines."""
    persistent_id = str((entry or {}).get("persistent_id") or "")
    if not persistent_id or persistent_id in set_chosen:
        return
    snapshot = (snapshots or {}).get(persistent_id)
    if snapshot is None:
        report["warnings"].append(
            "No snapshot for {}; skipped.".format(persistent_id)
        )
        return
    set_chosen.add(persistent_id)
    elements = tracking_set.setdefault("elements", {})
    untracked = list(tracking_set.get("untracked_ids") or [])
    if persistent_id in untracked:
        untracked = [item for item in untracked if item != persistent_id]
        elements.pop(persistent_id, None)
    tracking_set["untracked_ids"] = untracked

    linker_meta = copy.deepcopy(entry.get("linker_meta")) or {}
    linker_meta.setdefault("role", str(entry.get("role") or "child"))
    linker_meta["synced_at"] = now
    parent_persistent_id = entry.get("parent_persistent_id") or None

    record = elements.get(persistent_id)
    if record is None or record.get("state") == models.ELEMENT_REMOVED:
        was_removed = record is not None
        new_record = models.new_tracked_element(
            snapshot, baseline=True, track_location=True
        )
        new_record["parent_persistent_id"] = parent_persistent_id
        new_record["linker_meta"] = linker_meta
        elements[persistent_id] = new_record
        record = new_record
        if was_removed:
            report["restored"] += 1
        else:
            report["added"] += 1
    else:
        record["parent_persistent_id"] = parent_persistent_id
        record["linker_meta"] = linker_meta
        if not record.get("track_location"):
            location = copy.deepcopy(snapshot.get("location"))
            record["track_location"] = True
            record["current_location"] = location
            record["accepted_location"] = copy.deepcopy(location)
        # Adopt the snapshot's self-pointing device relationship so circuit
        # context (Select Circuit) works on records synced before it existed.
        snapshot_relationship = snapshot.get("relationship")
        if snapshot_relationship and not record.get("relationship"):
            record["relationship"] = copy.deepcopy(snapshot_relationship)
        if record.get("relationship"):
            record["relationship_context"] = copy.deepcopy(
                snapshot.get("relationship_context")
            ) or record.get("relationship_context")
        report["refreshed"] += 1
    comparison_engine.recompute_record(record, tracking_set)


def apply_sync_membership(store, entries, snapshots, source_descriptor, now_text=None):
    """Idempotently apply the chosen membership, one sync set per category.

    ``entries`` is a list of dicts: ``{"persistent_id", "role" (child/parent),
    "parent_persistent_id" (children only), "linker_meta"}``. Children are
    bucketed by their PARENT's category into "Element Linker - {Category}"
    sets (the parent equipment defines the set; its category falls back to
    the child's when the parent has no snapshot); each parent is filed into
    every category set that holds one of its children (so per-set children
    lists and follow-moves stay self-contained).
    Existing records keep their accepted baselines untouched. Stale sync
    records (linker_meta set but no longer chosen for that set) are removed
    when clean, kept + reported otherwise; sync sets left empty (e.g. the
    legacy single mixed set) are deleted.
    """
    now = now_text or models.utc_now_text()
    result = copy.deepcopy(store or models.new_project_store())

    report = {
        "sets_created": [],
        "sets_updated": [],
        "sets_removed": [],
        "added": 0,
        "refreshed": 0,
        "restored": 0,
        "stale_removed": 0,
        "stale_kept": [],
        "warnings": [],
    }

    parent_entries = {}
    child_entries = []
    for entry in list(entries or []):
        persistent_id = str((entry or {}).get("persistent_id") or "")
        if not persistent_id:
            continue
        if str(entry.get("role") or "child") == "parent":
            parent_entries[persistent_id] = entry
        else:
            child_entries.append(entry)

    buckets = {}
    for entry in child_entries:
        persistent_id = entry["persistent_id"]
        snapshot = (snapshots or {}).get(persistent_id)
        if snapshot is None:
            report["warnings"].append(
                "No snapshot for {}; skipped.".format(persistent_id)
            )
            continue
        parent_pid = str(entry.get("parent_persistent_id") or "")
        parent_snapshot = (snapshots or {}).get(parent_pid) if parent_pid else None
        category_source = parent_snapshot if parent_snapshot is not None else snapshot
        buckets.setdefault(_category_of(category_source), []).append(entry)

    chosen_per_set = {}
    primary_set_id = None
    for category_name in sorted(buckets.keys()):
        bucket = buckets[category_name]
        tracking_set = _find_category_set(result, category_name)
        if tracking_set is None:
            tracking_set = _new_category_set(source_descriptor, category_name)
            result.setdefault("tracking_sets", []).append(tracking_set)
            report["sets_created"].append(str(tracking_set.get("name")))
        else:
            report["sets_updated"].append(str(tracking_set.get("name")))
        tracking_set["membership"] = models.MEMBERSHIP_EXPLICIT
        tracking_set["origin"] = models.SET_ORIGIN_ELEMENT_LINKER
        set_chosen = chosen_per_set.setdefault(
            str(tracking_set.get("set_id") or ""), set()
        )
        bucket_entries = list(bucket)
        for entry in bucket:
            parent_pid = str(entry.get("parent_persistent_id") or "")
            if parent_pid and parent_pid in parent_entries:
                bucket_entries.append(parent_entries[parent_pid])
        for entry in bucket_entries:
            _apply_entry(tracking_set, entry, snapshots, now, report, set_chosen)
        comparison_engine.refresh_set_status(tracking_set)
        tracking_set["updated_at"] = now
        if primary_set_id is None:
            primary_set_id = tracking_set.get("set_id")

    remaining_sets = []
    for tracking_set in list(result.get("tracking_sets") or []):
        if str(tracking_set.get("origin") or "") != models.SET_ORIGIN_ELEMENT_LINKER:
            remaining_sets.append(tracking_set)
            continue
        chosen = chosen_per_set.get(str(tracking_set.get("set_id") or ""), set())
        elements = tracking_set.get("elements") or {}
        removed_any = False
        for persistent_id, record in list(elements.items()):
            if record.get("linker_meta") is None:
                continue
            if persistent_id in chosen:
                continue
            clean = (
                record.get("state") == models.ELEMENT_TRACKED
                and int(record.get("change_count", 0) or 0) == 0
            )
            if clean:
                del elements[persistent_id]
                report["stale_removed"] += 1
                removed_any = True
            else:
                label = ((record.get("metadata") or {}).get("friendly_name")
                         or persistent_id)
                report["stale_kept"].append(str(label))
        if not elements and not list(tracking_set.get("untracked_ids") or []):
            report["sets_removed"].append(str(tracking_set.get("name")))
            continue
        if removed_any:
            comparison_engine.refresh_set_status(tracking_set)
            tracking_set["updated_at"] = now
        remaining_sets.append(tracking_set)
    result["tracking_sets"] = remaining_sets

    if primary_set_id is None:
        family = find_sync_sets(result)
        primary_set_id = family[0].get("set_id") if family else None
    result["updated_at"] = now
    return result, primary_set_id, report


def format_sync_report(report):
    """Human-readable multi-line summary for the output console."""
    report = report or {}
    lines = ["Element Linker sync complete."]
    lines.append(
        "All linker elements monitored: {} child(ren) under {} parent(s) "
        "({} host, {} from linked models).".format(
            int(report.get("children_found", 0) or 0),
            int(report.get("groups", 0) or 0),
            int(report.get("children_host", 0) or 0),
            int(report.get("children_linked", 0) or 0),
        )
    )
    lines.append(
        "Parents resolved: {} host, {} linked-model; {} linker element(s) "
        "skipped (no resolvable parent - not tracked).".format(
            int(report.get("parents_host", 0) or 0),
            int(report.get("parents_linked", 0) or 0),
            int(report.get("no_parent", 0) or 0),
        )
    )
    lines.append(
        "Monitor records: {} added, {} refreshed, {} restored, "
        "{} stale removed.".format(
            int(report.get("added", 0) or 0),
            int(report.get("refreshed", 0) or 0),
            int(report.get("restored", 0) or 0),
            int(report.get("stale_removed", 0) or 0),
        )
    )
    sets_created = list(report.get("sets_created") or [])
    sets_updated = list(report.get("sets_updated") or [])
    sets_removed = list(report.get("sets_removed") or [])
    if sets_created:
        lines.append("Category sets created: {}.".format(", ".join(sets_created)))
    if sets_updated:
        lines.append("Category sets updated: {}.".format(", ".join(sets_updated)))
    if sets_removed:
        lines.append("Empty sync sets removed: {}.".format(", ".join(sets_removed)))
    stale_kept = list(report.get("stale_kept") or [])
    if stale_kept:
        lines.append(
            "Stale records kept (unresolved changes): {}.".format(
                ", ".join(stale_kept)
            )
        )
    for warning in list(report.get("warnings") or []):
        lines.append("Warning: {}".format(warning))
    return "\n".join(lines)
