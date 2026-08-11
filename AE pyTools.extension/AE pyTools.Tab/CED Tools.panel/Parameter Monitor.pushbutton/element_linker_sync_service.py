# -*- coding: utf-8 -*-
"""Revit-facing Element Linker sync: collect and register children + parents.

Runs inside the ExternalEvent handler (API context). Every host element with
an Element_Linker payload and a live host parent is registered for
monitoring, together with its parent. The membership decision + store
mutation lives in the pure ``sync_logic``.
"""

from __future__ import print_function

try:
    from pyrevit import DB, forms
except Exception:
    DB = None
    forms = None

from Snippets import revit_helpers

import mep_linker_bridge
import models
import source_service
import sync_logic
import tracking_service

TITLE = "Parameter Monitor"


def _id_value(value):
    return revit_helpers.get_elementid_value(value)


def _element_family_name(element):
    """Family name for FamilyInstance/Group; mirrors the MEPRFP host_name
    validation so linked-doc ElementId collisions are rejected."""
    if element is None or DB is None:
        return ""
    if isinstance(element, DB.FamilyInstance):
        symbol = getattr(element, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol is not None else None
        return str(getattr(family, "Name", "") or "")
    if isinstance(element, DB.Group):
        group_type = getattr(element, "GroupType", None)
        return str(getattr(group_type, "Name", "") or "")
    return ""


def _link_contexts(document):
    """Loaded link instances with their documents and total transforms."""
    contexts = []
    if DB is None:
        return contexts
    for link_instance in DB.FilteredElementCollector(document).OfClass(DB.RevitLinkInstance):
        try:
            link_document = link_instance.GetLinkDocument()
        except Exception:
            link_document = None
        if link_document is None:
            continue
        contexts.append({
            "unique_id": str(getattr(link_instance, "UniqueId", "") or ""),
            "document": link_document,
            "transform": source_service.link_total_transform(link_instance),
        })
    return contexts


def _resolve_parent(document, link_contexts, parent_element_id, host_name):
    """Find the linker parent in the host doc or any loaded linked doc.

    ElementIds are not globally unique across documents; when the linker
    carries the parent's family name (``host_name``) it validates which
    candidate is the real parent. Without it, host wins, then first link.
    Returns a dict {persistent_id, element, document, transform, where}
    or None.
    """
    try:
        element_id = revit_helpers.elementid_from_value(int(parent_element_id))
    except Exception:
        return None
    candidates = []
    try:
        host_element = document.GetElement(element_id)
    except Exception:
        host_element = None
    if host_element is not None:
        unique_id = str(getattr(host_element, "UniqueId", "") or "")
        if unique_id:
            candidates.append({
                "persistent_id": "host:{}".format(unique_id),
                "element": host_element,
                "document": document,
                "transform": None,
                "where": "host",
            })
    for context in link_contexts:
        try:
            link_element = context["document"].GetElement(element_id)
        except Exception:
            link_element = None
        if link_element is None:
            continue
        unique_id = str(getattr(link_element, "UniqueId", "") or "")
        if not unique_id or not context["unique_id"]:
            continue
        candidates.append({
            "persistent_id": "link:{}:{}".format(context["unique_id"], unique_id),
            "element": link_element,
            "document": context["document"],
            "transform": context["transform"],
            "where": "link",
        })
    if not candidates:
        return None
    target = str(host_name or "").strip().lower()
    if target:
        for candidate in candidates:
            if _element_family_name(candidate["element"]).strip().lower() == target:
                return candidate
    return candidates[0]


def collect_linked_children(document, directive_index):
    """All host elements with an Element_Linker payload.

    Parents are resolved in the host document AND every loaded linked
    document; children whose parent cannot be resolved are still returned
    (tracked standalone). Returns ``(children, targets, counts)`` where
    ``targets`` maps persistent_id -> {element, document, transform} for
    snapshotting children and parents.
    """
    children = []
    targets = {}
    counts = {"with_linker": 0, "parents_host": 0, "parents_linked": 0,
              "no_parent": 0}
    link_contexts = _link_contexts(document)
    seen_parent_pids = set()
    collected = []
    for element_class in (DB.FamilyInstance, DB.Group):
        collector = DB.FilteredElementCollector(document).OfClass(element_class)
        collected.extend(list(collector.WhereElementIsNotElementType().ToElements()))
    for element in collected:
        try:
            linker = mep_linker_bridge.read_linker(element)
        except mep_linker_bridge.MepBridgeError:
            raise
        except Exception:
            linker = None
        if linker is None:
            continue
        unique_id = str(getattr(element, "UniqueId", "") or "")
        if not unique_id:
            continue
        counts["with_linker"] += 1
        parent = None
        parent_element_id = linker.get("parent_element_id")
        if parent_element_id is not None:
            parent = _resolve_parent(
                document,
                link_contexts,
                parent_element_id,
                linker.get("host_name"),
            )
        if parent is None:
            counts["no_parent"] += 1
        elif parent["where"] == "link":
            counts["parents_linked"] += 1
        else:
            counts["parents_host"] += 1
        led_id = str(linker.get("led_id") or "")
        children.append({
            "unique_id": unique_id,
            "element_id": _id_value(getattr(element, "Id", None)),
            "name": str(getattr(element, "Name", "") or ""),
            "parent_persistent_id": parent["persistent_id"] if parent else None,
            "parent_unique_id": (
                str(getattr(parent["element"], "UniqueId", "") or "")
                if parent else ""
            ),
            "parent_name": (
                str(getattr(parent["element"], "Name", "") or "Parent")
                if parent else ""
            ),
            "led_id": led_id,
            "set_id": str(linker.get("set_id") or ""),
            "space_profile_id": str(linker.get("space_profile_id") or ""),
            "has_directives": bool(directive_index.get(led_id, False)),
        })
        targets["host:{}".format(unique_id)] = {
            "element": element,
            "document": document,
            "transform": None,
        }
        if parent is not None and parent["persistent_id"] not in seen_parent_pids:
            seen_parent_pids.add(parent["persistent_id"])
            targets[parent["persistent_id"]] = {
                "element": parent["element"],
                "document": parent["document"],
                "transform": parent["transform"],
            }
    return children, targets, counts


def run_sync(document, uidocument, store, logger=None):
    """Full Sync Element Linker operation: monitor every linker-populated
    element, plus its parent when one resolves (host or linked document).
    Returns ``(store, set_id, message)`` or None when there is nothing to
    register."""
    if not mep_linker_bridge.is_available():
        raise ValueError(
            "The MEPRFP Automation 2.0 lib folder was not found, so "
            "Element_Linker payloads cannot be read."
        )
    profile_data, profile_warning = mep_linker_bridge.load_profile_data(document)
    directive_index = sync_logic.led_directive_index(profile_data)

    children, targets, counts = collect_linked_children(document, directive_index)
    if not children:
        forms.alert(
            "No host elements with a populated Element_Linker parameter "
            "were found.",
            title=TITLE,
        )
        return None

    groups = sync_logic.group_children(children)
    report = {
        "children_found": len(children),
        "groups": len(groups),
        "parents_host": counts.get("parents_host", 0),
        "parents_linked": counts.get("parents_linked", 0),
        "no_parent": counts.get("no_parent", 0),
        "warnings": [],
    }
    if profile_warning:
        report["warnings"].append(profile_warning)

    host_descriptor = source_service.host_source_descriptor(document)
    # Snapshot with the union of parameters tracked across the category
    # sync sets so existing per-set parameter tracking keeps its data.
    descriptors = []
    descriptor_keys = set()
    for existing_set in sync_logic.find_sync_sets(store):
        for descriptor in list(existing_set.get("tracked_properties") or []):
            key = str((descriptor or {}).get("key") or "")
            if key and key not in descriptor_keys:
                descriptor_keys.add(key)
                descriptors.append(descriptor)

    entries = []
    snapshots = {}
    parent_entry_pids = set()
    now = models.utc_now_text()
    for child in children:
        child_pid = "host:{}".format(child["unique_id"])
        parent_pid = child.get("parent_persistent_id")
        entries.append({
            "persistent_id": child_pid,
            "role": "child",
            "parent_persistent_id": parent_pid,
            "linker_meta": {
                "led_id": child.get("led_id"),
                "set_id": child.get("set_id"),
                "space_profile_id": child.get("space_profile_id"),
                "has_directives": bool(child.get("has_directives")),
            },
        })
        if parent_pid and parent_pid not in parent_entry_pids:
            parent_entry_pids.add(parent_pid)
            entries.append({
                "persistent_id": parent_pid,
                "role": "parent",
                "parent_persistent_id": None,
                "linker_meta": None,
            })

    type_cache = {}
    for persistent_id, target in targets.items():
        element = (target or {}).get("element")
        if element is None:
            continue
        snapshots[persistent_id] = tracking_service._snapshot_element(
            document,
            target.get("document") or document,
            host_descriptor,
            element,
            descriptors,
            type_cache,
            True,
            persistent_id_override=persistent_id,
            location_transform=target.get("transform"),
        )

    updated_store, sync_set_id, membership_report = sync_logic.apply_sync_membership(
        store, entries, snapshots, host_descriptor, now_text=now
    )
    earlier_warnings = list(report.get("warnings") or [])
    report.update(membership_report)
    report["warnings"] = earlier_warnings + list(membership_report.get("warnings") or [])
    if logger is not None:
        try:
            logger.info(
                "Parameter Monitor element-linker sync: %s children, %s parents, "
                "%s added, %s refreshed",
                report.get("children_found"),
                report.get("groups"),
                report.get("added"),
                report.get("refreshed"),
            )
        except Exception:
            pass
    return updated_store, sync_set_id, sync_logic.format_sync_report(report)
