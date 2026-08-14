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

import element_linker_codec
import location_service
import models
import relationship_service
import source_service
import sync_logic
import text_service
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
        return text_service.to_text(
            getattr(family, "Name", "") or "", context=u"Element Linker family name"
        )
    if isinstance(element, DB.Group):
        group_type = getattr(element, "GroupType", None)
        return text_service.to_text(
            getattr(group_type, "Name", "") or "", context=u"Element Linker group type name"
        )
    return ""


def _document_contexts(document):
    """The host document plus every loaded link, with pid prefixes and
    total transforms for host/world location conversion."""
    contexts = [{
        "kind": "host",
        "link_uid": None,
        "document": document,
        "transform": None,
    }]
    if DB is None:
        return contexts
    for link_instance in DB.FilteredElementCollector(document).OfClass(DB.RevitLinkInstance):
        try:
            link_document = link_instance.GetLinkDocument()
        except Exception:
            link_document = None
        if link_document is None:
            continue
        link_uid = text_service.to_text(
            getattr(link_instance, "UniqueId", ""), context=u"Element Linker link unique id"
        )
        if not link_uid:
            continue
        contexts.append({
            "kind": "link",
            "link_uid": link_uid,
            "document": link_document,
            "transform": source_service.link_total_transform(link_instance),
        })
    return contexts


def _context_pid(context, element):
    unique_id = text_service.to_text(
        getattr(element, "UniqueId", "") or "", context=u"Element Linker element unique id"
    )
    if not unique_id:
        return None
    if context["kind"] == "link":
        return "link:{}:{}".format(context["link_uid"], unique_id)
    return "host:{}".format(unique_id)


def _world_point(element, transform):
    location = location_service.read_location(element, transform=transform)
    if location.get("state") != models.VALUE_VALID:
        return None
    return (
        float(location.get("x", 0.0)),
        float(location.get("y", 0.0)),
        float(location.get("z", 0.0)),
    )


def _resolve_parent(child_context, contexts, linker):
    """Find the linker parent for a child, validating against ElementId
    collisions.

    The linker's ``parent_element_id`` belongs to the id space of the
    document the placement ran in — always the child's own document — so
    that document is authoritative. Host children additionally search the
    loaded links (legacy payloads written before linked equipment moved).
    Cross-document candidates are validated: they must be a FamilyInstance
    or Group, prefer a ``host_name`` family match, and prefer the candidate
    closest to the linker's stored world parent location. Returns
    {persistent_id, element, document, transform, where} or None.
    """
    parent_element_id = linker.get("parent_element_id")
    if parent_element_id is None:
        return None
    try:
        element_id = revit_helpers.elementid_from_value(int(parent_element_id))
    except Exception:
        return None

    def _candidate(context):
        try:
            element = context["document"].GetElement(element_id)
        except Exception:
            element = None
        if element is None:
            return None
        pid = _context_pid(context, element)
        if pid is None:
            return None
        return {
            "persistent_id": pid,
            "element": element,
            "document": context["document"],
            "transform": context["transform"],
            "where": context["kind"],
        }

    # A link child's linker was written while its model was open, so its
    # own id space is authoritative. A HOST child's parent id may belong
    # to a linked document (linked equipment), so the host lookup gets no
    # priority — every document is a candidate, validated below.
    if child_context["kind"] == "link":
        return _candidate(child_context)

    candidates = []
    for context in contexts:
        candidate = _candidate(context)
        if candidate is None:
            continue
        if DB is not None and not isinstance(
            candidate["element"], (DB.FamilyInstance, DB.Group)
        ):
            continue
        candidates.append(candidate)
    if not candidates:
        return None
    target = text_service.to_text(linker.get("host_name") or "").strip().lower()
    if target:
        matches = [
            candidate for candidate in candidates
            if _element_family_name(candidate["element"]).strip().lower() == target
        ]
        if matches:
            candidates = matches
    if len(candidates) > 1:
        stored = linker.get("parent_location_ft")
        if isinstance(stored, (list, tuple)) and len(stored) == 3:
            def _distance(candidate):
                point = _world_point(candidate["element"], candidate["transform"])
                if point is None:
                    return 1.0e12
                dx = point[0] - float(stored[0])
                dy = point[1] - float(stored[1])
                dz = point[2] - float(stored[2])
                return (dx * dx) + (dy * dy) + (dz * dz)
            candidates.sort(key=_distance)
    return candidates[0]


def collect_linked_children(document):
    """Every element with an Element_Linker payload — in the host document
    AND in every loaded linked document.

    Returns ``(children, targets, counts, warnings)`` where ``targets``
    maps persistent_id -> {element, document, transform} for snapshotting
    children and parents. Payloads are read straight off the elements'
    Element_Linker parameter via the self-contained codec.
    """
    children = []
    targets = {}
    counts = {"with_linker": 0, "children_host": 0, "children_linked": 0,
              "parents_host": 0, "parents_linked": 0, "no_parent": 0}
    warnings = []
    contexts = _document_contexts(document)
    seen_parent_pids = set()
    for context in contexts:
        collected = []
        for element_class in (DB.FamilyInstance, DB.Group):
            collector = DB.FilteredElementCollector(
                context["document"]
            ).OfClass(element_class)
            collected.extend(
                list(collector.WhereElementIsNotElementType().ToElements())
            )
        for element in collected:
            try:
                linker = element_linker_codec.read_linker(element)
            except Exception:
                linker = None
            if linker is None:
                continue
            child_pid = _context_pid(context, element)
            if child_pid is None:
                continue
            counts["with_linker"] += 1
            parent = _resolve_parent(context, contexts, linker)
            if parent is None:
                # No resolvable parent means no parent/child relationship to
                # monitor — these elements are not tracked at all.
                counts["no_parent"] += 1
                continue
            if context["kind"] == "link":
                counts["children_linked"] += 1
            else:
                counts["children_host"] += 1
            if parent["where"] == "link":
                counts["parents_linked"] += 1
            else:
                counts["parents_host"] += 1
            led_id = text_service.to_text(linker.get("led_id") or "")
            children.append({
                "persistent_id": child_pid,
                "unique_id": text_service.to_text(
                    getattr(element, "UniqueId", "") or "", context=u"Element Linker child unique id"
                ),
                "element_id": _id_value(getattr(element, "Id", None)),
                "name": text_service.to_text(
                    getattr(element, "Name", "") or "", context=u"Element Linker child name"
                ),
                "parent_persistent_id": parent["persistent_id"],
                "parent_unique_id": text_service.to_text(
                    getattr(parent["element"], "UniqueId", "") or "",
                    context=u"Element Linker parent unique id",
                ),
                "parent_name": text_service.to_text(
                    getattr(parent["element"], "Name", "") or "Parent",
                    context=u"Element Linker parent name",
                ),
                "led_id": led_id,
                "set_id": text_service.to_text(linker.get("set_id") or ""),
                "space_profile_id": text_service.to_text(linker.get("space_profile_id") or ""),
            })
            targets[child_pid] = {
                "element": element,
                "document": context["document"],
                "transform": context["transform"],
            }
            if parent is not None and parent["persistent_id"] not in seen_parent_pids:
                seen_parent_pids.add(parent["persistent_id"])
                targets[parent["persistent_id"]] = {
                    "element": parent["element"],
                    "document": parent["document"],
                    "transform": parent["transform"],
                }
    return children, targets, counts, warnings


def run_sync(document, uidocument, store, logger=None):
    """Full Sync Element Linker operation: monitor every linker-populated
    element, plus its parent when one resolves (host or linked document).
    Returns ``(store, set_id, message)`` or None when there is nothing to
    register."""
    children, targets, counts, collect_warnings = collect_linked_children(document)
    if not children:
        forms.alert(
            "No elements with a populated Element_Linker parameter were "
            "found in the host model or any loaded link.",
            title=TITLE,
        )
        return None

    groups = sync_logic.group_children(children)
    report = {
        "children_found": len(children),
        "children_host": counts.get("children_host", 0),
        "children_linked": counts.get("children_linked", 0),
        "groups": len(groups),
        "parents_host": counts.get("parents_host", 0),
        "parents_linked": counts.get("parents_linked", 0),
        "no_parent": counts.get("no_parent", 0),
        "warnings": list(collect_warnings or []),
    }

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
        child_pid = child["persistent_id"]
        parent_pid = child.get("parent_persistent_id")
        entries.append({
            "persistent_id": child_pid,
            "role": "child",
            "parent_persistent_id": parent_pid,
            "linker_meta": {
                "led_id": child.get("led_id"),
                "set_id": child.get("set_id"),
                "space_profile_id": child.get("space_profile_id"),
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

    host_child_pids = set([
        child["persistent_id"] for child in children
        if str(child.get("persistent_id") or "").startswith("host:")
    ])
    type_cache = {}
    for persistent_id, target in targets.items():
        element = (target or {}).get("element")
        if element is None:
            continue
        # Host children carry a self-pointing device relationship so their
        # circuit context resolves now and refreshes on every scan (this is
        # what enables Select Circuit for Profile children). Linked-model
        # children can't: their circuits live in the link document.
        relationship = None
        if persistent_id in host_child_pids:
            relationship = relationship_service.relationship_from_device(element)
        snapshots[persistent_id] = tracking_service._snapshot_element(
            document,
            target.get("document") or document,
            host_descriptor,
            element,
            descriptors,
            type_cache,
            True,
            relationship=relationship,
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
