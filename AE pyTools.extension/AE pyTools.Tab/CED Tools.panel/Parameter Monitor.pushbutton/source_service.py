# -*- coding: utf-8 -*-
"""Host/link source resolution and category-scoped collection."""

from __future__ import print_function

try:
    from pyrevit import DB
except Exception:
    DB = None

from Snippets import revit_helpers

import models
import text_service


def _id_value(value):
    return revit_helpers.get_elementid_value(value)


def document_key(document):
    if document is None:
        return ""
    try:
        return u"{}|{}".format(
            text_service.to_text(document.PathName or "", context=u"Document path"),
            text_service.to_text(document.Title or "", context=u"Document title"),
        )
    except Exception:
        return text_service.to_text(document)


def host_source_descriptor(document):
    return {
        "source_type": models.SOURCE_HOST,
        "display_name": u"Host - {}".format(text_service.to_text(
            getattr(document, "Title", "Current Model") or "Current Model",
            context=u"Host document title",
        )),
        "document_title": text_service.to_text(
            getattr(document, "Title", "") or "", context=u"Host document title"
        ),
    }


def link_source_descriptor(link_instance):
    link_doc = None
    try:
        link_doc = link_instance.GetLinkDocument()
    except Exception:
        pass
    return {
        "source_type": models.SOURCE_LINK,
        "display_name": text_service.to_text(
            getattr(link_instance, "Name", "Revit Link") or "Revit Link",
            context=u"Revit link instance name",
        ),
        "link_instance_unique_id": text_service.to_text(
            getattr(link_instance, "UniqueId", "") or "",
            context=u"Revit link instance unique id",
        ),
        "link_instance_id": _id_value(getattr(link_instance, "Id", None)),
        "linked_document_title": text_service.to_text(
            getattr(link_doc, "Title", "") or "", context=u"Linked document title"
        ),
        "loaded": link_doc is not None,
    }


def list_sources(host_document, include_unavailable=False):
    if DB is None or host_document is None:
        return []
    sources = [host_source_descriptor(host_document)]
    links = list(DB.FilteredElementCollector(host_document).OfClass(DB.RevitLinkInstance).ToElements())
    for link_instance in links:
        descriptor = link_source_descriptor(link_instance)
        if descriptor.get("loaded") or include_unavailable:
            sources.append(descriptor)
    sources[1:] = sorted(sources[1:], key=lambda item: item.get("display_name", "").lower())
    return sources


def _link_transform_state(link_instance):
    if link_instance is None:
        return {}
    transform = None
    for method_name in ("GetTotalTransform", "GetTransform"):
        try:
            transform = getattr(link_instance, method_name)()
            if transform is not None:
                break
        except Exception:
            transform = None
    if transform is None:
        return {}
    vectors = [transform.Origin, transform.BasisX, transform.BasisY, transform.BasisZ]
    matrix = []
    for vector in vectors:
        matrix.extend([float(vector.X), float(vector.Y), float(vector.Z)])
    return {"matrix": matrix}


def _workset_details(workset):
    """Return JSON-safe workset identity/open-state details when available."""
    if workset is None:
        return {}
    try:
        is_open = bool(workset.IsOpen)
    except Exception:
        # Do not guess.  This can be absent in a test double or unavailable
        # for a non-workshared document.
        return {}
    return {
        "id": _id_value(getattr(workset, "Id", None)),
        "name": text_service.to_text(
            getattr(workset, "Name", "") or "Workset", context=u"Workset name"
        ),
        "is_open": is_open,
    }


def element_workset_details(document, element):
    """Return the workset containing ``element`` as JSON-safe metadata.

    Each scanned linked element stores this identity.  On later scans we can
    distinguish a real removal from an element that Revit has not expanded
    because its linked-model workset is closed.
    """
    if document is None or element is None:
        return {}
    try:
        workset_id = element.WorksetId
        workset = document.GetWorksetTable().GetWorkset(workset_id)
    except Exception:
        return {}
    return _workset_details(workset)


def _user_worksets(document):
    """Return the readable user worksets in ``document``.

    System/view worksets cannot be selectively closed in the way that causes
    source element collection to become incomplete, so only user worksets are
    relevant to monitor scans.
    """
    if DB is None or document is None:
        return []
    try:
        collector = DB.FilteredWorksetCollector(document)
        worksets = collector.OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
    except Exception:
        return []
    details = []
    for workset in worksets:
        detail = _workset_details(workset)
        if detail:
            details.append(detail)
    return details


def _closed_user_worksets(document):
    return [
        item for item in _user_worksets(document)
        if item.get("is_open") is False
    ]


def _workset_names(worksets):
    names = [
        text_service.to_text(item.get("name") or "Workset")
        for item in list(worksets or [])
    ]
    return ", ".join(names)


def _record_workset_id(record):
    record = record or {}
    metadata = (
        record.get("current_metadata")
        or record.get("metadata")
        or {}
    )
    value = metadata.get("workset_id")
    if value is None or text_service.to_text(value) == "":
        return None
    try:
        return int(value)
    except Exception:
        return text_service.to_text(value)


def _link_instance_workset_message(host_document, link_instance):
    details = element_workset_details(host_document, link_instance)
    if details and details.get("is_open") is False:
        return (
            "The configured Revit link is in closed host workset '{}'. "
            "Open the workset before scanning."
        ).format(details.get("name") or "Workset")
    return ""


def resolve_source(host_document, descriptor):
    descriptor = descriptor or {}
    if descriptor.get("source_type") == models.SOURCE_HOST:
        return {
            "available": host_document is not None,
            "source_document": host_document,
            "link_instance": None,
            "source_state": {},
            "message": "",
        }
    unique_id = text_service.to_text(descriptor.get("link_instance_unique_id") or "")
    link_instance = None
    if host_document is not None and unique_id:
        try:
            link_instance = host_document.GetElement(unique_id)
        except Exception:
            link_instance = None
    if link_instance is None and host_document is not None:
        target_id = int(descriptor.get("link_instance_id") or 0)
        for candidate in DB.FilteredElementCollector(host_document).OfClass(DB.RevitLinkInstance):
            if _id_value(candidate.Id) == target_id:
                link_instance = candidate
                break
    if link_instance is None:
        return {
            "available": False,
            "source_document": None,
            "link_instance": None,
            "source_state": {},
            "message": "The configured Revit link instance no longer exists.",
        }
    workset_message = _link_instance_workset_message(host_document, link_instance)
    if workset_message:
        return {
            "available": False,
            "source_document": None,
            "link_instance": link_instance,
            "source_state": _link_transform_state(link_instance),
            "message": workset_message,
        }
    try:
        link_document = link_instance.GetLinkDocument()
    except Exception:
        link_document = None
    if link_document is None:
        return {
            "available": False,
            "source_document": None,
            "link_instance": link_instance,
            "source_state": _link_transform_state(link_instance),
            "message": "The configured Revit link is unloaded or unavailable.",
        }
    return {
        "available": True,
        "source_document": link_document,
        "link_instance": link_instance,
        "source_state": _link_transform_state(link_instance),
        "message": "",
    }


def evaluate_source_for_scan(host_document, tracking_set, require_complete=False):
    """Resolve a set source and reject incomplete linked-model collections.

    A Revit link can be present while closed linked-model worksets leave its
    category collector incomplete.  Comparing that partial collector against
    the stored baseline would incorrectly mark elements as removed.  This
    gate reports a whole-set link status instead and leaves record states
    untouched.  Older sets without stored workset metadata are handled
    conservatively whenever the linked document has closed user worksets.
    """
    tracking_set = tracking_set or {}
    source = tracking_set.get("source") or {}
    result = resolve_source(host_document, source)
    result["set_status"] = None
    result["closed_worksets"] = []

    if not result.get("available"):
        if source.get("source_type") == models.SOURCE_LINK:
            result["set_status"] = models.SET_LINK_UNAVAILABLE
        else:
            result["set_status"] = models.SET_SOURCE_UNAVAILABLE
        return result

    if source.get("source_type") != models.SOURCE_LINK:
        return result

    closed_worksets = _closed_user_worksets(result.get("source_document"))
    result["closed_worksets"] = closed_worksets
    if not closed_worksets:
        return result

    records = [
        record for record in list((tracking_set.get("elements") or {}).values())
        if (record or {}).get("state") != models.ELEMENT_REMOVED
    ]
    closed_ids = set([item.get("id") for item in closed_worksets])
    known_ids = set()
    unknown_records = 0
    for record in records:
        workset_id = _record_workset_id(record)
        if workset_id is None:
            unknown_records += 1
        else:
            known_ids.add(workset_id)
    affected_ids = known_ids.intersection(closed_ids)

    # A new baseline must include every source workset.  For existing sets,
    # unknown legacy metadata is intentionally treated as unsafe rather than
    # risking a false removal on the first scan after this protection ships.
    incomplete = bool(require_complete or affected_ids or unknown_records)
    if not incomplete:
        return result

    names = _workset_names(closed_worksets)
    if require_complete:
        message = (
            "The linked model has closed workset(s): {}. Open them before "
            "creating a complete baseline."
        ).format(names)
    elif unknown_records:
        message = (
            "Scan skipped: the linked model has closed workset(s): {}, and "
            "{} tracked element(s) do not yet have workset metadata. Existing "
            "element states were retained."
        ).format(names, unknown_records)
    else:
        message = (
            "Scan skipped: tracked element(s) are in closed linked-model "
            "workset(s): {}. Existing element states were retained."
        ).format(names)
    result["available"] = False
    result["set_status"] = models.SET_LINK_WORKSETS_UNAVAILABLE
    result["message"] = message
    return result


# Categories users may build tracking sets from. Matched case-insensitively
# against Revit's category names; "Fire Alarm Devices" covers the category
# Revit actually names that way.
ALLOWED_CATEGORY_NAMES = set([
    "mechanical control devices",
    "duct accessories",
    "mechanical equipment",
    "plumbing fixtures",
    "plumbing equipment",
    "pipe accessories",
    "electrical equipment",
    "electrical fixtures",
    "lighting fixtures",
    "lighting devices",
    "data devices",
    "security devices",
    "fire alarm devices",
    "fire alarm fixtures",
    "specialty equipment",
    "generic models",
])


def category_descriptor(category):
    value = _id_value(category.Id)
    return {
        "id": value,
        "builtin_id": value if value < 0 else None,
        "name": text_service.to_text(category.Name or "", context=u"Category name"),
    }


def list_categories(source_document):
    if DB is None or source_document is None:
        return []
    categories = {}
    collector = DB.FilteredElementCollector(source_document).WhereElementIsNotElementType()
    for element in collector:
        category = getattr(element, "Category", None)
        if category is None:
            continue
        try:
            if category.CategoryType != DB.CategoryType.Model:
                continue
        except Exception:
            pass
        value = _id_value(category.Id)
        if value == 0:
            continue
        if text_service.to_text(category.Name or "").strip().lower() not in ALLOWED_CATEGORY_NAMES:
            continue
        categories[value] = category_descriptor(category)
    return sorted(categories.values(), key=lambda item: item.get("name", "").lower())


def resolve_category(source_document, descriptor):
    descriptor = descriptor or {}
    target_id = descriptor.get("id")
    if target_id is not None:
        try:
            category = source_document.Settings.Categories.get_Item(
                revit_helpers.elementid_from_value(target_id)
            )
            if category is not None:
                return category_descriptor(category)
        except Exception:
            pass
    target_name = text_service.to_text(descriptor.get("name") or "").strip().lower()
    if target_name:
        for candidate in list_categories(source_document):
            if text_service.to_text(candidate.get("name") or "").strip().lower() == target_name:
                return candidate
    return None


def collect_elements(source_document, category):
    if DB is None or source_document is None:
        return []
    category_id = (category or {}).get("id")
    if category_id is None:
        return []
    element_filter = DB.ElementCategoryFilter(revit_helpers.elementid_from_value(category_id))
    collector = DB.FilteredElementCollector(source_document).WherePasses(element_filter)
    return list(collector.WhereElementIsNotElementType().ToElements())


def persistent_id(source_descriptor, element):
    element_unique_id = text_service.to_text(
        getattr(element, "UniqueId", "") or "", context=u"Element unique id"
    )
    if (source_descriptor or {}).get("source_type") == models.SOURCE_LINK:
        return "link:{}:{}".format(
            text_service.to_text(
                (source_descriptor or {}).get("link_instance_unique_id") or ""
            ),
            element_unique_id,
        )
    return "host:{}".format(element_unique_id)


def parse_persistent_id(persistent_key):
    """Split a persistent id into ``(kind, link_instance_unique_id, element_unique_id)``.

    ``kind`` is "host" or "link"; link_instance_unique_id is None for host ids.
    """
    text = text_service.to_text(persistent_key or "")
    if text.startswith("link:"):
        parts = text.split(":", 2)
        if len(parts) == 3:
            return "link", parts[1], parts[2]
        return "link", "", parts[-1]
    if text.startswith("host:"):
        return "host", None, text[len("host:"):]
    return "host", None, text


def unique_id_from_persistent_id(persistent_key):
    """Return the element UniqueId encoded in a persistent id ("host:..." / "link:...:...")."""
    return parse_persistent_id(persistent_key)[2]


def link_total_transform(link_instance):
    """Total transform of a link instance, or None."""
    if link_instance is None:
        return None
    for method_name in ("GetTotalTransform", "GetTransform"):
        try:
            transform = getattr(link_instance, method_name)()
            if transform is not None:
                return transform
        except Exception:
            pass
    return None


def resolve_member(host_document, persistent_key):
    """Resolve an explicit-membership persistent id to a live element.

    Returns ``(element, element_document, location_transform)``. Host members
    resolve directly; ``link:`` members resolve through their link instance,
    with the link's total transform for host/world location conversion.
    """
    kind, link_unique_id, element_unique_id = parse_persistent_id(persistent_key)
    if host_document is None or not element_unique_id:
        return None, None, None
    if kind == "host":
        try:
            element = host_document.GetElement(element_unique_id)
        except Exception:
            element = None
        return element, host_document, None
    link_instance = None
    if link_unique_id:
        try:
            link_instance = host_document.GetElement(link_unique_id)
        except Exception:
            link_instance = None
    link_document = None
    if link_instance is not None:
        try:
            link_document = link_instance.GetLinkDocument()
        except Exception:
            link_document = None
    if link_document is None:
        return None, None, None
    try:
        element = link_document.GetElement(element_unique_id)
    except Exception:
        element = None
    return element, link_document, link_total_transform(link_instance)


def collect_set_member_pairs(source_document, tracking_set):
    """Collect ``(persistent_id, element, element_document, location_transform)``
    tuples for a tracking set.

    Category-membership sets keep the original category sweep in the source
    document. Explicit sets resolve their recorded members (tracked +
    untracked) by persistent id; for explicit sets ``source_document`` must
    be the host document so ``link:`` members can resolve through their link
    instance.
    """
    tracking_set = tracking_set or {}
    membership = text_service.to_text(
        tracking_set.get("membership") or models.MEMBERSHIP_CATEGORY
    )
    if membership != models.MEMBERSHIP_EXPLICIT:
        pairs = []
        seen = set()
        for element in collect_elements(source_document, tracking_set.get("category") or {}):
            key = persistent_id(tracking_set.get("source") or {}, element)
            pairs.append((key, element, source_document, None))
            seen.add(key)
        # Manually linked children (Add Device) can be any category, so a
        # category sweep alone would drop them on every scan — resolve
        # parent-linked records explicitly.
        for key, record in list((tracking_set.get("elements") or {}).items()):
            key = text_service.to_text(key or "")
            if not key or key in seen:
                continue
            if not text_service.to_text((record or {}).get("parent_persistent_id") or ""):
                continue
            element, element_document, transform = resolve_member(source_document, key)
            if element is None:
                continue
            if DB is not None and isinstance(element, DB.ElementType):
                continue
            pairs.append((key, element, element_document, transform))
        return pairs
    if source_document is None:
        return []
    keys = list((tracking_set.get("elements") or {}).keys())
    keys.extend(list(tracking_set.get("untracked_ids") or []))
    pairs = []
    seen = set()
    for key in keys:
        key = text_service.to_text(key or "")
        if not key or key in seen:
            continue
        seen.add(key)
        element, element_document, transform = resolve_member(source_document, key)
        if element is None:
            continue
        if DB is not None and isinstance(element, DB.ElementType):
            continue
        pairs.append((key, element, element_document, transform))
    return pairs


def collect_set_elements(source_document, tracking_set):
    """Live elements a tracking set covers (element objects only)."""
    return [pair[1] for pair in collect_set_member_pairs(source_document, tracking_set)]

