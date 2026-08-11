# -*- coding: utf-8 -*-
"""Host/link source resolution and category-scoped collection."""

from __future__ import print_function

try:
    from pyrevit import DB
except Exception:
    DB = None

from Snippets import revit_helpers

import models


def _id_value(value):
    return revit_helpers.get_elementid_value(value)


def document_key(document):
    if document is None:
        return ""
    try:
        return "{}|{}".format(document.PathName or "", document.Title or "")
    except Exception:
        return str(document)


def host_source_descriptor(document):
    return {
        "source_type": models.SOURCE_HOST,
        "display_name": "Host - {}".format(getattr(document, "Title", "Current Model")),
        "document_title": str(getattr(document, "Title", "") or ""),
    }


def link_source_descriptor(link_instance):
    link_doc = None
    try:
        link_doc = link_instance.GetLinkDocument()
    except Exception:
        pass
    return {
        "source_type": models.SOURCE_LINK,
        "display_name": str(getattr(link_instance, "Name", "Revit Link") or "Revit Link"),
        "link_instance_unique_id": str(getattr(link_instance, "UniqueId", "") or ""),
        "link_instance_id": _id_value(getattr(link_instance, "Id", None)),
        "linked_document_title": str(getattr(link_doc, "Title", "") or ""),
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
    unique_id = str(descriptor.get("link_instance_unique_id") or "")
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
        "name": str(category.Name or ""),
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
        if str(category.Name or "").strip().lower() not in ALLOWED_CATEGORY_NAMES:
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
    target_name = str(descriptor.get("name") or "").strip().lower()
    if target_name:
        for candidate in list_categories(source_document):
            if str(candidate.get("name") or "").strip().lower() == target_name:
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
    element_unique_id = str(getattr(element, "UniqueId", "") or "")
    if (source_descriptor or {}).get("source_type") == models.SOURCE_LINK:
        return "link:{}:{}".format(
            str((source_descriptor or {}).get("link_instance_unique_id") or ""),
            element_unique_id,
        )
    return "host:{}".format(element_unique_id)


def parse_persistent_id(persistent_key):
    """Split a persistent id into ``(kind, link_instance_unique_id, element_unique_id)``.

    ``kind`` is "host" or "link"; link_instance_unique_id is None for host ids.
    """
    text = str(persistent_key or "")
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
    membership = str(tracking_set.get("membership") or models.MEMBERSHIP_CATEGORY)
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
            key = str(key or "")
            if not key or key in seen:
                continue
            if not str((record or {}).get("parent_persistent_id") or ""):
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
        key = str(key or "")
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

