# -*- coding: utf-8 -*-
"""Selection and Show-in-Model operations with linked-model fallback."""

from __future__ import print_function

try:
    from pyrevit import DB
    from System.Collections.Generic import List
except Exception:
    DB = None
    List = None

from Snippets import revit_helpers

import models
import source_service


def _host_element(document, record):
    unique_id = str((record or {}).get("source_element_unique_id") or "")
    if not unique_id:
        return None
    try:
        return document.GetElement(unique_id)
    except Exception:
        return None


def _linked_parts(host_document, tracking_set, record):
    resolved = source_service.resolve_source(host_document, tracking_set.get("source") or {})
    if not resolved.get("available"):
        return None, None
    link_instance = resolved.get("link_instance")
    link_document = resolved.get("source_document")
    unique_id = str((record or {}).get("source_element_unique_id") or "")
    linked_element = None
    if link_document is not None and unique_id:
        try:
            linked_element = link_document.GetElement(unique_id)
        except Exception:
            linked_element = None
    return link_instance, linked_element


def select_tracked(uidocument, tracking_set, record):
    return select_tracked_many(uidocument, tracking_set, [record])


def _available_records(records):
    return [
        record for record in list(records or [])
        if record is not None and record.get("state") != models.ELEMENT_REMOVED
    ]


def select_tracked_many(uidocument, tracking_set, records):
    records = _available_records(records)
    if uidocument is None or not records:
        return False
    document = uidocument.Document
    if (tracking_set.get("source") or {}).get("source_type") == models.SOURCE_HOST:
        ids = [
            element.Id for element in [_host_element(document, record) for record in records]
            if element is not None
        ]
        if not ids:
            return False
        uidocument.Selection.SetElementIds(List[DB.ElementId](ids))
        return True

    resolved = source_service.resolve_source(document, tracking_set.get("source") or {})
    link_instance = resolved.get("link_instance") if resolved.get("available") else None
    link_document = resolved.get("source_document") if resolved.get("available") else None
    if link_instance is None or link_document is None:
        return False
    linked_elements = []
    for record in records:
        unique_id = str(record.get("source_element_unique_id") or "")
        try:
            linked_element = link_document.GetElement(unique_id) if unique_id else None
        except Exception:
            linked_element = None
        if linked_element is not None:
            linked_elements.append(linked_element)
    if not linked_elements:
        return False
    try:
        references = List[DB.Reference]([
            DB.Reference(element).CreateLinkReference(link_instance)
            for element in linked_elements
        ])
        setter = getattr(uidocument.Selection, "SetReferences", None)
        if setter is not None:
            setter(references)
            return True
    except Exception:
        pass
    uidocument.Selection.SetElementIds(List[DB.ElementId]([link_instance.Id]))
    return True


def show_tracked(uidocument, tracking_set, record):
    return show_tracked_many(uidocument, tracking_set, [record])


def show_tracked_many(uidocument, tracking_set, records):
    records = _available_records(records)
    if uidocument is None or not records:
        return False
    document = uidocument.Document
    if (tracking_set.get("source") or {}).get("source_type") == models.SOURCE_HOST:
        ids = [
            element.Id for element in [_host_element(document, record) for record in records]
            if element is not None
        ]
        if not ids:
            return False
        uidocument.ShowElements(List[DB.ElementId](ids))
        return True
    link_instance, _linked_element = _linked_parts(document, tracking_set, records[0])
    if link_instance is None:
        return False
    uidocument.ShowElements(List[DB.ElementId]([link_instance.Id]))
    return True


def select_host_unique_id(uidocument, unique_id):
    if uidocument is None or not unique_id:
        return False
    element = uidocument.Document.GetElement(str(unique_id))
    if element is None:
        return False
    uidocument.Selection.SetElementIds(List[DB.ElementId]([element.Id]))
    return True


def show_host_unique_id(uidocument, unique_id):
    if uidocument is None or not unique_id:
        return False
    element = uidocument.Document.GetElement(str(unique_id))
    if element is None:
        return False
    uidocument.ShowElements(List[DB.ElementId]([element.Id]))
    return True


def select_circuit_id(uidocument, circuit_id):
    if uidocument is None or int(circuit_id or 0) <= 0:
        return False
    element_id = revit_helpers.elementid_from_value(circuit_id)
    if uidocument.Document.GetElement(element_id) is None:
        return False
    uidocument.Selection.SetElementIds(List[DB.ElementId]([element_id]))
    return True
