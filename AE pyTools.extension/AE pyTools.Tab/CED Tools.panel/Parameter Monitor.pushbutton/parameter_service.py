# -*- coding: utf-8 -*-
"""Parameter discovery, durable identity, and normalized reads."""

from __future__ import print_function

try:
    import System
    from pyrevit import DB
except Exception:
    System = None
    DB = None

from Snippets import revit_helpers

import models


def _id_value(value):
    return revit_helpers.get_elementid_value(value)


def _storage_type_name(storage_type):
    if DB is None:
        return str(storage_type or "none").lower()
    mapping = {
        DB.StorageType.String: "string",
        DB.StorageType.Integer: "integer",
        DB.StorageType.Double: "double",
        DB.StorageType.ElementId: "element_id",
    }
    return mapping.get(storage_type, "none")


def _definition_name(parameter):
    try:
        return str(parameter.Definition.Name or "")
    except Exception:
        return ""


def _shared_guid(parameter):
    try:
        if bool(parameter.IsShared):
            return str(parameter.GUID).lower()
    except Exception:
        pass
    return None


def _spec_type(parameter):
    try:
        data_type = parameter.Definition.GetDataType()
        return str(data_type.TypeId or data_type)
    except Exception:
        return ""


def descriptor_from_parameter(parameter, scope):
    scope = "type" if str(scope).lower() == "type" else "instance"
    name = _definition_name(parameter)
    param_id = _id_value(getattr(parameter, "Id", None))
    shared_guid = _shared_guid(parameter)
    if shared_guid:
        identity_kind = "shared"
        identity_value = shared_guid
    elif param_id < 0:
        identity_kind = "builtin"
        identity_value = str(param_id)
    elif param_id > 0:
        identity_kind = "project"
        identity_value = str(param_id)
    else:
        identity_kind = "name"
        identity_value = name.strip().lower()
    key = "{}:{}:{}".format(identity_kind, identity_value, scope)
    return {
        "key": key,
        "name": name,
        "scope": scope,
        "identity_kind": identity_kind,
        "identity_value": identity_value,
        "parameter_id": param_id if param_id else None,
        "shared_guid": shared_guid,
        "builtin_id": param_id if param_id < 0 else None,
        "storage_type": _storage_type_name(getattr(parameter, "StorageType", None)),
        "spec_type": _spec_type(parameter),
    }


def _iter_parameters(owner):
    if owner is None:
        return []
    try:
        return list(owner.Parameters)
    except Exception:
        return []


def discover_parameters(elements, source_document):
    elements = list(elements or [])
    descriptors = {}
    availability = {}
    total = len(elements)
    for element in elements:
        seen = set()
        for parameter in _iter_parameters(element):
            descriptor = descriptor_from_parameter(parameter, "instance")
            key = descriptor["key"]
            descriptors[key] = descriptor
            seen.add(key)
        type_element = revit_helpers.get_type_element(element, doc=source_document)
        if type_element is not None:
            for parameter in _iter_parameters(type_element):
                descriptor = descriptor_from_parameter(parameter, "type")
                key = descriptor["key"]
                descriptors[key] = descriptor
                seen.add(key)
        for key in seen:
            availability[key] = int(availability.get(key, 0) or 0) + 1
    result = []
    for key, descriptor in list(descriptors.items()):
        item = dict(descriptor)
        item["available_count"] = int(availability.get(key, 0) or 0)
        item["element_count"] = total
        result.append(item)
    return sorted(
        result,
        key=lambda item: (
            str(item.get("name") or "").lower(),
            0 if item.get("scope") == "instance" else 1,
            str(item.get("key") or ""),
        ),
    )


def _enum_built_in(value):
    if DB is None:
        return None
    try:
        return System.Enum.ToObject(DB.BuiltInParameter, int(value))
    except Exception:
        try:
            return DB.BuiltInParameter(int(value))
        except Exception:
            return None


def _parameter_by_name(owner, name):
    if owner is None or not name:
        return None
    try:
        parameter = owner.LookupParameter(str(name))
        if parameter is not None:
            return parameter
    except Exception:
        pass
    target = str(name).strip().lower()
    for candidate in _iter_parameters(owner):
        if _definition_name(candidate).strip().lower() == target:
            return candidate
    return None


def resolve_parameter(owner, descriptor):
    if owner is None:
        return None
    descriptor = descriptor or {}
    kind = descriptor.get("identity_kind")
    if kind == "shared" and descriptor.get("shared_guid") and System is not None:
        try:
            parameter = owner.get_Parameter(System.Guid(str(descriptor.get("shared_guid"))))
            if parameter is not None:
                return parameter
        except Exception:
            pass
    if kind == "builtin" and descriptor.get("builtin_id") is not None:
        built_in = _enum_built_in(descriptor.get("builtin_id"))
        if built_in is not None:
            try:
                parameter = owner.get_Parameter(built_in)
                if parameter is not None:
                    return parameter
            except Exception:
                pass
    if kind == "project" and descriptor.get("parameter_id") is not None:
        try:
            parameter = owner.get_Parameter(
                revit_helpers.elementid_from_value(descriptor.get("parameter_id"))
            )
            if parameter is not None:
                return parameter
        except Exception:
            pass
    return _parameter_by_name(owner, descriptor.get("name"))


def normalize_parameter(parameter, source_document=None):
    if parameter is None:
        return models.missing_value()
    storage_name = _storage_type_name(getattr(parameter, "StorageType", None))
    try:
        has_value = bool(parameter.HasValue)
    except Exception:
        has_value = True
    if not has_value:
        return models.normalized_value(models.VALUE_BLANK, storage_name, None, "")
    try:
        if storage_name == "string":
            raw = parameter.AsString()
            display = raw
            if display is None:
                display = parameter.AsValueString()
            if raw is None or str(raw) == "":
                return models.normalized_value(models.VALUE_BLANK, storage_name, raw, display or "")
        elif storage_name == "integer":
            raw = int(parameter.AsInteger())
            display = parameter.AsValueString()
        elif storage_name == "double":
            raw = float(parameter.AsDouble())
            display = parameter.AsValueString()
        elif storage_name == "element_id":
            element_id = parameter.AsElementId()
            raw = _id_value(element_id)
            display = parameter.AsValueString()
            if not display and source_document is not None:
                try:
                    referenced = source_document.GetElement(element_id)
                    display = str(getattr(referenced, "Name", "") or raw)
                except Exception:
                    display = str(raw)
        else:
            return models.normalized_value(
                models.VALUE_UNSUPPORTED,
                storage_name,
                None,
                "Unsupported",
                "Unsupported Revit StorageType.",
            )
        if display is None or str(display) == "":
            display = str(raw)
        return models.normalized_value(models.VALUE_VALID, storage_name, raw, str(display))
    except Exception as ex:
        return models.normalized_value(
            models.VALUE_READ_ERROR,
            storage_name,
            None,
            "Read Error",
            str(ex),
        )


def read_properties(element, source_document, descriptors, type_cache=None):
    type_cache = type_cache if type_cache is not None else {}
    type_element = None
    values = {}
    for descriptor in list(descriptors or []):
        owner = element
        if descriptor.get("scope") == "type":
            type_id = None
            try:
                type_id = element.GetTypeId()
            except Exception:
                pass
            type_key = _id_value(type_id)
            if type_key in type_cache:
                type_element = type_cache[type_key]
            elif type_id is not None:
                try:
                    type_element = source_document.GetElement(type_id)
                except Exception:
                    type_element = None
                type_cache[type_key] = type_element
            owner = type_element
        parameter = resolve_parameter(owner, descriptor)
        values[str(descriptor.get("key") or "")] = normalize_parameter(
            parameter, source_document=source_document
        )
    return values


def descriptor_matches(left, right):
    left = left or {}
    right = right or {}
    if (
        left.get("identity_kind") in ("shared", "builtin")
        and left.get("key")
        and left.get("key") == right.get("key")
    ):
        return True
    if left.get("shared_guid") and left.get("shared_guid") == right.get("shared_guid"):
        return left.get("scope") == right.get("scope")
    if left.get("builtin_id") is not None and left.get("builtin_id") == right.get("builtin_id"):
        return left.get("scope") == right.get("scope")
    return (
        str(left.get("name") or "").strip().lower()
        == str(right.get("name") or "").strip().lower()
        and left.get("scope") == right.get("scope")
    )
