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
import text_service


_SPEC_LABEL_CACHE = {}


def _id_value(value):
    numeric = revit_helpers.get_elementid_value(value, default=0)
    if numeric:
        return numeric
    try:
        return int(value or 0)
    except Exception:
        return 0


def _storage_type_name(storage_type):
    if DB is None:
        return text_service.to_text(storage_type or "none").lower()
    mapping = {
        DB.StorageType.String: "string",
        DB.StorageType.Integer: "integer",
        DB.StorageType.Double: "double",
        DB.StorageType.ElementId: "element_id",
    }
    return mapping.get(storage_type, "none")


def _definition_name(parameter):
    try:
        return text_service.to_text(
            parameter.Definition.Name or "", context=u"Parameter definition name"
        )
    except Exception:
        return ""


def _shared_guid(parameter):
    try:
        if bool(parameter.IsShared):
            guid = parameter.GUID
            try:
                guid = guid.ToString()
            except Exception:
                pass
            return text_service.to_text(
                guid, context=u"Shared parameter GUID"
            ).strip().lower()
    except Exception:
        pass
    return None


def spec_type_label(value):
    raw = text_service.to_text(value or "").strip()
    if not raw:
        return "Other"
    lowered = raw.lower()
    if (
        "forgetypeid" in lowered
        or ("object at" in lowered and "0x" in lowered)
        or raw.startswith("<")
    ):
        return "Other"
    cached = _SPEC_LABEL_CACHE.get(raw)
    if cached is not None:
        return cached
    token = raw.rsplit(":", 1)[-1]
    if "-" in token:
        candidate, version = token.rsplit("-", 1)
        if version and version[0].isdigit():
            token = candidate
    token = token.replace(".", " ").replace("_", " ").replace("-", " ")
    label = " ".join([part.capitalize() for part in token.split()])
    _SPEC_LABEL_CACHE[raw] = label
    return label


def _spec_type_info(parameter):
    try:
        data_type = parameter.Definition.GetDataType()
        type_id = text_service.to_text(
            getattr(data_type, "TypeId", "") or ""
        ).strip()
    except Exception:
        return "", "Other"
    if not type_id:
        return "", "Other"
    cached = _SPEC_LABEL_CACHE.get(type_id)
    if cached is not None:
        return type_id, cached
    label = ""
    if DB is not None:
        try:
            label = text_service.to_text(DB.LabelUtils.GetLabelForSpec(data_type))
        except Exception:
            pass
    label = label or spec_type_label(type_id)
    _SPEC_LABEL_CACHE[type_id] = label
    return type_id, label


def _spec_type(parameter):
    return _spec_type_info(parameter)[0]


def descriptor_from_parameter(parameter, scope):
    scope = "type" if text_service.to_text(scope).lower() == "type" else "instance"
    name = _definition_name(parameter)
    param_id = _id_value(getattr(parameter, "Id", None))
    shared_guid = _shared_guid(parameter)
    storage_type = _storage_type_name(getattr(parameter, "StorageType", None))
    spec_type, _spec_label = _spec_type_info(parameter)
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
        # A name-only identity is a legacy/last-resort case. Include the value
        # shape so two API parameters with the same localized display name do
        # not collapse into one discovery row.
        identity_value = "{}|{}|{}".format(
            name.strip().lower(), storage_type, spec_type
        )
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
        "storage_type": storage_type,
        "spec_type": spec_type,
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
            text_service.to_text(item.get("name") or "").lower(),
            0 if item.get("scope") == "instance" else 1,
            text_service.to_text(item.get("key") or ""),
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


def _parameters_by_name(owner, name):
    """Return every matching parameter; never use Revit's random first match.

    Autodesk explicitly documents that ``LookupParameter`` returns the first
    of potentially several same-name parameters and that the match is not
    deterministic. ``GetParameters`` is therefore the primary API here, with
    enumeration retained only for test doubles and defensive compatibility.
    """
    if owner is None or not name:
        return []
    requested = text_service.to_text(name)
    try:
        matches = list(owner.GetParameters(requested) or [])
        if matches:
            return matches
    except Exception:
        pass
    target = requested.strip().lower()
    return [
        candidate
        for candidate in _iter_parameters(owner)
        if _definition_name(candidate).strip().lower() == target
    ]


def _parameter_shape_matches(parameter, descriptor):
    descriptor = descriptor or {}
    storage_type = text_service.to_text(descriptor.get("storage_type") or "").lower()
    spec_type = text_service.to_text(descriptor.get("spec_type") or "")
    if storage_type and storage_type != _storage_type_name(
        getattr(parameter, "StorageType", None)
    ):
        return False
    if spec_type and spec_type != _spec_type(parameter):
        return False
    return True


def _unique_parameter_by_shape(parameters, descriptor):
    matches = [
        item for item in list(parameters or [])
        if _parameter_shape_matches(item, descriptor)
    ]
    if len(matches) == 1:
        return matches[0]
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
        # Never degrade an exact GUID identity to a display-name lookup.
        return None
    if kind == "builtin" and descriptor.get("builtin_id") is not None:
        built_in = _enum_built_in(descriptor.get("builtin_id"))
        if built_in is not None:
            try:
                parameter = owner.get_Parameter(built_in)
                if parameter is not None:
                    return parameter
            except Exception:
                pass
        # Never degrade a built-in identity to a localized display name.
        return None
    if kind == "project" and descriptor.get("parameter_id") is not None:
        target_id = _id_value(descriptor.get("parameter_id"))
        # Element.get_Parameter has overloads for BuiltInParameter,
        # Definition, and Guid -- not for a project parameter ElementId.
        # Get every same-name candidate and match Parameter.Id instead.
        for parameter in _parameters_by_name(owner, descriptor.get("name")):
            if _id_value(getattr(parameter, "Id", None)) == target_id:
                return parameter
        return None
    # Legacy imported/name-only definitions are allowed only when name plus
    # stored value shape identifies one unambiguous parameter.
    return _unique_parameter_by_shape(
        _parameters_by_name(owner, descriptor.get("name")), descriptor
    )


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
        parameter_name = _definition_name(parameter) or u"Unnamed parameter"
        text_context = u"Parameter '{}' value".format(parameter_name)
        if storage_name == "string":
            raw = parameter.AsString()
            raw = text_service.to_text(raw, context=text_context) if raw is not None else None
            display = raw
            if display is None:
                display_value = parameter.AsValueString()
                display = (
                    text_service.to_text(display_value, context=text_context)
                    if display_value is not None else u""
                )
            if raw is None or raw == "":
                return models.normalized_value(models.VALUE_BLANK, storage_name, raw, display or "")
        elif storage_name == "integer":
            raw = int(parameter.AsInteger())
            display_value = parameter.AsValueString()
            display = (
                text_service.to_text(display_value, context=text_context)
                if display_value is not None else None
            )
        elif storage_name == "double":
            raw = float(parameter.AsDouble())
            display_value = parameter.AsValueString()
            display = (
                text_service.to_text(display_value, context=text_context)
                if display_value is not None else None
            )
        elif storage_name == "element_id":
            element_id = parameter.AsElementId()
            raw = _id_value(element_id)
            display = parameter.AsValueString()
            if not display and source_document is not None:
                try:
                    referenced = source_document.GetElement(element_id)
                    display = text_service.to_text(
                        getattr(referenced, "Name", "") or raw,
                        context=u"Referenced element name for {}".format(parameter_name),
                    )
                except Exception:
                    display = text_service.to_text(raw, context=text_context)
        else:
            return models.normalized_value(
                models.VALUE_UNSUPPORTED,
                storage_name,
                None,
                "Unsupported",
                "Unsupported Revit StorageType.",
            )
        if display is None or display == "":
            display = text_service.to_text(raw, context=text_context)
        return models.normalized_value(models.VALUE_VALID, storage_name, raw, display)
    except Exception as ex:
        return models.normalized_value(
            models.VALUE_READ_ERROR,
            storage_name,
            None,
            "Read Error",
            text_service.diagnostic_text(ex, u"Revit parameter read failed."),
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
        values[text_service.to_text(descriptor.get("key") or "")] = normalize_parameter(
            parameter, source_document=source_document
        )
    return values


def descriptor_matches(left, right):
    left = left or {}
    right = right or {}
    if left.get("scope") != right.get("scope"):
        return False
    if left.get("key") and left.get("key") == right.get("key"):
        return True
    left_kind = text_service.to_text(left.get("identity_kind") or "name")
    right_kind = text_service.to_text(right.get("identity_kind") or "name")
    if left_kind == "shared":
        return (
            right_kind == "shared"
            and bool(left.get("shared_guid"))
            and left.get("shared_guid") == right.get("shared_guid")
        )
    if left_kind == "builtin":
        return (
            right_kind == "builtin"
            and left.get("builtin_id") is not None
            and left.get("builtin_id") == right.get("builtin_id")
        )
    if left_kind == "project":
        if right_kind != "project":
            return False
        if (
            left.get("parameter_id") is not None
            and left.get("parameter_id") == right.get("parameter_id")
        ):
            return True
    return (
        text_service.to_text(left.get("name") or "").strip().lower()
        == text_service.to_text(right.get("name") or "").strip().lower()
        and text_service.to_text(left.get("storage_type") or "").lower()
        == text_service.to_text(right.get("storage_type") or "").lower()
        and text_service.to_text(left.get("spec_type") or "")
        == text_service.to_text(right.get("spec_type") or "")
    )


def find_matching_descriptor(imported, available):
    """Map a stored/imported descriptor without choosing an ambiguous name."""
    candidates = [
        item for item in list(available or [])
        if descriptor_matches(imported, item)
    ]
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        exact_key = text_service.to_text((imported or {}).get("key") or "")
        exact = [
            item for item in candidates
            if exact_key and text_service.to_text(item.get("key") or "") == exact_key
        ]
        if len(exact) == 1:
            return exact[0]
    return None
