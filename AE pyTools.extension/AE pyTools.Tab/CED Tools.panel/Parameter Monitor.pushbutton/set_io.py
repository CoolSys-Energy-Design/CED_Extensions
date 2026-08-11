# -*- coding: utf-8 -*-
"""Versioned import/export for portable Tracking Set definitions."""

from __future__ import print_function

import copy
import io
import json

import models


class DefinitionFormatError(ValueError):
    pass


def portable_definition(tracking_set):
    tracking_set = tracking_set or {}
    return {
        "name": str(tracking_set.get("name") or "Tracking Set"),
        "category": copy.deepcopy(tracking_set.get("category") or {}),
        "tracked_properties": copy.deepcopy(tracking_set.get("tracked_properties") or []),
        "location_defaults": copy.deepcopy(tracking_set.get("location_defaults") or {}),
        "active": bool(tracking_set.get("active", True)),
        "scan_policy": "manual",
        "source_hint": {
            "source_type": str((tracking_set.get("source") or {}).get("source_type") or models.SOURCE_HOST),
            "display_name": str((tracking_set.get("source") or {}).get("display_name") or ""),
        },
    }


def build_export_document(tracking_sets):
    return {
        "schema_version": models.DEFINITION_SCHEMA_VERSION,
        "tool": "CED Parameter Monitor",
        "tool_version": models.TOOL_VERSION,
        "tracking_sets": [portable_definition(item) for item in list(tracking_sets or [])],
    }


def dumps(tracking_sets):
    return json.dumps(build_export_document(tracking_sets), indent=2, sort_keys=True)


def dump_file(path, tracking_sets):
    with io.open(path, "w", encoding="utf-8") as stream:
        stream.write(dumps(tracking_sets))


def _validate_definition(definition):
    if not isinstance(definition, dict):
        raise DefinitionFormatError("Each Tracking Set definition must be an object.")
    category = definition.get("category")
    if not isinstance(category, dict) or not (category.get("id") is not None or category.get("name")):
        raise DefinitionFormatError("A Tracking Set definition has no resolvable category.")
    properties = definition.get("tracked_properties")
    if not isinstance(properties, list):
        raise DefinitionFormatError("tracked_properties must be a list.")
    for descriptor in properties:
        if not isinstance(descriptor, dict) or not descriptor.get("key"):
            raise DefinitionFormatError("Every tracked property requires a stable key.")
    result = copy.deepcopy(definition)
    result.setdefault("name", category.get("name") or "Tracking Set")
    result.setdefault("location_defaults", {})
    result.setdefault("active", True)
    result["scan_policy"] = "manual"
    result.setdefault("source_hint", {"source_type": models.SOURCE_HOST, "display_name": "Host Model"})
    return result


def loads(text):
    try:
        document = json.loads(text)
    except Exception as ex:
        raise DefinitionFormatError("Invalid JSON: {}".format(ex))
    if not isinstance(document, dict):
        raise DefinitionFormatError("Definition file must contain a JSON object.")
    version = int(document.get("schema_version", 0) or 0)
    if version != models.DEFINITION_SCHEMA_VERSION:
        raise DefinitionFormatError(
            "Definition schema {} is not supported (expected {}).".format(
                version, models.DEFINITION_SCHEMA_VERSION
            )
        )
    definitions = document.get("tracking_sets")
    if not isinstance(definitions, list):
        raise DefinitionFormatError("tracking_sets must be a list.")
    return [_validate_definition(item) for item in definitions]


def load_file(path):
    with io.open(path, "r", encoding="utf-8-sig") as stream:
        return loads(stream.read())

