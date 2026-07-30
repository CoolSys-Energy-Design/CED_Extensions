# -*- coding: utf-8 -*-
"""
Extensible Storage primitives for the MEPRFP 2.0 active YAML store.

The schema is independent from the original MEP Automation panel. Tools
in the 2.0 panel only ever read and write this schema, and the legacy
panel never sees it.

Layout (v4 — current). Four typed map fields keep the schema future-
extensible without GUID changes; new keys land in the appropriate map
without re-building the schema. ``IntMap`` is ``Int64`` on Revit 2024+
and falls back to ``Int32`` on older builds; the read path tolerates
both.

Keys consumed by the application::

  StringMap[KEY_YAML_TEXT]            canonical v100 YAML
  StringMap[KEY_SOURCE_PATH]          last imported file path
  StringMap[KEY_LAST_MODIFIED_UTC]    ISO-8601 timestamp of last save
  IntMap   [KEY_STORE_VERSION]        layout version (currently 1)
  IntMap   [KEY_SCHEMA_VERSION]       internal MEPRFP schema version (100)

The ``BoolMap`` and ``DoubleMap`` fields are declared but currently
empty — they exist so future flags / numeric ratios don't require a
schema bump.

Reads fall back to the legacy v1 schema (simple-fields layout under
GUID ``a7d4e2f1-…``) if the v4 entity is missing. Writes always go to
v4; the legacy entity is left orphaned for older tooling that may
still want to read it.
"""

import clr  # noqa: F401  -- needed before importing Autodesk.Revit.DB

from Autodesk.Revit.DB.ExtensibleStorage import (  # noqa: E402
    Entity,
    Schema,
    SchemaBuilder,
    AccessLevel,
)
from System import Guid, Int32, String  # noqa: E402

import _es_v4  # noqa: E402


# ---------------------------------------------------------------------
# Schema GUIDs
# ---------------------------------------------------------------------

# v4 (current) — 4-map layout.
SCHEMA_GUID_STR = "e3f9b6a4-5d2c-4f81-9b3e-7c1a8f6d4e2b"
SCHEMA_GUID = Guid(SCHEMA_GUID_STR)
SCHEMA_NAME = "MEPRFP_Automation_2_YamlStore_v4"
SCHEMA_DOC = (
    "MEPRFP Automation 2.0 active YAML storage. Four typed map fields "
    "keyed by string; YAML text lives in StringMap['yaml_text']."
)

# v1 (legacy in-2.0) — simple-fields layout. Read-only fallback.
LEGACY_V1_SCHEMA_GUID_STR = "a7d4e2f1-9c3b-4e8a-b6d5-f3c1a8e9b2d4"
LEGACY_V1_SCHEMA_GUID = Guid(LEGACY_V1_SCHEMA_GUID_STR)
LEGACY_V1_SCHEMA_NAME = "MEPRFP_Automation_2_YamlStore"
LEGACY_V1_FIELD_STORE_VERSION = "StoreVersion"
LEGACY_V1_FIELD_YAML_TEXT = "YamlText"
LEGACY_V1_FIELD_SOURCE_PATH = "SourcePath"
LEGACY_V1_FIELD_SCHEMA_VERSION = "SchemaVersion"
LEGACY_V1_FIELD_LAST_MODIFIED_UTC = "LastModifiedUtc"

STORE_LAYOUT_VERSION = 1


# ---------------------------------------------------------------------
# Map keys
# ---------------------------------------------------------------------

KEY_YAML_TEXT = "yaml_text"
KEY_SOURCE_PATH = "source_path"
KEY_LAST_MODIFIED_UTC = "last_modified_utc"
KEY_STORE_VERSION = "store_version"
KEY_SCHEMA_VERSION = "schema_version"


# ---------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------

StorageError = _es_v4.StorageError


# ---------------------------------------------------------------------
# Schema accessors
# ---------------------------------------------------------------------

def get_or_create_schema():
    """Build (or look up) the v4 schema."""
    return _es_v4.get_or_create_schema(SCHEMA_GUID, SCHEMA_NAME, SCHEMA_DOC)


def _legacy_v1_schema():
    """Look up the legacy v1 schema; return None if it doesn't exist
    in this Revit session yet."""
    return Schema.Lookup(LEGACY_V1_SCHEMA_GUID)


# ---------------------------------------------------------------------
# Read
# ---------------------------------------------------------------------

def read_payload(doc):
    """Return the stored payload as a dict, or ``None`` if no entity exists.

    Tries the v4 entity first; falls back to legacy v1 (simple fields)
    if v4 is missing. The returned dict shape is the same regardless::

        {
          "store_version": int,
          "yaml_text": str,
          "source_path": str,
          "schema_version": int,
          "last_modified_utc": str,
          "_legacy_v1": bool,    # True if data came from the v1 fallback
        }
    """
    payload = _read_v4(doc)
    if payload is not None:
        payload["_legacy_v1"] = False
        return payload

    payload = _read_legacy_v1(doc)
    if payload is not None:
        payload["_legacy_v1"] = True
        return payload

    return None


def _read_v4(doc):
    schema = get_or_create_schema()
    entity = _es_v4.get_entity(doc, schema)
    if entity is None:
        return None
    maps = _es_v4.read_maps(entity)
    if maps is None:
        return None
    sm = maps["string_map"]
    im = maps["int_map"]
    return {
        "store_version": int(im.get(KEY_STORE_VERSION) or STORE_LAYOUT_VERSION),
        "yaml_text": sm.get(KEY_YAML_TEXT) or "",
        "source_path": sm.get(KEY_SOURCE_PATH) or "",
        "schema_version": int(im.get(KEY_SCHEMA_VERSION) or 0),
        "last_modified_utc": sm.get(KEY_LAST_MODIFIED_UTC) or "",
    }


def _read_legacy_v1(doc):
    schema = _legacy_v1_schema()
    if schema is None:
        return None
    # v1 stored its entity on ProjectInformation, not DataStorage.
    entity = _es_v4.get_legacy_project_info_entity(doc, schema)
    if entity is None:
        return None
    return {
        "store_version": int(entity.Get[Int32](LEGACY_V1_FIELD_STORE_VERSION) or 0),
        "yaml_text": entity.Get[String](LEGACY_V1_FIELD_YAML_TEXT) or "",
        "source_path": entity.Get[String](LEGACY_V1_FIELD_SOURCE_PATH) or "",
        "schema_version": int(
            entity.Get[Int32](LEGACY_V1_FIELD_SCHEMA_VERSION) or 0
        ),
        "last_modified_utc": entity.Get[String](LEGACY_V1_FIELD_LAST_MODIFIED_UTC) or "",
    }


# ---------------------------------------------------------------------
# Write
# ---------------------------------------------------------------------

def write_payload(doc, yaml_text, source_path, schema_version, last_modified_utc):
    """Persist the payload onto a DataStorage element in the v4 entity.

    Always writes v4. If a legacy v1 entity exists on
    ``ProjectInformation``, it is left in place — older 2.0 builds
    reading the same project will see stale data, but no data is lost.
    Caller manages the Revit transaction.
    """
    schema = get_or_create_schema()
    entity = _es_v4.build_entity(
        schema,
        string_map={
            KEY_YAML_TEXT: yaml_text or "",
            KEY_SOURCE_PATH: source_path or "",
            KEY_LAST_MODIFIED_UTC: last_modified_utc or "",
        },
        int_map={
            KEY_STORE_VERSION: STORE_LAYOUT_VERSION,
            KEY_SCHEMA_VERSION: int(schema_version) if schema_version else 0,
        },
    )
    _es_v4.set_entity(doc, entity, ds_name=SCHEMA_NAME)


def clear_payload(doc):
    """Delete the v4 stored entity (its DataStorage element). Caller
    manages the transaction.

    The legacy v1 entity on ``ProjectInformation`` (if present) is
    left untouched. Use ``clear_legacy_v1_payload`` to remove it.
    """
    schema = get_or_create_schema()
    _es_v4.delete_entity(doc, schema)


def clear_legacy_v1_payload(doc):
    """Delete the legacy v1 entity off ProjectInformation. No-op
    if no v1 entity exists in this project."""
    schema = _legacy_v1_schema()
    if schema is None:
        return
    pi = _es_v4.project_info_or_raise(doc)
    try:
        pi.DeleteEntity(schema)
    except Exception:
        pass


# ---------------------------------------------------------------------
# Convenience
# ---------------------------------------------------------------------

def has_v4_entity(doc):
    schema = Schema.Lookup(SCHEMA_GUID)
    if schema is None:
        return False
    return _es_v4.get_entity(doc, schema) is not None


def has_legacy_v1_entity(doc):
    schema = _legacy_v1_schema()
    if schema is None:
        return False
    return _es_v4.get_legacy_project_info_entity(doc, schema) is not None
