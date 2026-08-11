# -*- coding: utf-8 -*-
"""Self-contained codec for the shared ``Element_Linker`` text parameter.

Parameter Monitor only needs to read and rewrite the payload that lives on
elements in the model, so this module carries the (stable) MEPRFP 2.0
payload contract itself — codec version 1 JSON, plus the legacy
"Key: value" text format for old projects — with no dependency on the
MEPRFP Automation panel being present.

Pure Python except for the two functions that touch a Revit parameter;
everything else is testable outside Revit.
"""

from __future__ import print_function

import json
import re

PARAMETER_NAME = "Element_Linker"
CODEC_VERSION = 1

FIELDS = (
    "led_id",
    "set_id",
    "location_ft",
    "rotation_deg",
    "parent_rotation_deg",
    "parent_element_id",
    "level_id",
    "element_id",
    "facing",
    "host_name",
    "parent_location_ft",
    "ckt_circuit_number",
    "ckt_panel",
    "space_id",
    "space_profile_id",
)

_LEGACY_FIELD_MAP = {
    "Linked Element Definition ID": "led_id",
    "Set Definition ID": "set_id",
    "Location XYZ (ft)": "location_ft",
    "Rotation (deg)": "rotation_deg",
    "Parent Rotation (deg)": "parent_rotation_deg",
    "Parent ElementId": "parent_element_id",
    "Parent Element ID": "parent_element_id",
    "LevelId": "level_id",
    "Level Id": "level_id",
    "ElementId": "element_id",
    "Element ID": "element_id",
    "Element Id": "element_id",
    "FacingOrientation": "facing",
    "Host Name": "host_name",
    "Parent_location": "parent_location_ft",
    "CKT_Circuit Number_CEDT": "ckt_circuit_number",
    "CKT_Panel_CEDT": "ckt_panel",
}

_TUPLE_FIELDS = set(["location_ft", "facing", "parent_location_ft"])
_INT_FIELDS = set(["parent_element_id", "level_id", "element_id", "space_id"])
_FLOAT_FIELDS = set(["rotation_deg", "parent_rotation_deg"])

_LEGACY_INLINE_KEY_RE = re.compile(
    r"({}):\s*".format("|".join(re.escape(key) for key in _LEGACY_FIELD_MAP))
)


class LinkerCodecError(Exception):
    pass


def empty_linker():
    return dict((name, None) for name in FIELDS)


def _coerce_legacy_value(field_name, raw):
    if raw is None:
        return None
    raw = raw.strip()
    if raw == "" or raw == "Not found":
        return None
    if field_name in _INT_FIELDS:
        try:
            return int(raw)
        except (ValueError, TypeError):
            return None
    if field_name in _FLOAT_FIELDS:
        try:
            return float(raw)
        except (ValueError, TypeError):
            return None
    if field_name in _TUPLE_FIELDS:
        parts = [part.strip() for part in raw.split(",")]
        if len(parts) != 3:
            return None
        try:
            return [float(part) for part in parts]
        except (ValueError, TypeError):
            return None
    return raw


def _parse_legacy_kv(text):
    """Parse legacy text (multiline or inline) into {legacy_key: raw_string}."""
    text = (text or "").strip()
    if not text:
        return {}
    if "\n" in text:
        out = {}
        for line in text.splitlines():
            if ":" not in line:
                continue
            key, _sep, value = line.partition(":")
            out[key.strip()] = value.strip()
        return out
    matches = list(_LEGACY_INLINE_KEY_RE.finditer(text))
    if not matches:
        return {}
    out = {}
    for index, match in enumerate(matches):
        key = match.group(1)
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        out[key] = text[start:end].rstrip().rstrip(",").strip(" ,")
    return out


def parse_payload(text):
    """Parse an Element_Linker payload string into a field dict, or None.

    JSON (codec v1) first, then the legacy text formats. Unreadable or
    blank payloads return None.
    """
    if text is None or not text.strip():
        return None
    linker = None
    try:
        data = json.loads(text)
    except (ValueError, TypeError):
        data = None
    if isinstance(data, dict):
        version = data.get("v")
        if version is not None and version != CODEC_VERSION:
            return None
        linker = empty_linker()
        for name in FIELDS:
            if name in data:
                linker[name] = data[name]
        return linker
    raw = _parse_legacy_kv(text)
    if not raw:
        return None
    linker = empty_linker()
    for legacy_key, value in raw.items():
        name = _LEGACY_FIELD_MAP.get(legacy_key)
        if name is None:
            continue
        if linker.get(name) is not None:
            continue
        linker[name] = _coerce_legacy_value(name, value)
    return linker


def serialize_payload(linker):
    """Serialize a field dict to the exact JSON the MEPRFP codec writes."""
    payload = {"v": CODEC_VERSION}
    for name in FIELDS:
        payload[name] = (linker or {}).get(name)
    return json.dumps(payload, separators=(",", ":"), sort_keys=False)


def apply_updates(linker, updates):
    """Return a copy of a field dict with updates applied; unknown field
    names raise LinkerCodecError."""
    result = dict(linker or empty_linker())
    for name, value in (updates or {}).items():
        if name not in FIELDS:
            raise LinkerCodecError(
                "Unknown Element_Linker field '{}'.".format(name)
            )
        result[name] = value
    return result


def _lookup_param(element):
    if element is None:
        return None
    try:
        return element.LookupParameter(PARAMETER_NAME)
    except Exception:
        return None


def read_linker(element):
    """Element_Linker payload of an element as a plain dict, or None."""
    param = _lookup_param(element)
    if param is None:
        return None
    try:
        text = param.AsString()
    except Exception:
        return None
    return parse_payload(text)


def update_linker(element, updates):
    """Rewrite selected Element_Linker fields on an element.

    Caller must hold an open transaction. Raises LinkerCodecError when the
    parameter is absent, read-only, or an update names an unknown field.
    """
    param = _lookup_param(element)
    if param is None:
        raise LinkerCodecError(
            "Element has no '{}' parameter.".format(PARAMETER_NAME)
        )
    try:
        read_only = bool(param.IsReadOnly)
    except Exception:
        read_only = False
    if read_only:
        raise LinkerCodecError(
            "The '{}' parameter is read-only on this element.".format(
                PARAMETER_NAME
            )
        )
    try:
        existing = parse_payload(param.AsString())
    except Exception:
        existing = None
    updated = apply_updates(existing or empty_linker(), updates)
    param.Set(serialize_payload(updated))
