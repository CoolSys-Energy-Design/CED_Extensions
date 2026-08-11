# -*- coding: utf-8 -*-
"""Bridge to the MEPRFP Automation 2.0 lib (Element_Linker codec + profiles).

The MEPRFP lib is the codec of record for the shared ``Element_Linker``
parameter; importing it (instead of vendoring a copy) keeps the payload
round-trip identical to what the MEPRFP tools write. All imports are lazy
and the lib dir is APPENDED to sys.path so it can never shadow this
bundle's own modules or CEDLib. Because names like ``storage`` are
painfully generic, every import is verified to have resolved from the
MEPRFP lib dir; a foreign cached module raises a descriptive error
instead of being silently used.
"""

from __future__ import print_function

import os
import sys
import tempfile

THIS_DIR = os.path.dirname(os.path.abspath(__file__))

MEP_LIB_SUBPATH = os.path.join(
    "..", "..", "MEPRFP Automation 2.0.panel", "lib"
)


class MepBridgeError(Exception):
    pass


def mep_lib_dir():
    """Absolute path to the MEPRFP lib dir, or None when it is missing."""
    path = os.path.abspath(os.path.join(THIS_DIR, MEP_LIB_SUBPATH))
    return path if os.path.isdir(path) else None


def is_available():
    return mep_lib_dir() is not None


def _import_mep(name):
    lib_dir = mep_lib_dir()
    if lib_dir is None:
        raise MepBridgeError(
            "MEPRFP Automation 2.0 lib folder was not found next to this "
            "extension; Element Linker features are unavailable."
        )
    if lib_dir not in sys.path:
        sys.path.append(lib_dir)
    module = __import__(name)
    module_file = os.path.abspath(getattr(module, "__file__", "") or "")
    if not module_file.lower().startswith(lib_dir.lower()):
        raise MepBridgeError(
            "Module '{}' resolved to '{}' instead of the MEPRFP lib; a "
            "same-named module from another tool is already loaded.".format(
                name, module_file or "<unknown>"
            )
        )
    return module


def read_linker(element):
    """Element_Linker payload of an element as a plain dict, or None."""
    element_linker_io = _import_mep("element_linker_io")
    linker = element_linker_io.read_from_element(element)
    if linker is None:
        return None
    return linker.to_dict()


def update_linker(element, updates):
    """Rewrite selected Element_Linker fields on an element.

    Reads the existing payload, applies ``updates`` (field name -> value),
    and writes it back. Caller must hold an open transaction. Raises
    MepBridgeError when the element has no readable/writable payload.
    """
    element_linker_io = _import_mep("element_linker_io")
    linker = element_linker_io.read_from_element(element)
    if linker is None:
        raise MepBridgeError(
            "Element has no Element_Linker payload to update."
        )
    for field_name, value in (updates or {}).items():
        try:
            setattr(linker, field_name, value)
        except AttributeError:
            raise MepBridgeError(
                "Unknown Element_Linker field '{}'.".format(field_name)
            )
    try:
        element_linker_io.write_to_element(element, linker)
    except Exception as ex:
        raise MepBridgeError(str(ex))


def _parse_yaml_via_pyrevit(yaml_text):
    """Parse YAML with pyRevit's YamlDotNet wrapper (IronPython-safe).

    The MEPRFP lib's vendored PyYAML is 6.x (Python-3-only), so under this
    bundle's IronPython 2.7 engine ``yaml_io.parse`` raises at import time.
    ``pyrevit.coreutils.yaml`` wraps the .NET YamlDotNet library and works on
    every pyRevit engine; it is file-based, so the text takes a round trip
    through a temp file. Scalars come back as strings, which is fine for
    structure-only consumers like directive detection.
    """
    from pyrevit.coreutils import yaml as core_yaml
    handle, path = tempfile.mkstemp(suffix=".yaml", prefix="ced_pm_profiles_")
    try:
        stream = os.fdopen(handle, "wb")
        try:
            stream.write(yaml_text.encode("utf-8"))
        finally:
            stream.close()
        return core_yaml.load_as_dict(path) or {}
    finally:
        try:
            os.remove(path)
        except Exception:
            pass


def _parse_yaml_text(yaml_text):
    """Parse YAML text, preferring the MEPRFP codec, falling back to
    pyRevit's YamlDotNet wrapper. Raises when both fail."""
    first_error = None
    try:
        yaml_io = _import_mep("yaml_io")
        return yaml_io.parse(yaml_text)
    except Exception as ex:
        first_error = ex
    try:
        return _parse_yaml_via_pyrevit(yaml_text)
    except Exception as ex:
        raise MepBridgeError(
            "YAML parse failed via MEPRFP lib ({}) and via pyRevit "
            "YamlDotNet ({}).".format(first_error, ex)
        )


def load_profile_data(document):
    """Active MEPRFP profile payload as a dict, plus an optional warning.

    Returns ``(data_dict, warning_or_None)``. Any failure (missing lib,
    no stored payload, YAML parse error) degrades to ``({}, warning)`` so
    the sync can proceed treating every child as directive-less.
    """
    try:
        storage = _import_mep("storage")
    except MepBridgeError as ex:
        return {}, str(ex)
    try:
        payload = storage.read_payload(document)
    except Exception as ex:
        return {}, "MEPRFP profile storage could not be read: {}".format(ex)
    if not payload or not payload.get("yaml_text"):
        return {}, "No MEPRFP profile data is stored in this project."
    try:
        data = _parse_yaml_text(payload.get("yaml_text"))
    except Exception as ex:
        return {}, "MEPRFP profile YAML could not be parsed: {}".format(ex)
    if not isinstance(data, dict):
        return {}, "MEPRFP profile YAML has an unexpected shape."
    return data, None
