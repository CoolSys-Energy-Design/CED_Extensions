# -*- coding: utf-8 -*-
"""Startup hooks for CED ElecTools extension."""

import imp
import hashlib
import os
import sys

from pyrevit import forms, script

_CIRCUIT_MANAGER_REGISTERED = False
_CED_LIB_SIGNATURE_ENVVAR = "CED_LIB_RUNTIME_SIGNATURE"
_CED_MODULE_PREFIXES = (
    "CEDElectrical",
    "ExtensibleStorage",
    "LogicClasses",
    "Selection",
    "Snippets",
    "UIClasses",
    "pyrevitmep",
)


def _extension_root():
    return os.path.abspath(os.path.dirname(__file__))


def _workspace_root():
    return os.path.abspath(os.path.join(_extension_root(), ".."))


def _cedlib_root():
    return os.path.abspath(os.path.join(_workspace_root(), "CEDLib.lib"))


def _normalized_path(value):
    try:
        return os.path.normcase(os.path.realpath(os.path.abspath(str(value or ""))))
    except Exception:
        return ""


def _path_is_within(path, root):
    candidate = _normalized_path(path)
    root_path = _normalized_path(root)
    if not candidate or not root_path:
        return False
    return candidate == root_path or candidate.startswith(root_path + os.sep)


def _prioritize_syspath(path):
    target = _normalized_path(path)
    if not target:
        return
    retained = []
    for entry in list(sys.path):
        if _normalized_path(entry) == target:
            continue
        retained.append(entry)
    sys.path[:] = [os.path.abspath(path)] + retained


def _library_signature(lib_root):
    """Return a content signature so persistent engines detect library updates."""
    digest = hashlib.sha1()
    try:
        for current_root, dirs, files in os.walk(lib_root):
            dirs[:] = sorted([name for name in dirs if name != "__pycache__"])
            for file_name in sorted(files):
                if not file_name.lower().endswith(".py"):
                    continue
                file_path = os.path.join(current_root, file_name)
                relative_path = os.path.relpath(file_path, lib_root).replace("\\", "/")
                digest.update(relative_path.encode("utf-8"))
                with open(file_path, "rb") as source:
                    while True:
                        chunk = source.read(65536)
                        if not chunk:
                            break
                        digest.update(chunk)
        return digest.hexdigest()
    except Exception:
        return None


def _is_ced_module_name(module_name):
    text = str(module_name or "")
    for prefix in _CED_MODULE_PREFIXES:
        if text == prefix or text.startswith(prefix + "."):
            return True
    return False


def _module_is_from_root(module, lib_root):
    module_file = getattr(module, "__file__", None)
    return bool(module_file and _path_is_within(module_file, lib_root))


def _has_foreign_cached_ced_modules(lib_root):
    for module_name, module in list(sys.modules.items()):
        if not _is_ced_module_name(module_name):
            continue
        if module is None or not _module_is_from_root(module, lib_root):
            return True
    return False


def _clear_cached_ced_modules():
    cleared = 0
    module_names = [name for name in list(sys.modules.keys()) if _is_ced_module_name(name)]
    module_names.sort(key=lambda value: (value.count("."), len(value)), reverse=True)
    for module_name in module_names:
        try:
            del sys.modules[module_name]
            cleared += 1
        except Exception:
            pass
    return cleared


def _seed_runtime_paths():
    logger = script.get_logger()
    ext_root = _extension_root()
    lib_root = _cedlib_root()
    if os.path.isdir(lib_root):
        _prioritize_syspath(lib_root)
        signature = _library_signature(lib_root)
        try:
            previous_signature = script.get_envvar(_CED_LIB_SIGNATURE_ENVVAR)
        except Exception:
            previous_signature = None
        reset_modules = bool(
            _has_foreign_cached_ced_modules(lib_root)
            or (signature and str(previous_signature or "") != str(signature))
        )
        if reset_modules:
            cleared = _clear_cached_ced_modules()
            logger.info("CEDLib runtime cache refreshed. modules_cleared=%s", int(cleared))
        if signature:
            try:
                script.set_envvar(_CED_LIB_SIGNATURE_ENVVAR, signature)
            except Exception:
                pass
        logger.info("CEDLib prioritized on sys.path from CED ElecTools startup.")
    try:
        script.set_envvar("CED_EXTENSION_ROOT", ext_root)
    except Exception:
        pass
    try:
        script.set_envvar("CED_WORKSPACE_ROOT", _workspace_root())
    except Exception:
        pass
    try:
        script.set_envvar("CED_LIB_ROOT", lib_root if os.path.isdir(lib_root) else "")
    except Exception:
        pass


def _find_circuit_manager_panel_path():
    return os.path.abspath(
        os.path.join(
            _extension_root(),
            "AE pyTools.tab",
            "Electrical.panel",
            "Circuit Manager.pushbutton",
            "CircuitBrowserPanel.py",
        )
    )


def _register_circuit_manager_panel():
    global _CIRCUIT_MANAGER_REGISTERED
    logger = script.get_logger()
    if _CIRCUIT_MANAGER_REGISTERED:
        logger.info("Circuit Manager panel registration skipped (already registered in this startup run).")
        return

    panel_path = _find_circuit_manager_panel_path()
    if not panel_path or not os.path.exists(panel_path):
        logger.warning("Circuit Manager panel file not found under CED ElecTools extension.")
        return

    try:
        panel_module = imp.load_source("ced_electools_circuit_manager_panel", panel_path)
    except Exception as exc:
        logger.warning("Failed to load Circuit Manager panel: %s", exc)
        return

    panel_cls = getattr(panel_module, "CircuitBrowserPanel", None)
    if panel_cls is None:
        logger.warning("Circuit Manager panel class not found in: %s", panel_path)
        return

    try:
        if not forms.is_registered_dockable_panel(panel_cls):
            forms.register_dockable_panel(panel_cls, default_visible=False)
            logger.info("Circuit Manager panel registered successfully.")
        else:
            logger.info("Circuit Manager panel already registered.")
        _CIRCUIT_MANAGER_REGISTERED = True
    except Exception as exc:
        logger.warning("Failed to register Circuit Manager panel: %s", exc)


_seed_runtime_paths()
_register_circuit_manager_panel()
