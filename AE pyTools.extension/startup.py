# -*- coding: utf-8 -*-
"""
Startup hook for after-sync parent parameter conflict checks.
"""

import getpass
import imp
import json
import os
import shutil
import time

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from pyrevit import forms, script, telemetry
try:
    import telemetry_route
except Exception:
    telemetry_route = None

try:
    from Autodesk.Revit.UI.Events import DocumentSynchronizedWithCentralEventArgs as UiSyncArgs
except Exception:
    UiSyncArgs = None

try:
    from Autodesk.Revit.DB.Events import DocumentSynchronizedWithCentralEventArgs as DbSyncArgs
except Exception:
    DbSyncArgs = None

try:
    from System import EventHandler
except Exception:
    EventHandler = None

try:
    from System.Diagnostics import Process
except Exception:
    Process = None

_SYNC_HANDLER_UI = None
_SYNC_HANDLER_APP = None
_MODULE = None
_IS_RUNNING = False

_DOCKABLE_REGISTERED = False

LOCAL_TELEMETRY_CLEANUP_VERSION = 1
PYTOOLS_ABOUT_METADATA_RELATIVE_PATH = os.path.join(
    "AE pyTools.Tab",
    "CED Tools.panel",
    "About.pushbutton",
    "about.yaml",
)

def _telemetry_source_folder():
    if telemetry_route is not None:
        return telemetry_route.telemetry_source_folder()
    appdata = os.environ.get("APPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Roaming"))
    return os.path.join(appdata, "pyRevit", "Extensions", "CED_pyTelemetry")


def _ensure_telemetry_source_folder():
    if telemetry_route is not None:
        return telemetry_route.ensure_telemetry_source_folder()
    source_folder = _telemetry_source_folder()
    if os.path.exists(source_folder):
        return source_folder, True, None
    try:
        os.makedirs(source_folder)
        return source_folder, True, None
    except Exception as exc:
        return source_folder, False, exc

def _normalize_path(value):
    if value in (None, ""):
        return ""
    return os.path.normcase(os.path.normpath(value))


def _canonical_telemetry_folder(value):
    if value in (None, ""):
        return ""
    try:
        return os.path.abspath(os.path.normpath(str(value)))
    except Exception:
        return str(value or "")


def _telemetry_folder_matches(current_value, expected_value):
    try:
        current_text = os.path.normcase(str(current_value or "").strip())
        expected_text = os.path.normcase(str(expected_value or "").strip())
        return current_text == expected_text
    except Exception:
        return current_value == expected_value


def _fallback_acc_root_is_viable(candidate_root):
    if not candidate_root or not os.path.isdir(candidate_root):
        return False

    project_files_path = os.path.join(candidate_root, "Project Files")
    usage_base_path = os.path.join(project_files_path, "03 Automations", "Usage")
    if not os.path.isdir(project_files_path) or not os.path.isdir(usage_base_path):
        return False

    try:
        project_folders = [
            name for name in os.listdir(project_files_path)
            if os.path.isdir(os.path.join(project_files_path, name))
        ]
    except Exception:
        return False

    has_automations = any(
        _normalize_path(name).lower() == _normalize_path("03 Automations").lower()
        for name in project_folders
    )
    if len(project_folders) == 1 and has_automations:
        return False

    return True


def _clean_yaml_scalar(value):
    value_text = str(value or "").split("#", 1)[0].strip()
    if len(value_text) >= 2:
        first_char = value_text[0]
        last_char = value_text[-1]
        if first_char == last_char and first_char in ("'", '"'):
            value_text = value_text[1:-1].strip()
    return value_text


def _read_pytools_release_metadata(metadata_path=None):
    result = {
        "toolbar_version": "",
        "build_version": "",
    }
    if not metadata_path:
        metadata_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            PYTOOLS_ABOUT_METADATA_RELATIVE_PATH,
        )

    try:
        with open(metadata_path, "r") as metadata_file:
            for raw_line in metadata_file:
                stripped = str(raw_line or "").strip()
                if not stripped or stripped.startswith("#"):
                    continue
                key, separator, value = stripped.partition(":")
                if not separator:
                    continue
                key = key.strip().lower()
                if key == "toolbar_version":
                    result["toolbar_version"] = _clean_yaml_scalar(value)
                elif key == "build":
                    result["build_version"] = _clean_yaml_scalar(value)
    except Exception:
        pass
    return result


_PYTOOLS_RELEASE_METADATA = _read_pytools_release_metadata()
PYTOOLS_TOOLBAR_VERSION = _PYTOOLS_RELEASE_METADATA.get("toolbar_version", "")
PYTOOLS_BUILD_VERSION = _PYTOOLS_RELEASE_METADATA.get("build_version", "")


def _configure_pyrevit_telemetry():
    logger = script.get_logger()
    source_folder, folder_ok, folder_error = _ensure_telemetry_source_folder()
    if not folder_ok:
        logger.warning("Telemetry folder not available: %s", folder_error)
        return
    source_folder = _canonical_telemetry_folder(source_folder)

    try:
        telemetry_cfg = script.get_config("telemetry")
        expected_settings = {
            "utc_timestamps": True,
            "active": True,
            "telemetry_file_dir": source_folder,
            "include_hooks": True,
        }

        changed_settings = []
        for setting_name, expected_value in expected_settings.items():
            current_value = telemetry_cfg.get_option(
                setting_name,
                default_value="",
            )
            if setting_name == "telemetry_file_dir":
                values_match = _telemetry_folder_matches(
                    current_value,
                    expected_value,
                )
            else:
                values_match = current_value == expected_value
            if values_match:
                continue
            telemetry_cfg.set_option(setting_name, value=expected_value)
            changed_settings.append(setting_name)

        if changed_settings:
            # Persist configuration for pyRevit to apply during its next normal
            # startup. CED must not initialize telemetry or create a second
            # session file inside the current Revit process.
            script.save_config()
            logger.info(
                "pyRevit telemetry configuration saved for next launch. "
                "changed=%s file_dir=%s",
                ", ".join(changed_settings),
                source_folder,
            )
        else:
            logger.info(
                "pyRevit telemetry already matched required settings. "
                "No config write needed."
            )
    except Exception as exc:
        logger.warning("Failed to configure pyRevit telemetry: %s", exc)

def _find_acc_root():
    if telemetry_route is not None:
        try:
            resolution = telemetry_route.resolve_usage_route(persist=True)
            root = resolution.get("resolved_root")
            if root:
                return root
            return None
        except Exception:
            return None
    candidates = [
        r"C:\ACC\ACCDocs\CoolSys\CED Content Collection",
        os.path.join(os.path.expanduser("~"), "DC", "ACCDocs", "CoolSys", "CED Content Collection"),
    ]
    for path in candidates:
        if _fallback_acc_root_is_viable(path):
            return path
    return None

ENV_HANDLER_KEY = "ced_parent_param_sync_handler_registered"
ENV_LAST_RUN_KEY = "ced_parent_param_sync_last_run"
ENV_RUNNING_KEY = "ced_parent_param_sync_running"
ENV_APP_CLOSING_HANDLER_KEY = "ced_app_closing_handler_registered"


def _module_path():
    return os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "AE pyTools.Tab",
            "MEP Automation.panel",
            "Parameter Flag Settings.pushbutton",
            "parent_param_conflicts.py",
        )
    )


def _load_checker():
    global _MODULE
    if _MODULE is not None:
        return _MODULE
    path = _module_path()
    if not os.path.exists(path):
        return None
    try:
        _MODULE = imp.load_source("ced_parent_param_conflicts", path)
        return _MODULE
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to load parent param conflict checker: %s", exc)
        return None


def _on_doc_sync(sender, args):
    global _IS_RUNNING
    doc = None
    try:
        doc = getattr(args, "Document", None)
    except Exception:
        doc = None
    if doc is None:
        try:
            doc = __revit__.ActiveUIDocument.Document
        except Exception:
            doc = None
    if doc is None:
        return
    if _IS_RUNNING:
        return
    if not _should_run_sync(doc):
        return
    checker = _load_checker()
    if checker is None:
        return
    try:
        _IS_RUNNING = True
        _set_env(ENV_RUNNING_KEY, "1")
        checker.run_sync_check(doc)
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Parent param conflict check failed: %s", exc)
    finally:
        _set_env(ENV_RUNNING_KEY, "0")
        _IS_RUNNING = False


def _sync_guard_host():
    uiapp = None
    try:
        uiapp = __revit__
    except Exception:
        uiapp = None
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    return app or uiapp


def _get_doc_key(doc):
    if doc is None:
        return None
    try:
        return doc.PathName or doc.Title
    except Exception:
        return None


def _get_env(name, default=None):
    try:
        value = script.get_envvar(name)
    except Exception:
        return default
    if value in (None, ""):
        return default
    return value


def _set_env(name, value):
    try:
        script.set_envvar(name, value)
    except Exception:
        return False
    return True


def _load_env_payload(raw_value):
    if isinstance(raw_value, dict):
        return raw_value
    try:
        return json.loads(raw_value)
    except Exception:
        return {}


def _should_run_sync(doc):
    running = _get_env(ENV_RUNNING_KEY)
    if str(running).strip() == "1":
        return False
    doc_key = _get_doc_key(doc)
    if not doc_key:
        return True
    now = time.time()
    payload = _load_env_payload(_get_env(ENV_LAST_RUN_KEY, "{}"))
    last_key = payload.get("doc_key")
    last_ts = payload.get("timestamp") or 0.0
    if last_key == doc_key and now - last_ts < 20.0:
        return False
    payload = {"doc_key": doc_key, "timestamp": now}
    _set_env(ENV_LAST_RUN_KEY, json.dumps(payload))
    return True


def _handler_registry(uiapp):
    if uiapp is None:
        return None
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    host = app or uiapp
    registry = getattr(host, "_ced_parent_param_sync_handlers", None)
    if registry is None:
        registry = {}
        try:
            setattr(host, "_ced_parent_param_sync_handlers", registry)
        except Exception:
            return None
    return registry


def _register_sync_handler():
    global _SYNC_HANDLER_UI, _SYNC_HANDLER_APP
    logger = script.get_logger()
    if EventHandler is None:
        logger.warning("Parent parameter sync handler not registered: EventHandler missing.")
        return
    if _get_env(ENV_HANDLER_KEY):
        return
    uiapp = None
    try:
        uiapp = __revit__
    except Exception:
        uiapp = None
    registry = _handler_registry(uiapp)
    app = None
    try:
        app = getattr(uiapp, "Application", None)
    except Exception:
        app = None
    if registry is not None and registry.get("registered"):
        return
    if app is not None and DbSyncArgs is not None and _SYNC_HANDLER_APP is None:
        try:
            handler = EventHandler[DbSyncArgs](_on_doc_sync)
            app.DocumentSynchronizedWithCentral += handler
            _SYNC_HANDLER_APP = handler
            if registry is not None:
                registry["registered"] = "app"
            _set_env(ENV_HANDLER_KEY, "app")
            logger.info("Parent parameter conflict app sync handler registered.")
            return
        except Exception as exc:
            logger.warning("App sync handler not registered: %s", exc)
    if uiapp is not None and UiSyncArgs is not None and _SYNC_HANDLER_UI is None:
        try:
            handler = EventHandler[UiSyncArgs](_on_doc_sync)
            uiapp.DocumentSynchronizedWithCentral += handler
            _SYNC_HANDLER_UI = handler
            if registry is not None:
                registry["registered"] = "ui"
            _set_env(ENV_HANDLER_KEY, "ui")
            logger.info("Parent parameter conflict UI sync handler registered.")
        except Exception as exc:
            logger.warning("UI sync handler not registered: %s", exc)


def _dockable_panel_path():
    return os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "AE pyTools.Tab",
            "MEP Automation.panel",
            "Place Single Profile.pushbutton",
            "PlaceSingleProfilePanel.py",
        )
    )


def _register_place_single_profile_panel():
    global _DOCKABLE_REGISTERED
    if _DOCKABLE_REGISTERED:
        return
    panel_path = _dockable_panel_path()
    if not os.path.exists(panel_path):
        return
    try:
        panel_module = imp.load_source("ced_place_single_profile_panel", panel_path)
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to load Place Single Profile panel: %s", exc)
        return
    panel_cls = getattr(panel_module, "PlaceSingleProfilePanel", None)
    if panel_cls is None:
        return
    try:
        if not forms.is_registered_dockable_panel(panel_cls):
            forms.register_dockable_panel(panel_cls, default_visible=False)
        _DOCKABLE_REGISTERED = True
    except Exception as exc:
        logger = script.get_logger()
        logger.warning("Failed to register Place Single Profile panel: %s", exc)


def _utc_now():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def _to_int(value):
    try:
        return int(value or 0)
    except Exception:
        return 0


def _is_telemetry_file_name(file_name):
    try:
        return str(file_name or "").lower().endswith("_telemetry.json")
    except Exception:
        return False


def _is_last_revit_process(process_type=None, diagnostics=None):
    """Return True only when no other live Revit is in this Windows session."""
    details = diagnostics if isinstance(diagnostics, dict) else {}
    details["status"] = "inspection_failed"
    details["error"] = ""
    details["other_process_id"] = None

    process_api = process_type or Process
    if process_api is None:
        details["error"] = "System.Diagnostics.Process is unavailable."
        return False

    try:
        current_process = process_api.GetCurrentProcess()
        current_process_id = int(current_process.Id)
        current_session_id = int(current_process.SessionId)
        details["current_process_id"] = current_process_id
        details["current_session_id"] = current_session_id

        revit_processes = process_api.GetProcessesByName("Revit")
        for revit_process in revit_processes:
            process_id = int(revit_process.Id)
            if process_id == current_process_id:
                continue

            session_id = int(revit_process.SessionId)
            if session_id != current_session_id:
                continue

            if not bool(revit_process.HasExited):
                details["status"] = "another_revit_process"
                details["other_process_id"] = process_id
                return False

        details["status"] = "last_revit_process"
        return True
    except Exception as ex:
        details["status"] = "inspection_failed"
        details["error"] = str(ex)
        return False


def _empty_current_file_result(current_path=""):
    return {
        "status": "not_found",
        "source_path": current_path or "",
        "file_name": "",
        "destination_path": "",
        "eligible": False,
        "found": False,
        "copy_attempted": False,
        "copied": False,
        "copy_failed": False,
        "copy_error": "",
        "delete_attempted": False,
        "deleted": False,
        "delete_failed": False,
        "delete_error": "",
        "dropped": False,
    }


def _transfer_and_delete_telemetry_file(source_path, destination_folder):
    result = _empty_current_file_result(current_path=source_path)
    file_name = os.path.basename(source_path or "")
    result["file_name"] = file_name
    result["eligible"] = _is_telemetry_file_name(file_name)

    if not result["eligible"]:
        result["status"] = "ineligible_name"
        return result

    if not os.path.isfile(source_path):
        result["status"] = "not_found"
        return result

    result["found"] = True
    if destination_folder:
        destination_path = os.path.join(destination_folder, file_name)
        result["destination_path"] = destination_path
        result["copy_attempted"] = True
        try:
            # Exact-name copy intentionally overwrites an existing destination.
            shutil.copyfile(source_path, destination_path)
            result["copied"] = True
        except Exception as ex:
            result["copy_failed"] = True
            result["copy_error"] = str(ex)
    else:
        result["copy_failed"] = True
        result["copy_error"] = "destination_unavailable"
        result["status"] = "destination_unavailable"
        return result

    # Local deletion is required even when the copy fails. Failed telemetry is
    # intentionally not retained for a later replay.
    result["delete_attempted"] = True
    try:
        os.remove(source_path)
        result["deleted"] = True
    except Exception as ex:
        result["delete_failed"] = True
        result["delete_error"] = str(ex)

    result["dropped"] = bool(result["copy_failed"] and result["deleted"])
    if result["delete_failed"]:
        result["status"] = "delete_failed"
    elif result["copy_failed"]:
        result["status"] = "dropped"
    else:
        result["status"] = "success"
    return result


def _local_telemetry_files(source_folder, excluded_paths=None):
    if not os.path.isdir(source_folder):
        return [], ""

    excluded = set()
    for path in list(excluded_paths or []):
        if path:
            excluded.add(_normalize_path(path))

    state_file_name = ".ced_usage_route_status.json"
    if telemetry_route is not None:
        state_file_name = getattr(
            telemetry_route,
            "STATE_FILE_NAME",
            state_file_name,
        )

    try:
        telemetry_files = []
        for file_name in sorted(os.listdir(source_folder)):
            if str(file_name).lower() == str(state_file_name).lower():
                continue
            if not _is_telemetry_file_name(file_name):
                continue
            source_path = os.path.join(source_folder, file_name)
            if _normalize_path(source_path) in excluded:
                continue
            if os.path.isfile(source_path):
                telemetry_files.append(source_path)
        return telemetry_files, ""
    except Exception as ex:
        return [], str(ex)


def _empty_legacy_cleanup_result():
    return {
        "status": "not_required",
        "ran": False,
        "deferred": False,
        "defer_reason": "",
        "files_found": 0,
        "files_deleted": 0,
        "files_failed": 0,
        "files_remaining": 0,
        "error": "",
        "complete": False,
    }


def _run_legacy_local_cleanup(source_folder):
    result = _empty_legacy_cleanup_result()
    result["ran"] = True
    result["status"] = "running"

    telemetry_files, list_error = _local_telemetry_files(source_folder)
    if list_error:
        result["status"] = "enumeration_failed"
        result["error"] = list_error
        return result

    result["files_found"] = len(telemetry_files)
    for source_path in telemetry_files:
        try:
            os.remove(source_path)
            result["files_deleted"] += 1
        except Exception as ex:
            result["files_failed"] += 1
            result["error"] = str(ex)

    remaining_files, verify_error = _local_telemetry_files(source_folder)
    result["files_remaining"] = len(remaining_files)
    if verify_error:
        result["status"] = "verification_failed"
        result["error"] = verify_error
        return result

    if result["files_failed"] or result["files_remaining"]:
        result["status"] = "partial_failure"
        return result

    result["status"] = "success"
    result["complete"] = True
    return result


def _empty_sweep_result():
    return {
        "status": "not_required",
        "ran": False,
        "deferred": False,
        "defer_reason": "",
        "files_found": 0,
        "files_copied": 0,
        "files_copy_failed": 0,
        "files_deleted": 0,
        "files_dropped": 0,
        "files_delete_failed": 0,
        "error": "",
    }


def _run_post_migration_sweep(source_folder, destination_folder, current_path=""):
    result = _empty_sweep_result()
    result["ran"] = True
    result["status"] = "running"

    telemetry_files, list_error = _local_telemetry_files(
        source_folder,
        excluded_paths=[current_path],
    )
    if list_error:
        result["status"] = "enumeration_failed"
        result["error"] = list_error
        return result

    result["files_found"] = len(telemetry_files)
    for source_path in telemetry_files:
        file_result = _transfer_and_delete_telemetry_file(
            source_path,
            destination_folder,
        )
        if file_result.get("copied"):
            result["files_copied"] += 1
        if file_result.get("copy_failed"):
            result["files_copy_failed"] += 1
            result["error"] = file_result.get("copy_error", "")
        if file_result.get("deleted"):
            result["files_deleted"] += 1
        if file_result.get("dropped"):
            result["files_dropped"] += 1
        if file_result.get("delete_failed"):
            result["files_delete_failed"] += 1
            result["error"] = file_result.get("delete_error", "")

    if result["files_copy_failed"] or result["files_delete_failed"]:
        result["status"] = "partial_success"
    else:
        result["status"] = "success"
    return result


def _resolve_shutdown_destination(username, source_folder):
    result = {
        "status": "route_unresolved",
        "route_status": "",
        "route_reason": "",
        "resolved_root": "",
        "destination_folder": "",
        "error": "",
    }

    acc_root = None
    if telemetry_route is not None:
        try:
            route_result = telemetry_route.resolve_usage_route(
                username=username,
                source_folder=source_folder,
                persist=True,
            )
            result["route_status"] = route_result.get("status", "")
            result["route_reason"] = route_result.get("reason", "")
            acc_root = route_result.get("resolved_root")
        except Exception as ex:
            result["route_status"] = "error"
            result["route_reason"] = str(ex)

    if not acc_root:
        try:
            acc_root = _find_acc_root()
        except Exception as ex:
            result["error"] = str(ex)

    if not acc_root:
        return result

    result["resolved_root"] = acc_root
    usage_base = os.path.join(
        acc_root,
        "Project Files",
        "03 Automations",
        "Usage",
    )
    if not os.path.isdir(usage_base):
        result["status"] = "usage_base_missing"
        return result

    try:
        if telemetry_route is not None and hasattr(telemetry_route, "ensure_user_folder"):
            folder_result = telemetry_route.ensure_user_folder(
                acc_root,
                username=username,
            )
            destination_folder = folder_result.get("path") or os.path.join(
                usage_base,
                username,
            )
            if not folder_result.get("ok"):
                result["status"] = "failed_create_user_folder"
                result["error"] = folder_result.get(
                    "reason",
                    "user_folder_unavailable",
                )
                return result
        else:
            destination_folder = os.path.join(usage_base, username)
            if not os.path.isdir(destination_folder):
                os.mkdir(destination_folder)
    except Exception as ex:
        result["status"] = "failed_create_user_folder"
        result["error"] = str(ex)
        return result

    result["status"] = "ready"
    result["destination_folder"] = destination_folder
    return result


def _state_snapshot_paths(source_folder, destination_folder):
    state_file_name = ".ced_usage_route_status.json"
    if telemetry_route is not None:
        state_file_name = getattr(
            telemetry_route,
            "STATE_FILE_NAME",
            state_file_name,
        )

    if telemetry_route is not None and hasattr(telemetry_route, "state_file_path"):
        source_path = telemetry_route.state_file_path(source_folder)
    else:
        source_path = os.path.join(source_folder, state_file_name)

    destination_path = ""
    if destination_folder:
        destination_path = os.path.join(destination_folder, state_file_name)
    return state_file_name, source_path, destination_path


def _planned_state_snapshot_result(source_folder, destination_folder):
    file_name, source_path, destination_path = _state_snapshot_paths(
        source_folder,
        destination_folder,
    )
    result = {
        "status": "destination_unavailable",
        "file_name": file_name,
        "source_path": source_path,
        "destination_path": destination_path,
        "copy_attempted": False,
        "copied": False,
        "overwritten": False,
        "source_retained": True,
        "delete_attempted": False,
        "deleted": False,
        "error": "",
        "updated_utc": _utc_now(),
    }
    if destination_path:
        result["status"] = "success"
        result["copy_attempted"] = True
        result["copied"] = True
        try:
            result["overwritten"] = os.path.isfile(destination_path)
        except Exception:
            result["overwritten"] = False
    return result


def _record_shutdown_state(log_data, source_folder):
    if telemetry_route is None:
        return False, "telemetry_route_unavailable"

    if hasattr(telemetry_route, "load_state") and hasattr(telemetry_route, "save_state"):
        snapshot_result = None
        try:
            snapshot_result = _planned_state_snapshot_result(
                source_folder,
                log_data.get("destination_folder", ""),
            )
            log_data["pytools_toolbar_version"] = PYTOOLS_TOOLBAR_VERSION
            log_data["pytools_build_version"] = PYTOOLS_BUILD_VERSION
            log_data["state_snapshot_copy"] = snapshot_result

            state_payload = telemetry_route.load_state(source_folder=source_folder)
            shutdown_state = dict(log_data)
            set_cleanup_version = bool(
                shutdown_state.pop("_set_cleanup_version", False)
            )
            shutdown_state["updated_utc"] = _utc_now()
            state_payload["last_shutdown_transfer"] = shutdown_state
            state_payload["pytools_toolbar_version"] = PYTOOLS_TOOLBAR_VERSION
            state_payload["pytools_build_version"] = PYTOOLS_BUILD_VERSION
            state_payload["last_state_snapshot_copy"] = dict(snapshot_result)

            legacy_cleanup = log_data.get("legacy_cleanup") or {}
            if legacy_cleanup.get("ran") or legacy_cleanup.get("deferred"):
                cleanup_state = dict(legacy_cleanup)
                cleanup_state["required_version"] = LOCAL_TELEMETRY_CLEANUP_VERSION
                cleanup_state["previous_version"] = log_data.get(
                    "cleanup_version_before",
                    0,
                )
                cleanup_state["updated_utc"] = _utc_now()
                state_payload["last_local_telemetry_cleanup"] = cleanup_state

            if set_cleanup_version:
                saved_cleanup_version = _to_int(
                    state_payload.get("local_telemetry_cleanup_version", 0)
                )
                if saved_cleanup_version < LOCAL_TELEMETRY_CLEANUP_VERSION:
                    cleanup_utc = _utc_now()
                    state_payload["local_telemetry_cleanup_version"] = (
                        LOCAL_TELEMETRY_CLEANUP_VERSION
                    )
                    state_payload["local_telemetry_cleanup_utc"] = cleanup_utc

            current_file = log_data.get("current_file") or {}
            sweep = log_data.get("post_migration_sweep") or {}
            files_found = int(bool(
                current_file.get("found") and current_file.get("eligible")
            )) + _to_int(sweep.get("files_found", 0))
            files_copied = int(bool(current_file.get("copied"))) + _to_int(
                sweep.get("files_copied", 0)
            )
            files_failed = int(bool(current_file.get("copy_failed"))) + _to_int(
                sweep.get("files_copy_failed", 0)
            )
            state_payload["last_transfer"] = {
                "status": log_data.get("status", ""),
                "username": log_data.get("username", ""),
                "resolved_root": log_data.get("resolved_route", ""),
                "destination_folder": log_data.get("destination_folder", ""),
                "files_found": files_found,
                "files_copied": files_copied,
                "files_failed": files_failed,
                "files_deleted": int(bool(current_file.get("deleted"))) + _to_int(
                    sweep.get("files_deleted", 0)
                ),
                "files_dropped": int(bool(current_file.get("dropped"))) + _to_int(
                    sweep.get("files_dropped", 0)
                ),
                "note": "Exact-name telemetry transfer; local deletion attempted after each copy attempt.",
                "updated_utc": _utc_now(),
            }
            telemetry_route.save_state(state_payload, source_folder=source_folder)

            if snapshot_result.get("copy_attempted"):
                try:
                    shutil.copyfile(
                        snapshot_result.get("source_path", ""),
                        snapshot_result.get("destination_path", ""),
                    )
                except Exception as snapshot_error:
                    snapshot_result["status"] = "copy_failed"
                    snapshot_result["copied"] = False
                    snapshot_result["error"] = str(snapshot_error)
                    snapshot_result["source_retained"] = os.path.isfile(
                        snapshot_result.get("source_path", "")
                    )
                    snapshot_result["updated_utc"] = _utc_now()
                    log_data["state_snapshot_copy"] = snapshot_result
                    state_payload["last_state_snapshot_copy"] = dict(
                        snapshot_result
                    )
                    state_payload["last_shutdown_transfer"][
                        "state_snapshot_copy"
                    ] = dict(snapshot_result)
                    telemetry_route.save_state(
                        state_payload,
                        source_folder=source_folder,
                    )
            return True, ""
        except Exception as ex:
            if snapshot_result is not None:
                if snapshot_result.get("status") == "success":
                    snapshot_result["status"] = "state_save_failed"
                    snapshot_result["copy_attempted"] = False
                    snapshot_result["copied"] = False
                snapshot_result["error"] = str(ex)
                snapshot_result["updated_utc"] = _utc_now()
                log_data["state_snapshot_copy"] = snapshot_result
            return False, str(ex)

    if hasattr(telemetry_route, "record_transfer_state"):
        try:
            current_file = log_data.get("current_file") or {}
            sweep = log_data.get("post_migration_sweep") or {}
            telemetry_route.record_transfer_state(
                status=log_data.get("status", ""),
                username=log_data.get("username"),
                resolved_root=log_data.get("resolved_route", ""),
                files_found=int(bool(current_file.get("found"))) + _to_int(
                    sweep.get("files_found", 0)
                ),
                files_copied=int(bool(current_file.get("copied"))) + _to_int(
                    sweep.get("files_copied", 0)
                ),
                files_failed=int(bool(current_file.get("copy_failed"))) + _to_int(
                    sweep.get("files_copy_failed", 0)
                ),
                source_folder=source_folder,
                note="Exact-name telemetry transfer; local deletion attempted after each copy attempt.",
            )
            return True, ""
        except Exception as ex:
            return False, str(ex)

    return False, "state_writer_unavailable"


def _on_app_closing(sender, args):
    source_folder = _telemetry_source_folder()
    log_data = {
        "status": "unknown",
        "error": "",
        "username": "",
        "source_folder": source_folder,
        "route_status": "",
        "route_reason": "",
        "resolved_route": "",
        "destination_folder": "",
        "current_file_lookup_error": "",
        "current_file_handled": False,
        "current_file": _empty_current_file_result(),
        "is_last_revit_process": False,
        "process_check_status": "not_run",
        "process_check_error": "",
        "another_revit_process_open": False,
        "cleanup_deferred": False,
        "cleanup_defer_reason": "",
        "cleanup_version_required": LOCAL_TELEMETRY_CLEANUP_VERSION,
        "cleanup_version_before": 0,
        "cleanup_version_after": 0,
        "legacy_cleanup": _empty_legacy_cleanup_result(),
        "post_migration_sweep": _empty_sweep_result(),
        "_set_cleanup_version": False,
    }

    try:
        try:
            username = getpass.getuser()
        except Exception:
            username = os.environ.get("USERNAME", "UnknownUser")
        log_data["username"] = username

        route_info = _resolve_shutdown_destination(username, source_folder)
        log_data["route_status"] = route_info.get("route_status", "")
        log_data["route_reason"] = route_info.get("route_reason", "")
        log_data["resolved_route"] = route_info.get("resolved_root", "")
        log_data["destination_folder"] = route_info.get(
            "destination_folder",
            "",
        )

        current_path = ""
        try:
            current_path = telemetry.get_telemetry_file_path() or ""
        except Exception as ex:
            log_data["current_file_lookup_error"] = str(ex)

        log_data["current_file"] = _transfer_and_delete_telemetry_file(
            current_path,
            log_data["destination_folder"],
        )
        current_file = log_data.get("current_file") or {}
        log_data["current_file_handled"] = bool(
            not log_data.get("current_file_lookup_error")
            and (
                not current_file.get("found")
                or not current_file.get("eligible")
                or current_file.get("copy_attempted")
            )
        )

        process_details = {}
        is_last_process = _is_last_revit_process(diagnostics=process_details)
        log_data["is_last_revit_process"] = is_last_process
        log_data["process_check_status"] = process_details.get("status", "")
        log_data["process_check_error"] = process_details.get("error", "")
        log_data["another_revit_process_open"] = (
            process_details.get("status") == "another_revit_process"
        )

        state_payload = {}
        state_available = False
        if telemetry_route is not None and hasattr(telemetry_route, "load_state"):
            try:
                state_payload = telemetry_route.load_state(
                    source_folder=source_folder,
                )
                state_available = hasattr(telemetry_route, "save_state")
            except Exception as ex:
                log_data["cleanup_defer_reason"] = "state_load_failed"
                log_data["error"] = str(ex)

        cleanup_version = _to_int(
            state_payload.get("local_telemetry_cleanup_version", 0)
        )
        log_data["cleanup_version_before"] = cleanup_version
        log_data["cleanup_version_after"] = cleanup_version

        if not is_last_process:
            defer_reason = process_details.get("status") or "process_check_failed"
            log_data["cleanup_deferred"] = True
            log_data["cleanup_defer_reason"] = defer_reason
            if cleanup_version < LOCAL_TELEMETRY_CLEANUP_VERSION:
                log_data["legacy_cleanup"]["status"] = "deferred"
                log_data["legacy_cleanup"]["deferred"] = True
                log_data["legacy_cleanup"]["defer_reason"] = defer_reason
            else:
                log_data["post_migration_sweep"]["status"] = "deferred"
                log_data["post_migration_sweep"]["deferred"] = True
                log_data["post_migration_sweep"]["defer_reason"] = defer_reason
        elif not log_data.get("current_file_handled"):
            if log_data.get("current_file_lookup_error"):
                current_defer_reason = "current_file_lookup_failed"
            else:
                current_defer_reason = (
                    (log_data.get("current_file") or {}).get("status")
                    or "current_file_not_handled"
                )
            log_data["cleanup_deferred"] = True
            log_data["cleanup_defer_reason"] = current_defer_reason
            if cleanup_version < LOCAL_TELEMETRY_CLEANUP_VERSION:
                log_data["legacy_cleanup"]["status"] = "deferred"
                log_data["legacy_cleanup"]["deferred"] = True
                log_data["legacy_cleanup"]["defer_reason"] = current_defer_reason
            else:
                log_data["post_migration_sweep"]["status"] = "deferred"
                log_data["post_migration_sweep"]["deferred"] = True
                log_data["post_migration_sweep"]["defer_reason"] = (
                    current_defer_reason
                )
        elif not state_available:
            log_data["cleanup_deferred"] = True
            log_data["cleanup_defer_reason"] = "state_unavailable"
            log_data["legacy_cleanup"]["status"] = "deferred"
            log_data["legacy_cleanup"]["deferred"] = True
            log_data["legacy_cleanup"]["defer_reason"] = "state_unavailable"
        elif cleanup_version < LOCAL_TELEMETRY_CLEANUP_VERSION:
            legacy_cleanup = _run_legacy_local_cleanup(source_folder)
            log_data["legacy_cleanup"] = legacy_cleanup
            if legacy_cleanup.get("complete"):
                log_data["_set_cleanup_version"] = True
                log_data["cleanup_version_after"] = (
                    LOCAL_TELEMETRY_CLEANUP_VERSION
                )
        elif not log_data.get("destination_folder"):
            log_data["cleanup_deferred"] = True
            log_data["cleanup_defer_reason"] = "destination_unavailable"
            log_data["post_migration_sweep"]["status"] = "deferred"
            log_data["post_migration_sweep"]["deferred"] = True
            log_data["post_migration_sweep"]["defer_reason"] = (
                "destination_unavailable"
            )
        else:
            log_data["post_migration_sweep"] = _run_post_migration_sweep(
                source_folder,
                log_data["destination_folder"],
                current_path=current_path,
            )

        current_file = log_data.get("current_file") or {}
        legacy_cleanup = log_data.get("legacy_cleanup") or {}
        sweep = log_data.get("post_migration_sweep") or {}
        has_transfer_failure = bool(
            current_file.get("copy_failed")
            or current_file.get("delete_failed")
            or legacy_cleanup.get("status") in (
                "enumeration_failed",
                "verification_failed",
                "partial_failure",
            )
            or sweep.get("status") in ("enumeration_failed", "partial_success")
        )

        if has_transfer_failure or log_data.get("current_file_lookup_error"):
            log_data["status"] = "partial_success"
        elif route_info.get("status") != "ready":
            log_data["status"] = route_info.get("status", "route_unresolved")
        else:
            log_data["status"] = "success"
    except Exception as ex:
        log_data["status"] = "fatal_error"
        log_data["error"] = str(ex)

    state_saved, state_error = _record_shutdown_state(log_data, source_folder)
    log_data["state_saved"] = state_saved
    log_data["state_error"] = state_error
    return log_data

def _register_shutdown_hook():
    logger = script.get_logger()
    if _get_env(ENV_APP_CLOSING_HANDLER_KEY):
        logger.info("ApplicationClosing hook already registered; skipping.")
        return

    try:
        app = __revit__
        if app is None:
            logger.warning("ApplicationClosing hook not registered: UIApplication unavailable.")
            return
        app.ApplicationClosing += _on_app_closing
        _set_env(ENV_APP_CLOSING_HANDLER_KEY, "1")
        logger.info("ApplicationClosing hook registered.")

    except Exception as exc:
        logger.warning("Failed to register ApplicationClosing hook: %s", exc)

def _check_acc_sync():
    if telemetry_route is not None:
        try:
            route_result = telemetry_route.resolve_usage_route(persist=True)
            if route_result.get("status") == "resolved":
                return
            if int(route_result.get("candidate_count", 0) or 0) > 0:
                # Candidates exist; user can resolve manually from ACC Path Resolver.
                return
        except Exception:
            pass
    elif _find_acc_root() is not None:
        return
    from System.Windows import Window, SizeToContent, WindowStartupLocation, Thickness, TextWrapping, HorizontalAlignment
    from System.Windows.Controls import StackPanel, Image, TextBlock, Button, ScrollViewer
    from System.Windows.Media.Imaging import BitmapImage
    from System import Uri, UriKind

    img_dir = os.path.normpath(os.path.join(
        os.path.dirname(os.path.abspath(__file__)), os.pardir,
        "WM Tools.extension", "AE pyTools.Tab", "WM Tools.panel",
        "WM Tools.pulldown", "Load Electrical Content.pushbutton",
    ))
    sync_img = os.path.join(img_dir, "sync_instruction.png")
    explorer_img = os.path.join(img_dir, "file_explorer_instruction.png")

    win = Window()
    win.Title = "ACC Sync Required"
    win.SizeToContent = SizeToContent.Width
    win.Height = 700
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    scroll = ScrollViewer()
    panel = StackPanel()
    panel.Margin = Thickness(15)

    req = TextBlock()
    req.Text = "REQUIRED FOR COOLSYS EMPLOYEES:"
    req.FontSize = 14
    req.FontWeight = __import__("System.Windows", fromlist=["FontWeights"]).FontWeights.Bold
    req.Margin = Thickness(0, 0, 0, 5)
    panel.Children.Add(req)

    header = TextBlock()
    header.Text = "CED Content Collection is not synced"
    header.FontSize = 16
    header.FontWeight = __import__("System.Windows", fromlist=["FontWeights"]).FontWeights.Bold
    header.Margin = Thickness(0, 0, 0, 10)
    panel.Children.Add(header)

    msg = TextBlock()
    msg.TextWrapping = TextWrapping.Wrap
    msg.MaxWidth = 620
    msg.Text = (
        "This extension requires the CED Content Collection ACC project "
        "to be synced via Autodesk Desktop Connector.\n\n"
        "1. Click the Desktop Connector tray icon on your taskbar.\n"
        "2. Click 'Select Projects' and check 'CED Content Collection' "
        "from the CoolSys directory.\n"
        "3. Once synced, restart Revit."
    )
    msg.Margin = Thickness(0, 0, 0, 15)
    panel.Children.Add(msg)

    for img_path, caption, max_w in [(sync_img, "Select Projects in Desktop Connector", 620),
                                      (explorer_img, "ACC folder in File Explorer", 310)]:
        if os.path.exists(img_path):
            lbl = TextBlock()
            lbl.Text = caption
            lbl.FontWeight = __import__("System.Windows", fromlist=["FontWeights"]).FontWeights.SemiBold
            lbl.Margin = Thickness(0, 0, 0, 5)
            panel.Children.Add(lbl)
            img = Image()
            img.Source = BitmapImage(Uri(img_path, UriKind.Absolute))
            img.MaxWidth = max_w
            img.HorizontalAlignment = HorizontalAlignment.Left
            img.Margin = Thickness(0, 0, 0, 15)
            panel.Children.Add(img)

    btn = Button()
    btn.Content = "OK"
    btn.Width = 80
    btn.Height = 28
    btn.HorizontalAlignment = HorizontalAlignment.Center
    btn.Click += lambda s, e: win.Close()
    panel.Children.Add(btn)

    scroll.Content = panel
    win.Content = scroll
    win.ShowDialog()

_configure_pyrevit_telemetry()
_check_acc_sync()
_register_shutdown_hook()
#_register_sync_handler()
# Temporarily disabled to prevent startup-time dockable panel activity.
# _register_place_single_profile_panel()
