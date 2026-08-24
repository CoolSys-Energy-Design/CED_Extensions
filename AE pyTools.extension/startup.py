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
    from pyrevit.userconfig import user_config
except Exception:
    user_config = None
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


# ---------------------------------------------------------------------------
# Revit 2027+ CPython engine primer
#
# On Revit 2027 (netcore / .NET 10) the FIRST CPython engine init of a
# session, if it happens after a document has been opened, can abort
# mid-assembly-scan and leave pythonnet's ``clr`` module without its
# ``_available_namespaces`` table — every ``#! python3`` button then dies
# on its first CLR import for the rest of the session
# (pyrevitlabs/pyRevit #3341 / #3346 / #3510 family).
#
# Initializing the engine HERE — at session load, before any document —
# sidesteps that window. pyRevit's own CPythonEngine.Start() skips init
# when PythonEngine.IsInitialized is already true (and tolerates
# re-setting the same PythonDLL), so this is transparent to normal runs.
#
# A raw embedded init can also come up with no stdlib on sys.path
# ("No module named 'traceback'"), and CPythonEngine.StoreSearchPaths()
# snapshots sys.path on each command's first run — so the warm-up below
# seeds the stdlib zip + engine dir and proves traceback/clr/System/
# Autodesk.Revit.DB all import before any button ever runs.
# ---------------------------------------------------------------------------
_PRIMER_LOG = os.path.join(
    os.environ.get("APPDATA", ""), "pyRevit", "CED_cpython_primer.log")


def _primer_note(msg):
    """File breadcrumb — pyRevit 6.5.4 writes no 2027 runtime log."""
    try:
        with open(_PRIMER_LOG, "a") as f:
            f.write("%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


def _prime_cpython_engine():
    logger = script.get_logger()
    try:
        import re as _re
        from pyrevit import HOST_APP
        import pyrevit as _pyrevit
        from System import AppDomain, Type, Array, Object
        from System.Reflection import Assembly

        try:
            if int(str(HOST_APP.version)) < 2027:
                return
        except Exception:
            return
        _primer_note("--- session start (Revit %s, pid %s) ---" % (
            HOST_APP.version, os.getpid()))

        pn = None
        bin_dir = None
        for _asm in AppDomain.CurrentDomain.GetAssemblies():
            try:
                name = _asm.GetName().Name
            except Exception:
                continue
            if name == "pyRevitLabs.PythonNet":
                pn = _asm
            elif bin_dir is None and name.startswith("pyRevitLabs.PyRevit"):
                try:
                    loc = _asm.Location
                    if loc:
                        # ...\bin\<netcore|netfx>\pyRevitLabs.*.dll
                        bin_dir = os.path.dirname(loc)
                except Exception:
                    pass
        if pn is None:
            # Assembly loading is lazy — at startup time nothing has
            # touched the CPython types yet, so load it ourselves.
            candidates = []
            if bin_dir:
                candidates.append(
                    os.path.join(bin_dir, "pyRevitLabs.PythonNet.dll"))
            candidates.append(os.path.join(
                _pyrevit.HOME_DIR, "bin", "netcore",
                "pyRevitLabs.PythonNet.dll"))
            for cand in candidates:
                if os.path.isfile(cand):
                    try:
                        pn = Assembly.LoadFrom(cand)
                        _primer_note("loaded PythonNet from %s" % cand)
                        break
                    except Exception as load_exc:
                        _primer_note(
                            "LoadFrom failed for %s: %s" % (cand, load_exc))
        if pn is None:
            _primer_note("ABORT: pyRevitLabs.PythonNet not loadable")
            return
        engine_t = pn.GetType("Python.Runtime.PythonEngine")
        runtime_t = pn.GetType("Python.Runtime.Runtime")
        py_t = pn.GetType("Python.Runtime.Py")
        if engine_t is None or runtime_t is None or py_t is None:
            _primer_note("ABORT: Python.Runtime types missing")
            return

        # newest CPython engine bundled with this clone
        cengines = os.path.join(_pyrevit.HOME_DIR, "bin", "cengines")
        if not os.path.isdir(cengines):
            _primer_note("ABORT: no cengines dir at %s" % cengines)
            return
        engine_dirs = sorted(
            d for d in os.listdir(cengines)
            if _re.match(r"^CPY\d+$", d)
            and os.path.isdir(os.path.join(cengines, d))
        )
        if not engine_dirs:
            _primer_note("ABORT: no CPY* engine dirs in %s" % cengines)
            return
        engine_dir = os.path.join(cengines, engine_dirs[-1])
        dlls = sorted(
            f for f in os.listdir(engine_dir)
            if _re.match(r"^python3\d+\.dll$", f.lower())
        )
        if not dlls:
            _primer_note("ABORT: no python3*.dll in %s" % engine_dir)
            return
        python_dll = os.path.join(engine_dir, dlls[-1])
        stdlib_zip = os.path.splitext(python_dll)[0] + ".zip"

        if not engine_t.GetProperty("IsInitialized").GetValue(None, None):
            runtime_t.GetProperty("PythonDLL").SetValue(None, python_dll, None)
            engine_t.GetProperty("ProgramName").SetValue(None, "pyrevit", None)
            engine_t.GetMethod("Initialize", Type.EmptyTypes).Invoke(None, None)
            _primer_note("engine initialized (dll=%s)" % python_dll)
        else:
            _primer_note("engine was already initialized")

        # CRITICAL: pythonnet's Runtime.PythonDLL setter throws
        # unconditionally once Runtime._isInitialized is true — and
        # CPythonEngine.Start() assigns PythonDLL on every fresh engine
        # instance, so pre-initializing here would crash the FIRST run
        # of every command ("This property must be set before runtime is
        # initialized"), caching an engine with empty _sysPaths (-> "No
        # module named 'configparser'" on the second run). Flipping the
        # private guard field back to False lets Start() assign the (now
        # inert) property and fall through to StoreSearchPaths(), which
        # snapshots our seeded sys.path. PythonEngine.IsInitialized is a
        # separate flag and stays True, so Start() still skips its own
        # Initialize() call.
        from System.Reflection import BindingFlags
        init_fld = runtime_t.GetField(
            "_isInitialized", BindingFlags.NonPublic | BindingFlags.Static)
        if init_fld is not None:
            init_fld.SetValue(None, False)
            _primer_note("Runtime._isInitialized flipped to False "
                         "(defuses PythonDLL setter for first Start)")
        else:
            _primer_note("WARN: Runtime._isInitialized field not found — "
                         "first click per session may fail once")

        snippet = (
            "import sys\n"
            "for _p in (r'{zip}', r'{dir}'):\n"
            "    if _p not in sys.path:\n"
            "        sys.path.insert(0, _p)\n"
            "import traceback\n"
            "import clr\n"
            "import System\n"
            "import Autodesk.Revit.DB\n"
        ).format(zip=stdlib_zip, dir=engine_dir)
        gil = py_t.GetMethod("GIL").Invoke(None, None)
        try:
            rc = engine_t.GetMethod("RunSimpleString").Invoke(
                None, Array[Object]([snippet]))
        finally:
            gil.Dispose()
        if rc == 0:
            _primer_note("OK: stdlib seeded + warm-up imports passed")
            logger.info("CPython engine primed (dll=%s)", python_dll)
        else:
            _primer_note("WARM-UP FAILED (rc=%s)" % rc)
            logger.warning(
                "CPython engine primer warm-up failed (rc=%s) — "
                "python3 buttons may not work this session", rc)
    except Exception as exc:
        _primer_note("EXCEPTION: %s" % exc)
        try:
            logger.warning("CPython engine primer failed: %s", exc)
        except Exception:
            pass


try:
    _prime_cpython_engine()
except Exception:
    pass


_SYNC_HANDLER_UI = None
_SYNC_HANDLER_APP = None
_MODULE = None
_IS_RUNNING = False

_DOCKABLE_REGISTERED = False

# Emergency release switches. Keep both False in distributed builds.
# Set both True only in a local test build after telemetry is deliberately
# re-enabled. The disabled path changes only pyRevit's active state; it never
# assigns, creates, clears, or transfers a telemetry file or folder.
ENABLE_PYREVIT_TELEMETRY = False
ENABLE_DESKTOP_CONNECTOR_TELEMETRY_TRANSFER = False


def _telemetry_state_is_enabled(value):
    return str(value or "").strip().lower() in (
        "true",
        "1",
        "yes",
        "on",
    )


def _set_pyrevit_telemetry_active(enabled):
    """Set only pyRevit's global active setting when it differs."""
    logger = script.get_logger()
    config_enabled = None
    runtime_enabled = None

    if user_config is not None:
        try:
            config_enabled = _telemetry_state_is_enabled(
                user_config.telemetry_status
            )
        except Exception as exc:
            logger.warning("Could not read pyRevit telemetry config: %s", exc)

    try:
        runtime_enabled = _telemetry_state_is_enabled(
            telemetry.get_telemetry_state()
        )
    except Exception as exc:
        logger.warning("Could not read pyRevit telemetry state: %s", exc)

    if config_enabled == enabled and runtime_enabled == enabled:
        logger.info("pyRevit telemetry active state already %s.", enabled)
        return

    try:
        telemetry.set_telemetry_state(enabled)
        config_saved = False
        if user_config is not None and config_enabled != enabled:
            user_config.save_changes()
            config_saved = True
        logger.info(
            "pyRevit telemetry active state set to %s. config_saved=%s",
            enabled,
            config_saved,
        )
    except Exception as exc:
        logger.warning("Failed to set pyRevit telemetry active state: %s", exc)


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


def _event_flags_to_int(value):
    if value in (None, ""):
        return 0
    try:
        value_text = str(value).strip()
        if value_text.lower().startswith("0x"):
            return int(value_text, 16)
        return int(value_text)
    except Exception:
        return 0


def _configure_pyrevit_telemetry():
    if not ENABLE_PYREVIT_TELEMETRY:
        # Release behavior: turn telemetry off without touching any telemetry
        # path or invoking pyRevit setup, which owns session-file creation.
        _set_pyrevit_telemetry_active(False)
        return

    # Local test behavior: restore the retained telemetry setup below.
    _set_pyrevit_telemetry_active(True)
    logger = script.get_logger()
    source_folder, folder_ok, folder_error = _ensure_telemetry_source_folder()
    if not folder_ok:
        logger.warning("Telemetry folder not available: %s", folder_error)
        return

    try:
        telemetry_cfg = script.get_config("telemetry")

        expected_settings = {
            "utc_timestamps": True,
            "active": True,
            "telemetry_file_dir": source_folder,
            "telemetry_server_url": "",
            "include_hooks": True,
            "active_app": False,
            "apptelemetry_server_url": "",
            "apptelemetry_event_flags": "0x0",
        }

        current_settings = {
            setting_name: telemetry_cfg.get_option(setting_name, default_value="")
            for setting_name in expected_settings
        }

        setting_setters = {
            "utc_timestamps": telemetry.set_telemetry_utc_timestamp,
            "active": telemetry.set_telemetry_state,
            "telemetry_file_dir": telemetry.set_telemetry_file_dir,
            "telemetry_server_url": telemetry.set_telemetry_server_url,
            "include_hooks": telemetry.set_telemetry_include_hooks,
            "active_app": telemetry.set_apptelemetry_state,
            "apptelemetry_server_url": telemetry.set_apptelemetry_server_url,
            "apptelemetry_event_flags": lambda _: telemetry.set_apptelemetry_event_flags(0),
        }

        value_normalizers = {
            "telemetry_file_dir": _normalize_path,
            "apptelemetry_event_flags": _event_flags_to_int,
        }

        changed_settings = []
        for setting_name, expected_value in expected_settings.items():
            current_value = current_settings.get(setting_name)
            normalizer = value_normalizers.get(setting_name)
            if normalizer:
                current_value = normalizer(current_value)
                expected_value = normalizer(expected_value)
            if current_value != expected_value:
                setting_setters[setting_name](expected_settings[setting_name])
                changed_settings.append(setting_name)

        if changed_settings:
            # setup_telemetry() applies derived runtime state (session file path,
            # handlers, env vars) and persists the updated config once.
            telemetry.setup_telemetry()
            logger.info(
                "pyRevit telemetry updated via telemetry API. changed=%s file_dir=%s",
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


def _nonclobber_path(dst_path):
    if not os.path.exists(dst_path):
        return dst_path
    base, ext = os.path.splitext(dst_path)
    tick = int(time.time())
    candidate = "{}_{}{}".format(base, tick, ext)
    if not os.path.exists(candidate):
        return candidate
    index = 1
    while True:
        candidate = "{}_{}_{}{}".format(base, tick, index, ext)
        if not os.path.exists(candidate):
            return candidate
        index += 1


def _on_app_closing(sender, args):
    if not ENABLE_DESKTOP_CONNECTOR_TELEMETRY_TRANSFER:
        return {
            "status": "disabled_by_release_switch",
            "transfer_disabled": True,
        }

    log_data = {
        "username": None,
        "files_found": 0,
        "files_copied": 0,
        "files_failed": 0,
        "status": "unknown",
        "error": None,
        "route_status": None,
        "route_reason": None,
        "route_root": None,
        "recovery_status": None,
        "recovery_error": None,
    }

    try:
        # Username
        try:
            username = getpass.getuser()
        except Exception:
            username = os.environ.get("USERNAME", "UnknownUser")

        log_data["username"] = username

        source_folder = _telemetry_source_folder()
        if not os.path.exists(source_folder):
            log_data["status"] = "no_source_folder"
            return

        acc_root = None
        if telemetry_route is not None:
            try:
                route_result = telemetry_route.resolve_usage_route(username=username, persist=True)
                log_data["route_status"] = route_result.get("status")
                log_data["route_reason"] = route_result.get("reason")
                acc_root = route_result.get("resolved_root")
            except Exception as ex:
                log_data["route_status"] = "error"
                log_data["route_reason"] = str(ex)

        if not acc_root:
            acc_root = _find_acc_root()

        log_data["route_root"] = acc_root
        if acc_root is None:
            log_data["status"] = "route_unresolved"
            if telemetry_route is not None:
                telemetry_route.record_transfer_state(
                    status="route_unresolved",
                    username=username,
                    resolved_root="",
                    files_found=0,
                    files_copied=0,
                    files_failed=0,
                    source_folder=source_folder,
                    note=log_data.get("route_reason") or "No ACC root resolved.",
                )
            return

        base_path = os.path.join(acc_root, "Project Files", "03 Automations", "Usage")
        if not os.path.isdir(base_path):
            log_data["status"] = "usage_base_missing"
            if telemetry_route is not None:
                telemetry_route.record_transfer_state(
                    status="usage_base_missing",
                    username=username,
                    resolved_root=acc_root,
                    files_found=0,
                    files_copied=0,
                    files_failed=0,
                    source_folder=source_folder,
                    note="Usage base folder not found. Transfer canceled.",
                )
            return

        try:
            if telemetry_route is not None and hasattr(telemetry_route, "ensure_user_folder"):
                folder_result = telemetry_route.ensure_user_folder(acc_root, username=username)
                user_folder = folder_result.get("path") or os.path.join(base_path, username)
                if not folder_result.get("ok"):
                    raise Exception(folder_result.get("reason", "user_folder_unavailable"))
            else:
                user_folder = os.path.join(base_path, username)
                if not os.path.exists(user_folder):
                    # Intentionally create only the username folder under an existing Usage base.
                    os.mkdir(user_folder)
        except Exception as e:
            log_data["status"] = "failed_create_user_folder"
            log_data["error"] = str(e)
            if telemetry_route is not None:
                telemetry_route.record_transfer_state(
                    status="failed_create_user_folder",
                    username=username,
                    resolved_root=acc_root,
                    files_found=0,
                    files_copied=0,
                    files_failed=0,
                    source_folder=source_folder,
                    note=str(e),
                )
            # from Snippets import hooks_logger
            # hooks_logger.log_hook(__file__, log_data)
            return

        recovery_result = None
        if telemetry_route is not None and hasattr(telemetry_route, "recover_stale_usage_jsons"):
            try:
                recovery_result = telemetry_route.recover_stale_usage_jsons(
                    acc_root,
                    username=username,
                    source_folder=source_folder,
                )
                log_data["recovery_status"] = recovery_result.get("status")
                if hasattr(telemetry_route, "record_recovery_state"):
                    telemetry_route.record_recovery_state(
                        recovery_result,
                        source_folder=source_folder,
                    )
            except Exception as e:
                recovery_result = {
                    "status": "error",
                    "username": username,
                    "resolved_root": acc_root,
                    "destination_folder": user_folder,
                    "stale_folders_checked": [],
                    "files_found": 0,
                    "files_moved": 0,
                    "files_failed": 0,
                    "error": str(e),
                }
                log_data["recovery_status"] = "error"
                log_data["recovery_error"] = str(e)
                if hasattr(telemetry_route, "record_recovery_state"):
                    try:
                        telemetry_route.record_recovery_state(
                            recovery_result,
                            source_folder=source_folder,
                        )
                    except Exception:
                        pass

        files = os.listdir(source_folder)
        log_data["files_found"] = len(files)

        for fname in files:
            try:
                src = os.path.join(source_folder, fname)

                if not os.path.isfile(src):
                    continue

                if telemetry_route is not None and fname == telemetry_route.STATE_FILE_NAME:
                    # Keep route status local; never transfer it to ACC.
                    continue

                dst = _nonclobber_path(os.path.join(user_folder, fname))
                # Copy instead of move so local telemetry remains available.
                shutil.copy2(src, dst)
                log_data["files_copied"] += 1

            except Exception:
                log_data["files_failed"] += 1

        if log_data["files_failed"] > 0:
            log_data["status"] = "partial_success"
        else:
            log_data["status"] = "success"

        if telemetry_route is not None:
            telemetry_route.record_transfer_state(
                status=log_data["status"],
                username=username,
                resolved_root=acc_root,
                files_found=log_data["files_found"],
                files_copied=log_data["files_copied"],
                files_failed=log_data["files_failed"],
                source_folder=source_folder,
                note="Copied telemetry files to ACC; local files retained.",
            )

    except Exception as e:
        log_data["status"] = "fatal_error"
        log_data["error"] = str(e)
        if telemetry_route is not None:
            try:
                telemetry_route.record_transfer_state(
                    status="fatal_error",
                    username=log_data.get("username"),
                    resolved_root=log_data.get("route_root") or "",
                    files_found=log_data.get("files_found", 0),
                    files_copied=log_data.get("files_copied", 0),
                    files_failed=log_data.get("files_failed", 0),
                    source_folder=_telemetry_source_folder(),
                    note=str(e),
                )
            except Exception:
                pass

    # Always log
    # try:
    #     from Snippets import hooks_logger
    #     hooks_logger.log_hook(__file__, log_data)
    # except:
    #     pass

def _register_shutdown_hook():
    logger = script.get_logger()
    if not ENABLE_DESKTOP_CONNECTOR_TELEMETRY_TRANSFER:
        logger.info(
            "Desktop Connector telemetry transfer disabled; "
            "ApplicationClosing hook not registered."
        )
        return
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
