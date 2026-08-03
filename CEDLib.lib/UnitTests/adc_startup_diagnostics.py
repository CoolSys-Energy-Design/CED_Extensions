# -*- coding: utf-8 -*-
r"""Standalone ADC routing, startup, and diagnostics tests.

Run tests outside Revit:
    py -3 C:\Users\Aevelina\CED_Extensions\CEDLib.lib\UnitTests\adc_startup_diagnostics.py

Print live, read-only route diagnostics outside Revit:
    py -3 C:\Users\Aevelina\CED_Extensions\CEDLib.lib\UnitTests\adc_startup_diagnostics.py --live
"""

from __future__ import print_function

import importlib.util
import json
import os
import runpy
import sys
import tempfile
import types
import unittest
import warnings
from unittest import mock

TEST_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(TEST_DIR, os.pardir, os.pardir))
EXTENSION_ROOT = os.path.join(REPO_ROOT, "AE pyTools.extension")
STARTUP_PATH = os.path.join(EXTENSION_ROOT, "startup.py")
TELEMETRY_ROUTE_PATH = os.path.join(EXTENSION_ROOT, "telemetry_route.py")
DIAGNOSTICS_PATH = os.path.join(
    EXTENSION_ROOT,
    "AE pyTools.Tab",
    "CED Tools.panel",
    "ADC Diagnostics.pushbutton",
    "script.py",
)


def _load_module(module_name, path):
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _make_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def _write_text(path, value):
    _make_dir(os.path.dirname(path))
    with open(path, "w", encoding="utf-8") as stream:
        stream.write(value)


def _read_text(path):
    with open(path, "r", encoding="utf-8") as stream:
        return stream.read()


def _read_json(path):
    return json.loads(_read_text(path))


def _make_candidate(
    base_dir,
    anchor_name,
    project_folders,
    usage_users,
    include_route_key=False,
):
    root = os.path.join(
        base_dir,
        anchor_name,
        "ACCDocs",
        "CoolSys",
        "CED Content Collection",
    )
    project_files = _make_dir(os.path.join(root, "Project Files"))
    for folder_name in project_folders:
        _make_dir(os.path.join(project_files, folder_name))

    usage_base = _make_dir(
        os.path.join(project_files, "03 Automations", "Usage")
    )
    for username in usage_users:
        _make_dir(os.path.join(usage_base, username))

    if include_route_key:
        route_key = os.path.join(
            project_files,
            "py",
            ".ced_usage_route_key.json",
        )
        _write_text(
            route_key,
            json.dumps(
                {
                    "route_id": "ced-telemetry-usage-v1",
                    "schema_version": 1,
                }
            ),
        )
    return root


class _FakeLogger(object):
    def __init__(self):
        self.messages = []

    def _record(self, level, message, *args):
        if args:
            try:
                message = message % args
            except Exception:
                message = "{} {}".format(message, args)
        self.messages.append((level, str(message)))

    def info(self, message, *args):
        self._record("info", message, *args)

    def warning(self, message, *args):
        self._record("warning", message, *args)

    def exception(self, message, *args):
        self._record("exception", message, *args)


class _FakeOutput(object):
    def __init__(self):
        self.lines = []

    def print_md(self, value):
        self.lines.append(str(value))


class _FakeConfig(object):
    def __init__(self, options=None):
        self.options = dict(options or {})

    def get_option(self, name, default_value=""):
        return self.options.get(name, default_value)


class _FakeTelemetry(types.ModuleType):
    def __init__(self, options=None):
        types.ModuleType.__init__(self, "pyrevit.telemetry")
        options = dict(options or {})
        self.calls = []
        self.setup_count = 0
        self.utc_timestamps = options.get("utc_timestamps", False)
        self.active = options.get("active", False)
        self.telemetry_file_dir = options.get("telemetry_file_dir", "")
        self.include_hooks = options.get("include_hooks", False)
        self.telemetry_file_path = options.get("telemetry_file_path", "")

    def _record(self, name, value):
        self.calls.append((name, value))

    def set_telemetry_utc_timestamp(self, value):
        self.utc_timestamps = value
        self._record("utc_timestamps", value)

    def set_telemetry_state(self, value):
        self.active = value
        self._record("active", value)

    def set_telemetry_file_dir(self, value):
        self.telemetry_file_dir = value
        self._record("telemetry_file_dir", value)

    def set_telemetry_server_url(self, value):
        self._record("telemetry_server_url", value)

    def set_telemetry_include_hooks(self, value):
        self.include_hooks = value
        self._record("include_hooks", value)

    def set_apptelemetry_state(self, value):
        self._record("active_app", value)

    def set_apptelemetry_server_url(self, value):
        self._record("apptelemetry_server_url", value)

    def set_apptelemetry_event_flags(self, value):
        self._record("apptelemetry_event_flags", value)

    def setup_telemetry(self):
        self.setup_count += 1
        if self.active and os.path.isdir(self.telemetry_file_dir):
            self.telemetry_file_path = os.path.join(
                self.telemetry_file_dir,
                "pyRevit_test_session_telemetry.json",
            )

    def get_telemetry_utc_timestamp(self):
        return self.utc_timestamps

    def get_telemetry_state(self):
        return self.active

    def get_telemetry_file_dir(self):
        return self.telemetry_file_dir

    def get_telemetry_include_hooks(self):
        return self.include_hooks

    def get_telemetry_file_path(self):
        return self.telemetry_file_path


class _FakeForms(types.ModuleType):
    def __init__(self):
        types.ModuleType.__init__(self, "pyrevit.forms")
        self.alerts = []

    def alert(self, message, **kwargs):
        self.alerts.append((str(message), dict(kwargs)))
        return False

    def pick_folder(self, **kwargs):
        return None

    def is_registered_dockable_panel(self, panel_cls):
        return False

    def register_dockable_panel(self, panel_cls, default_visible=False):
        return None


class _PyRevitShim(object):
    def __init__(self, source_folder, config_options=None):
        self.logger = _FakeLogger()
        self.output = _FakeOutput()
        self.config = _FakeConfig(config_options)
        self.telemetry = _FakeTelemetry(config_options)
        self.forms = _FakeForms()

        self.script = types.ModuleType("pyrevit.script")
        self.script.get_logger = lambda: self.logger
        self.script.get_output = lambda: self.output
        self.script.get_config = lambda section=None: self.config
        self.script.get_envvar = lambda name: None
        self.script.set_envvar = lambda name, value: None

        self.pyrevit = types.ModuleType("pyrevit")
        self.pyrevit.forms = self.forms
        self.pyrevit.script = self.script
        self.pyrevit.telemetry = self.telemetry

        self.clr = types.ModuleType("clr")
        self.clr.AddReference = lambda name: None

        self.route = types.ModuleType("telemetry_route")
        self.route.STATE_FILE_NAME = ".ced_usage_route_status.json"
        self.route.telemetry_source_folder = lambda: source_folder
        self.route.ensure_telemetry_source_folder = (
            lambda: (source_folder, True, None)
        )
        self.route.resolve_usage_route = lambda **kwargs: {
            "status": "resolved",
            "reason": "test",
            "resolved_root": "",
            "candidate_count": 1,
        }
        self.route.record_transfer_state = lambda **kwargs: None
        self.route.ensure_user_folder = lambda resolved_root, username=None: {
            "ok": True,
            "created": False,
            "reason": "already_exists",
            "path": os.path.join(
                resolved_root,
                "Project Files",
                "03 Automations",
                "Usage",
                username or "UnknownUser",
            ),
        }
        self.route.recover_stale_usage_jsons = lambda *args, **kwargs: {
            "status": "no_stale_jsons",
            "username": kwargs.get("username", ""),
            "resolved_root": args[0] if args else "",
            "destination_folder": "",
            "stale_folders_checked": [],
            "files_found": 0,
            "files_moved": 0,
            "files_failed": 0,
            "files_skipped_existing": 0,
            "error": "",
        }
        self.route.record_recovery_state = lambda *args, **kwargs: None

    def modules(self):
        return {
            "clr": self.clr,
            "pyrevit": self.pyrevit,
            "pyrevit.forms": self.forms,
            "pyrevit.script": self.script,
            "pyrevit.telemetry": self.telemetry,
            "telemetry_route": self.route,
        }


def _load_startup_with_shims(source_folder, config_options=None):
    shim = _PyRevitShim(source_folder, config_options=config_options)
    module_name = "ced_startup_test_{}".format(id(shim))
    with mock.patch.dict(sys.modules, shim.modules(), clear=False):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", DeprecationWarning)
            startup = _load_module(module_name, STARTUP_PATH)
    return startup, shim


def print_live_diagnostics():
    """Print the current machine's routing diagnostics without Revit."""
    route = _load_module(
        "ced_telemetry_route_live",
        TELEMETRY_ROUTE_PATH,
    )
    username = route.get_username()
    result = route.resolve_usage_route(
        username=username,
        persist=False,
    )
    state = route.load_state()

    print("CED Telemetry Route Diagnostics")
    print("=" * 31)
    print("Mode: read-only")
    print("Username: {}".format(username))
    print("State file: {}".format(result.get("state_file", "")))
    print("Approved root: {}".format(state.get("approved_root", "")))
    print("Status: {}".format(result.get("status", "")))
    print("Reason: {}".format(result.get("reason", "")))
    print("Resolved root: {}".format(result.get("resolved_root", "")))
    print(
        "Score / margin: {} / {}".format(
            result.get("best_score", 0),
            result.get("margin", 0),
        )
    )
    print(
        "Candidates / viable: {} / {}".format(
            result.get("candidate_count", 0),
            result.get("viable_count", 0),
        )
    )
    print("")

    candidates = list(result.get("scored_candidates") or [])
    if not candidates:
        print("No candidates found.")
        return result

    print("Candidates")
    print("-" * 10)
    for index, item in enumerate(candidates, 1):
        print(
            "{}. score={} root={}".format(
                index,
                item.get("score", 0),
                item.get("root", ""),
            )
        )
        print(
            "   exists={} usage_base={} usage_users={} project_folders={}".format(
                bool(item.get("root_exists")),
                bool(item.get("usage_base_exists")),
                item.get("usage_subfolder_count", 0),
                item.get("project_files_subfolder_count", 0),
            )
        )
        print(
            "   only_automations={} only_current_user={} key_match={}".format(
                bool(item.get("project_files_only_automations")),
                bool(item.get("usage_only_current_user")),
                bool(
                    (item.get("key_info") or {}).get(
                        "matches_expected_route"
                    )
                ),
            )
        )
    return result


class LiveRouteDiagnosticsTests(unittest.TestCase):
    def test_current_machine_route_report(self):
        print("")
        result = print_live_diagnostics()
        print("")

        self.assertIn(
            result.get("status"),
            ("resolved", "ambiguous", "not_found"),
        )
        self.assertIn("candidate_count", result)
        self.assertIn("scored_candidates", result)


class TelemetryRouteTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.route = _load_module(
            "ced_telemetry_route_test",
            TELEMETRY_ROUTE_PATH,
        )

    def test_full_tree_candidate_beats_partial_duplicate(self):
        with tempfile.TemporaryDirectory(prefix="ced_adc_route_") as temp_dir:
            full_root = _make_candidate(
                temp_dir,
                "ACC",
                [
                    "01 Projects",
                    "02 Standards",
                    "03 Automations",
                    "04 Resources",
                    "05 Archive",
                ],
                ["Alpha", "Beta", "Gamma"],
                include_route_key=True,
            )
            partial_root = _make_candidate(
                temp_dir,
                "DC",
                ["03 Automations"],
                ["TestUser"],
                include_route_key=False,
            )
            state_folder = _make_dir(os.path.join(temp_dir, "state"))

            with mock.patch.object(
                self.route,
                "build_candidate_roots",
                return_value=[partial_root, full_root],
            ):
                result = self.route.resolve_usage_route(
                    username="TestUser",
                    source_folder=state_folder,
                    persist=False,
                )

            self.assertEqual("resolved", result["status"])
            self.assertEqual(
                self.route._norm(full_root),
                result["resolved_root"],
            )
            self.assertEqual(2, result["candidate_count"])
            self.assertTrue(
                result["scored_candidates"][0]["key_info"][
                    "matches_expected_route"
                ]
            )
            partial = [
                item
                for item in result["scored_candidates"]
                if item["root"] == self.route._norm(partial_root)
            ][0]
            self.assertTrue(partial["project_files_only_automations"])
            self.assertTrue(partial["usage_only_current_user"])

    def test_discovery_scans_drive_root_and_home_root(self):
        home_root = r"C:\Users\TestUser"
        calls = []

        def fake_top_level_anchor_dirs(scope_root):
            calls.append(scope_root)
            if scope_root == home_root:
                return ["home_anchor", "shared_anchor"]
            return ["drive_anchor", "shared_anchor"]

        def fake_discover_project_roots(anchor):
            if anchor == "drive_anchor":
                return ["drive_candidate", "duplicate_candidate"]
            if anchor == "shared_anchor":
                return ["shared_candidate"]
            if anchor == "home_anchor":
                return ["home_candidate", "duplicate_candidate"]
            return []

        with mock.patch.object(
            self.route.os.path,
            "expanduser",
            return_value=home_root,
        ), mock.patch.object(
            self.route,
            "_top_level_anchor_dirs",
            side_effect=fake_top_level_anchor_dirs,
        ), mock.patch.object(
            self.route,
            "_discover_project_roots_under",
            side_effect=fake_discover_project_roots,
        ):
            candidates = self.route._discover_candidates_by_scope()

        self.assertEqual(2, len(calls))
        self.assertEqual(home_root, calls[1])
        self.assertEqual(
            [
                "drive_candidate",
                "duplicate_candidate",
                "shared_candidate",
                "home_candidate",
            ],
            candidates,
        )

    def test_recover_stale_usage_jsons_moves_unique_and_skips_collision(self):
        with tempfile.TemporaryDirectory(prefix="ced_adc_recover_") as temp_dir:
            correct_root = _make_candidate(
                temp_dir,
                "ACC",
                ["01 Projects", "03 Automations", "04 Resources"],
                ["ExistingUser"],
                include_route_key=True,
            )
            stale_root = _make_candidate(
                temp_dir,
                "DC",
                ["03 Automations"],
                ["TestUser"],
                include_route_key=False,
            )

            correct_user = os.path.join(
                correct_root,
                "Project Files",
                "03 Automations",
                "Usage",
                "TestUser",
            )
            stale_user = os.path.join(
                stale_root,
                "Project Files",
                "03 Automations",
                "Usage",
                "TestUser",
            )

            _make_dir(correct_user)
            unique_name = "unique_telemetry.json"
            collision_name = "collision_TELEMETRY.JSON"
            unique_source = os.path.join(stale_user, unique_name)
            collision_source = os.path.join(stale_user, collision_name)
            collision_destination = os.path.join(
                correct_user,
                collision_name,
            )
            _write_text(unique_source, '{"unique": true}')
            _write_text(collision_source, '{"stale": true}')
            _write_text(collision_destination, '{"approved": true}')
            _write_text(
                os.path.join(stale_user, "unrelated.json"),
                '{"unrelated": true}',
            )
            _write_text(os.path.join(stale_user, "notes.txt"), "leave this")
            _write_text(
                os.path.join(stale_user, self.route.STATE_FILE_NAME),
                '{"state": true}',
            )

            with mock.patch.object(
                self.route,
                "build_candidate_roots",
                return_value=[correct_root, stale_root],
            ):
                result = self.route.recover_stale_usage_jsons(
                    correct_root,
                    username="TestUser",
                    source_folder=os.path.join(temp_dir, "state"),
                )

            self.assertEqual("success", result["status"])
            self.assertEqual(2, result["files_found"])
            self.assertEqual(1, result["files_moved"])
            self.assertEqual(0, result["files_failed"])
            self.assertEqual(1, result["files_skipped_existing"])
            self.assertTrue(os.path.isfile(os.path.join(correct_user, unique_name)))
            self.assertFalse(os.path.isfile(unique_source))
            self.assertTrue(os.path.isfile(collision_source))
            self.assertEqual(
                '{"approved": true}',
                _read_text(collision_destination),
            )
            self.assertEqual(
                [collision_name, unique_name],
                sorted(
                    name for name in os.listdir(correct_user)
                    if name.lower().endswith("_telemetry.json")
                ),
            )
            self.assertTrue(
                os.path.isfile(os.path.join(stale_user, "unrelated.json"))
            )
            self.assertTrue(os.path.isfile(os.path.join(stale_user, "notes.txt")))
            self.assertTrue(os.path.isfile(os.path.join(stale_user, self.route.STATE_FILE_NAME)))
            self.assertTrue(os.path.isdir(stale_user))

            state_folder = _make_dir(os.path.join(temp_dir, "state"))
            self.route.record_recovery_state(
                result,
                source_folder=state_folder,
            )
            recovery_state = self.route.load_state(
                source_folder=state_folder,
            )["last_stale_recovery"]
            self.assertEqual(1, recovery_state["files_moved"])
            self.assertEqual(0, recovery_state["files_failed"])
            self.assertEqual(1, recovery_state["files_skipped_existing"])

    def test_manual_approval_and_user_folder_are_persisted(self):
        with tempfile.TemporaryDirectory(prefix="ced_adc_manual_") as temp_dir:
            root = _make_candidate(
                temp_dir,
                "ACC",
                ["01 Projects", "03 Automations", "04 Resources"],
                ["ExistingUser"],
                include_route_key=True,
            )
            state_folder = _make_dir(os.path.join(temp_dir, "state"))

            with mock.patch.object(
                self.route,
                "build_candidate_roots",
                return_value=[root],
            ):
                approval = self.route.set_manual_approved_root(
                    root,
                    username="TestUser",
                    source_folder=state_folder,
                )

            self.assertTrue(approval["success"])
            state = self.route.load_state(source_folder=state_folder)
            self.assertEqual(self.route._norm(root), state["approved_root"])

            folder_result = self.route.ensure_user_folder(
                root,
                username="TestUser",
            )
            self.assertTrue(folder_result["ok"])
            self.assertTrue(folder_result["created"])
            self.assertTrue(os.path.isdir(folder_result["path"]))


class StartupTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir_obj = tempfile.TemporaryDirectory(
            prefix="ced_adc_startup_"
        )
        self.temp_dir = self.temp_dir_obj.name
        self.source_folder = _make_dir(
            os.path.join(self.temp_dir, "telemetry")
        )
        matched_config = {
            "utc_timestamps": True,
            "active": True,
            "telemetry_file_dir": self.source_folder,
            "telemetry_file_path": os.path.join(
                self.source_folder,
                "pyRevit_existing_session_telemetry.json",
            ),
            "telemetry_server_url": "",
            "include_hooks": True,
            "active_app": False,
            "apptelemetry_server_url": "",
            "apptelemetry_event_flags": "0x0",
        }
        self.startup, self.shim = _load_startup_with_shims(
            self.source_folder,
            config_options=matched_config,
        )
        self.route = _load_module(
            "ced_telemetry_route_startup_test_{}".format(id(self)),
            TELEMETRY_ROUTE_PATH,
        )
        self.route.telemetry_source_folder = lambda: self.source_folder
        self.route.ensure_telemetry_source_folder = lambda: (
            self.source_folder,
            True,
            None,
        )
        self.startup.telemetry_route = self.route
        self.acc_root = ""
        self.destination_user = ""

    def tearDown(self):
        self.temp_dir_obj.cleanup()

    def _configure_resolved_route(self):
        self.acc_root = _make_candidate(
            self.temp_dir,
            "ACC",
            ["01 Projects", "03 Automations", "04 Resources"],
            ["TestUser"],
            include_route_key=True,
        )
        self.destination_user = os.path.join(
            self.acc_root,
            "Project Files",
            "03 Automations",
            "Usage",
            "TestUser",
        )
        self.route.resolve_usage_route = lambda **kwargs: {
            "status": "resolved",
            "reason": "test_route",
            "resolved_root": self.acc_root,
            "candidate_count": 1,
        }
        return self.destination_user

    def _set_cleanup_version(self, version):
        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        state_payload["local_telemetry_cleanup_version"] = version
        self.route.save_state(
            state_payload,
            source_folder=self.source_folder,
        )

    def _run_shutdown(
        self,
        is_last_process=True,
        process_status=None,
    ):
        if process_status is None:
            if is_last_process:
                process_status = "last_revit_process"
            else:
                process_status = "another_revit_process"

        def process_check(process_type=None, diagnostics=None):
            if diagnostics is not None:
                diagnostics["status"] = process_status
                diagnostics["error"] = ""
                if process_status == "another_revit_process":
                    diagnostics["other_process_id"] = 200
            return is_last_process

        with mock.patch.object(
            self.startup,
            "_is_last_revit_process",
            side_effect=process_check,
        ), mock.patch.object(
            self.startup.getpass,
            "getuser",
            return_value="TestUser",
        ):
            return self.startup._on_app_closing(None, None)

    def test_startup_path_and_telemetry_folder_normalizers(self):
        self.assertEqual(
            os.path.normcase(os.path.normpath("A/../B")),
            self.startup._normalize_path("A/../B"),
        )
        canonical_folder = os.path.abspath(os.path.normpath(self.source_folder))
        self.assertEqual(
            canonical_folder,
            self.startup._canonical_telemetry_folder(self.source_folder),
        )
        self.assertTrue(
            self.startup._telemetry_folder_matches(
                canonical_folder,
                canonical_folder,
            )
        )
        self.assertFalse(
            self.startup._telemetry_folder_matches(
                canonical_folder.replace("\\", "\\\\"),
                canonical_folder,
            )
        )

    def test_release_metadata_reader_parses_toolbar_and_build_versions(self):
        metadata_path = os.path.join(self.temp_dir, "about.yaml")
        _write_text(
            metadata_path,
            "toolbar_version: '9.8.7' # release\n"
            "build: \"20991231+2359\"\n",
        )

        metadata = self.startup._read_pytools_release_metadata(metadata_path)

        self.assertEqual("9.8.7", metadata["toolbar_version"])
        self.assertEqual("20991231+2359", metadata["build_version"])

    def test_release_metadata_reader_fails_silently_when_missing(self):
        metadata = self.startup._read_pytools_release_metadata(
            os.path.join(self.temp_dir, "missing_about.yaml")
        )

        self.assertEqual("", metadata["toolbar_version"])
        self.assertEqual("", metadata["build_version"])

    def test_configure_telemetry_canonicalizes_folder_without_touching_endpoints(self):
        doubled_folder = self.source_folder.replace("\\", "\\\\")
        self.shim.telemetry.utc_timestamps = False
        self.shim.telemetry.active = False
        self.shim.telemetry.telemetry_file_dir = doubled_folder
        self.shim.telemetry.include_hooks = False
        self.shim.telemetry.telemetry_file_path = ""
        self.shim.telemetry.calls = []
        self.shim.telemetry.setup_count = 0

        self.startup._configure_pyrevit_telemetry()

        changed = dict(self.shim.telemetry.calls)
        self.assertEqual(True, changed["utc_timestamps"])
        self.assertEqual(True, changed["active"])
        self.assertEqual(
            os.path.abspath(os.path.normpath(self.source_folder)),
            changed["telemetry_file_dir"],
        )
        self.assertEqual(True, changed["include_hooks"])
        self.assertNotIn("telemetry_server_url", changed)
        self.assertNotIn("active_app", changed)
        self.assertNotIn("apptelemetry_server_url", changed)
        self.assertNotIn("apptelemetry_event_flags", changed)
        self.assertEqual(1, self.shim.telemetry.setup_count)
        self.assertTrue(self.shim.telemetry.telemetry_file_path)

    def test_configure_telemetry_initializes_blank_runtime_file(self):
        self.shim.telemetry.telemetry_file_path = ""
        self.shim.telemetry.calls = []
        self.shim.telemetry.setup_count = 0

        self.startup._configure_pyrevit_telemetry()

        self.assertEqual([], self.shim.telemetry.calls)
        self.assertEqual(1, self.shim.telemetry.setup_count)
        self.assertTrue(self.shim.telemetry.telemetry_file_path)

    def test_configure_telemetry_leaves_initialized_runtime_untouched(self):
        self.shim.telemetry.calls = []
        self.shim.telemetry.setup_count = 0

        self.startup._configure_pyrevit_telemetry()

        self.assertEqual([], self.shim.telemetry.calls)
        self.assertEqual(0, self.shim.telemetry.setup_count)

    def test_check_acc_sync_does_not_open_ui_when_candidates_exist(self):
        self.startup.telemetry_route.resolve_usage_route = lambda **kwargs: {
            "status": "ambiguous",
            "candidate_count": 2,
        }

        self.startup._check_acc_sync()

    def test_last_revit_process_helper_is_session_scoped_and_injectable(self):
        class FakeProcess(object):
            def __init__(self, process_id, session_id, has_exited=False):
                self.Id = process_id
                self.SessionId = session_id
                self.HasExited = has_exited

        class FakeProcessApi(object):
            def __init__(self, current_process, processes):
                self.current_process = current_process
                self.processes = processes

            def GetCurrentProcess(self):
                return self.current_process

            def GetProcessesByName(self, process_name):
                if process_name != "Revit":
                    raise AssertionError("Unexpected process name")
                return list(self.processes)

        current = FakeProcess(100, 10)
        api = FakeProcessApi(
            current,
            [
                current,
                FakeProcess(200, 11, has_exited=False),
                FakeProcess(300, 10, has_exited=True),
            ],
        )
        diagnostics = {}
        self.assertTrue(
            self.startup._is_last_revit_process(
                process_type=api,
                diagnostics=diagnostics,
            )
        )
        self.assertEqual("last_revit_process", diagnostics["status"])

        api.processes.append(FakeProcess(400, 10, has_exited=False))
        diagnostics = {}
        self.assertFalse(
            self.startup._is_last_revit_process(
                process_type=api,
                diagnostics=diagnostics,
            )
        )
        self.assertEqual("another_revit_process", diagnostics["status"])
        self.assertEqual(400, diagnostics["other_process_id"])

    def test_last_revit_process_helper_fails_safe_on_enumeration_error(self):
        class FakeCurrentProcess(object):
            Id = 100
            SessionId = 10

        class FailingProcessApi(object):
            @staticmethod
            def GetCurrentProcess():
                return FakeCurrentProcess()

            @staticmethod
            def GetProcessesByName(process_name):
                raise RuntimeError("process enumeration failed")

        diagnostics = {}
        self.assertFalse(
            self.startup._is_last_revit_process(
                process_type=FailingProcessApi,
                diagnostics=diagnostics,
            )
        )
        self.assertEqual("inspection_failed", diagnostics["status"])
        self.assertIn("process enumeration failed", diagnostics["error"])

    def test_shutdown_overwrites_exact_names_and_retains_local_state(self):
        destination_user = self._configure_resolved_route()
        self._set_cleanup_version(
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION
        )

        telemetry_name = "original_telemetry.json"
        source_telemetry = os.path.join(
            self.source_folder,
            telemetry_name,
        )
        destination_telemetry = os.path.join(
            destination_user,
            telemetry_name,
        )
        unrelated_json = os.path.join(
            self.source_folder,
            "unrelated.json",
        )
        source_state = os.path.join(
            self.source_folder,
            self.route.STATE_FILE_NAME,
        )
        destination_state = os.path.join(
            destination_user,
            self.route.STATE_FILE_NAME,
        )
        _write_text(source_telemetry, '{"event": "new"}')
        _write_text(destination_telemetry, '{"event": "old"}')
        _write_text(destination_state, '{"snapshot": "old"}')
        _write_text(unrelated_json, '{"leave": true}')
        self.shim.telemetry.telemetry_file_path = source_telemetry

        with mock.patch.object(
            self.route,
            "recover_stale_usage_jsons",
        ) as recovery_mock:
            result = self._run_shutdown(is_last_process=True)

        recovery_mock.assert_not_called()
        self.assertEqual("success", result["status"])
        self.assertFalse(os.path.exists(source_telemetry))
        self.assertEqual(
            '{"event": "new"}',
            _read_text(destination_telemetry),
        )
        self.assertEqual(
            set([telemetry_name, self.route.STATE_FILE_NAME]),
            set(os.listdir(destination_user)),
        )
        self.assertEqual(
            os.path.normcase(os.path.normpath(destination_telemetry)),
            os.path.normcase(os.path.normpath(
                result["current_file"]["destination_path"]
            )),
        )
        self.assertTrue(result["current_file"]["copied"])
        self.assertTrue(result["current_file"]["deleted"])
        self.assertTrue(os.path.isfile(source_state))
        self.assertTrue(os.path.isfile(destination_state))
        self.assertTrue(os.path.isfile(unrelated_json))
        local_state = _read_json(source_state)
        destination_state_payload = _read_json(destination_state)
        self.assertEqual(local_state, destination_state_payload)
        self.assertTrue(self.startup.PYTOOLS_TOOLBAR_VERSION)
        self.assertTrue(self.startup.PYTOOLS_BUILD_VERSION)
        self.assertEqual(
            self.startup.PYTOOLS_TOOLBAR_VERSION,
            local_state["pytools_toolbar_version"],
        )
        self.assertEqual(
            self.startup.PYTOOLS_BUILD_VERSION,
            local_state["pytools_build_version"],
        )
        self.assertEqual(1, local_state["last_transfer"]["files_found"])
        self.assertEqual(1, local_state["last_transfer"]["files_copied"])
        self.assertEqual(1, local_state["last_transfer"]["files_deleted"])
        snapshot_result = local_state["last_state_snapshot_copy"]
        self.assertEqual("success", snapshot_result["status"])
        self.assertTrue(snapshot_result["copy_attempted"])
        self.assertTrue(snapshot_result["copied"])
        self.assertTrue(snapshot_result["overwritten"])
        self.assertTrue(snapshot_result["source_retained"])
        self.assertFalse(snapshot_result["delete_attempted"])
        self.assertFalse(snapshot_result["deleted"])
        self.assertEqual(
            os.path.normcase(os.path.normpath(destination_state)),
            os.path.normcase(os.path.normpath(
                snapshot_result["destination_path"]
            )),
        )
        self.assertFalse(
            os.path.exists(os.path.join(destination_user, "unrelated.json"))
        )
        self.assertEqual([], self.shim.forms.alerts)

    def test_other_revit_process_handles_only_current_and_defers_migration(self):
        destination_user = self._configure_resolved_route()
        current_name = "current_telemetry.json"
        legacy_name = "legacy_telemetry.json"
        current_path = os.path.join(self.source_folder, current_name)
        legacy_path = os.path.join(self.source_folder, legacy_name)
        _write_text(current_path, '{"current": true}')
        _write_text(legacy_path, '{"legacy": true}')
        self.shim.telemetry.telemetry_file_path = current_path

        with mock.patch.object(
            self.startup,
            "_local_telemetry_files",
            side_effect=AssertionError("Folder enumeration is forbidden"),
        ) as enumeration_mock:
            result = self._run_shutdown(is_last_process=False)

        enumeration_mock.assert_not_called()
        self.assertFalse(os.path.exists(current_path))
        self.assertTrue(os.path.isfile(legacy_path))
        self.assertTrue(
            os.path.isfile(os.path.join(destination_user, current_name))
        )
        self.assertFalse(
            os.path.exists(os.path.join(destination_user, legacy_name))
        )
        self.assertTrue(result["current_file"]["copied"])
        self.assertTrue(result["cleanup_deferred"])
        self.assertTrue(result["another_revit_process_open"])
        self.assertTrue(result["legacy_cleanup"]["deferred"])
        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        self.assertLess(
            state_payload.get("local_telemetry_cleanup_version", 0),
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION,
        )
        self.assertEqual(
            self.startup.PYTOOLS_BUILD_VERSION,
            state_payload["pytools_build_version"],
        )
        snapshot_result = state_payload["last_state_snapshot_copy"]
        self.assertEqual("success", snapshot_result["status"])
        self.assertTrue(snapshot_result["copy_attempted"])
        self.assertTrue(snapshot_result["copied"])
        self.assertTrue(snapshot_result["source_retained"])
        self.assertTrue(
            os.path.isfile(
                os.path.join(destination_user, self.route.STATE_FILE_NAME)
            )
        )
        self.assertEqual(
            "deferred",
            state_payload["last_local_telemetry_cleanup"]["status"],
        )

    def test_last_process_first_migration_deletes_legacy_without_copying(self):
        destination_user = self._configure_resolved_route()
        current_name = "current_telemetry.json"
        legacy_names = [
            "legacy_one_telemetry.json",
            "legacy_two_TELEMETRY.JSON",
        ]
        current_path = os.path.join(self.source_folder, current_name)
        unrelated_path = os.path.join(self.source_folder, "unrelated.json")
        _write_text(current_path, '{"current": true}')
        for legacy_name in legacy_names:
            _write_text(
                os.path.join(self.source_folder, legacy_name),
                '{"legacy": true}',
            )
        _write_text(unrelated_path, '{"leave": true}')
        self.shim.telemetry.telemetry_file_path = current_path

        result = self._run_shutdown(is_last_process=True)

        self.assertEqual("success", result["status"])
        self.assertTrue(result["current_file"]["copied"])
        self.assertTrue(result["legacy_cleanup"]["ran"])
        self.assertTrue(result["legacy_cleanup"]["complete"])
        self.assertEqual(2, result["legacy_cleanup"]["files_found"])
        self.assertEqual(2, result["legacy_cleanup"]["files_deleted"])
        self.assertEqual(0, result["legacy_cleanup"]["files_failed"])
        self.assertEqual(
            set([current_name, self.route.STATE_FILE_NAME]),
            set(os.listdir(destination_user)),
        )
        for legacy_name in legacy_names:
            self.assertFalse(
                os.path.exists(os.path.join(self.source_folder, legacy_name))
            )
            self.assertFalse(
                os.path.exists(os.path.join(destination_user, legacy_name))
            )
        self.assertTrue(os.path.isfile(unrelated_path))

        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        self.assertEqual(
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION,
            state_payload["local_telemetry_cleanup_version"],
        )
        self.assertTrue(state_payload["local_telemetry_cleanup_utc"])
        cleanup_state = state_payload["last_local_telemetry_cleanup"]
        self.assertEqual("success", cleanup_state["status"])
        self.assertEqual(2, cleanup_state["files_found"])
        self.assertEqual(2, cleanup_state["files_deleted"])
        self.assertEqual(0, cleanup_state["files_failed"])

    def test_failed_legacy_deletion_leaves_cleanup_version_unset(self):
        destination_user = self._configure_resolved_route()
        current_name = "current_telemetry.json"
        legacy_name = "undeletable_telemetry.json"
        current_path = os.path.join(self.source_folder, current_name)
        legacy_path = os.path.join(self.source_folder, legacy_name)
        _write_text(current_path, '{"current": true}')
        _write_text(legacy_path, '{"legacy": true}')
        self.shim.telemetry.telemetry_file_path = current_path

        original_remove = os.remove

        def selective_remove(path):
            if os.path.normcase(path) == os.path.normcase(legacy_path):
                raise OSError("legacy file is locked")
            return original_remove(path)

        with mock.patch.object(
            self.startup.os,
            "remove",
            side_effect=selective_remove,
        ):
            result = self._run_shutdown(is_last_process=True)

        self.assertTrue(result["current_file"]["copied"])
        self.assertFalse(os.path.exists(current_path))
        self.assertTrue(
            os.path.isfile(os.path.join(destination_user, current_name))
        )
        self.assertTrue(os.path.isfile(legacy_path))
        self.assertEqual("partial_failure", result["legacy_cleanup"]["status"])
        self.assertEqual(1, result["legacy_cleanup"]["files_failed"])
        self.assertFalse(result["legacy_cleanup"]["complete"])

        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        self.assertLess(
            state_payload.get("local_telemetry_cleanup_version", 0),
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION,
        )
        cleanup_state = state_payload["last_local_telemetry_cleanup"]
        self.assertEqual("partial_failure", cleanup_state["status"])
        self.assertEqual(1, cleanup_state["files_failed"])

    def test_last_process_post_migration_sweep_overwrites_exact_names(self):
        destination_user = self._configure_resolved_route()
        self._set_cleanup_version(
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION
        )
        current_name = "current_telemetry.json"
        remaining_names = [
            "reload_telemetry.json",
            "crashed_TELEMETRY.JSON",
        ]
        current_path = os.path.join(self.source_folder, current_name)
        _write_text(current_path, "new current")
        _write_text(
            os.path.join(destination_user, current_name),
            "old current",
        )
        for file_name in remaining_names:
            _write_text(
                os.path.join(self.source_folder, file_name),
                "new {}".format(file_name),
            )
            _write_text(
                os.path.join(destination_user, file_name),
                "old {}".format(file_name),
            )
        unrelated_path = os.path.join(self.source_folder, "unrelated.json")
        _write_text(unrelated_path, "leave this")
        self.shim.telemetry.telemetry_file_path = current_path

        result = self._run_shutdown(is_last_process=True)

        self.assertEqual("success", result["status"])
        self.assertEqual(2, result["post_migration_sweep"]["files_found"])
        self.assertEqual(2, result["post_migration_sweep"]["files_copied"])
        self.assertEqual(2, result["post_migration_sweep"]["files_deleted"])
        self.assertEqual(0, result["post_migration_sweep"]["files_dropped"])
        expected_names = set(
            [current_name, self.route.STATE_FILE_NAME] + remaining_names
        )
        self.assertEqual(expected_names, set(os.listdir(destination_user)))
        self.assertEqual("new current", _read_text(
            os.path.join(destination_user, current_name)
        ))
        for file_name in remaining_names:
            self.assertFalse(
                os.path.exists(os.path.join(self.source_folder, file_name))
            )
            self.assertEqual(
                "new {}".format(file_name),
                _read_text(os.path.join(destination_user, file_name)),
            )
        self.assertTrue(os.path.isfile(unrelated_path))
        self.assertTrue(
            os.path.isfile(
                os.path.join(self.source_folder, self.route.STATE_FILE_NAME)
            )
        )
        self.assertTrue(
            os.path.isfile(
                os.path.join(destination_user, self.route.STATE_FILE_NAME)
            )
        )
        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        self.assertEqual(3, state_payload["last_transfer"]["files_found"])
        self.assertEqual(3, state_payload["last_transfer"]["files_copied"])
        self.assertEqual(3, state_payload["last_transfer"]["files_deleted"])

    def test_current_copy_failure_deletes_source_and_records_drop(self):
        destination_user = self._configure_resolved_route()
        self._set_cleanup_version(
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION
        )
        current_name = "copy_failure_telemetry.json"
        current_path = os.path.join(self.source_folder, current_name)
        _write_text(current_path, '{"drop": true}')
        self.shim.telemetry.telemetry_file_path = current_path

        original_copyfile = self.startup.shutil.copyfile

        def fail_current_copy_only(source_path, destination_path):
            if os.path.normcase(source_path) == os.path.normcase(current_path):
                raise IOError("destination unavailable")
            return original_copyfile(source_path, destination_path)

        with mock.patch.object(
            self.startup.shutil,
            "copyfile",
            side_effect=fail_current_copy_only,
        ):
            result = self._run_shutdown(is_last_process=True)

        current_result = result["current_file"]
        self.assertEqual("partial_success", result["status"])
        self.assertTrue(current_result["copy_attempted"])
        self.assertTrue(current_result["copy_failed"])
        self.assertTrue(current_result["deleted"])
        self.assertTrue(current_result["dropped"])
        self.assertFalse(os.path.exists(current_path))
        self.assertFalse(
            os.path.exists(os.path.join(destination_user, current_name))
        )
        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        transfer_state = state_payload["last_transfer"]
        self.assertEqual(1, transfer_state["files_found"])
        self.assertEqual(0, transfer_state["files_copied"])
        self.assertEqual(1, transfer_state["files_failed"])
        self.assertEqual(1, transfer_state["files_deleted"])
        self.assertEqual(1, transfer_state["files_dropped"])
        self.assertEqual(
            "success",
            state_payload["last_state_snapshot_copy"]["status"],
        )
        self.assertTrue(
            os.path.isfile(
                os.path.join(destination_user, self.route.STATE_FILE_NAME)
            )
        )

    def test_state_snapshot_copy_failure_retains_local_without_rename(self):
        destination_user = self._configure_resolved_route()
        self._set_cleanup_version(
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION
        )
        current_name = "current_telemetry.json"
        current_path = os.path.join(self.source_folder, current_name)
        local_state_path = os.path.join(
            self.source_folder,
            self.route.STATE_FILE_NAME,
        )
        destination_state_path = os.path.join(
            destination_user,
            self.route.STATE_FILE_NAME,
        )
        _write_text(current_path, '{"current": true}')
        _write_text(destination_state_path, '{"snapshot": "old"}')
        self.shim.telemetry.telemetry_file_path = current_path

        original_copyfile = self.startup.shutil.copyfile
        state_copy_attempts = []

        def fail_state_copy_only(source_path, destination_path):
            if os.path.basename(source_path) == self.route.STATE_FILE_NAME:
                state_copy_attempts.append((source_path, destination_path))
                raise IOError("state destination unavailable")
            return original_copyfile(source_path, destination_path)

        with mock.patch.object(
            self.startup.shutil,
            "copyfile",
            side_effect=fail_state_copy_only,
        ):
            result = self._run_shutdown(is_last_process=True)

        self.assertEqual(1, len(state_copy_attempts))
        self.assertTrue(os.path.isfile(local_state_path))
        self.assertEqual(
            '{"snapshot": "old"}',
            _read_text(destination_state_path),
        )
        self.assertEqual(
            set([current_name, self.route.STATE_FILE_NAME]),
            set(os.listdir(destination_user)),
        )
        snapshot_result = _read_json(local_state_path)[
            "last_state_snapshot_copy"
        ]
        self.assertEqual("copy_failed", snapshot_result["status"])
        self.assertTrue(snapshot_result["copy_attempted"])
        self.assertFalse(snapshot_result["copied"])
        self.assertTrue(snapshot_result["overwritten"])
        self.assertTrue(snapshot_result["source_retained"])
        self.assertFalse(snapshot_result["delete_attempted"])
        self.assertFalse(snapshot_result["deleted"])
        self.assertIn("state destination unavailable", snapshot_result["error"])
        self.assertEqual("copy_failed", result["state_snapshot_copy"]["status"])
        self.assertEqual([], self.shim.forms.alerts)

    def test_state_snapshot_not_copied_when_local_state_save_fails(self):
        destination_user = self._configure_resolved_route()
        self._set_cleanup_version(
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION
        )
        current_name = "current_telemetry.json"
        current_path = os.path.join(self.source_folder, current_name)
        destination_state_path = os.path.join(
            destination_user,
            self.route.STATE_FILE_NAME,
        )
        _write_text(current_path, '{"current": true}')
        _write_text(destination_state_path, '{"snapshot": "old"}')
        self.shim.telemetry.telemetry_file_path = current_path

        original_copyfile = self.startup.shutil.copyfile
        state_copy_attempts = []

        def track_state_copy(source_path, destination_path):
            if os.path.basename(source_path) == self.route.STATE_FILE_NAME:
                state_copy_attempts.append((source_path, destination_path))
            return original_copyfile(source_path, destination_path)

        with mock.patch.object(
            self.startup.shutil,
            "copyfile",
            side_effect=track_state_copy,
        ), mock.patch.object(
            self.route,
            "save_state",
            side_effect=IOError("local state save failed"),
        ):
            result = self._run_shutdown(is_last_process=True)

        self.assertEqual([], state_copy_attempts)
        self.assertFalse(result["state_saved"])
        self.assertIn("local state save failed", result["state_error"])
        self.assertEqual(
            "state_save_failed",
            result["state_snapshot_copy"]["status"],
        )
        self.assertFalse(result["state_snapshot_copy"]["copy_attempted"])
        self.assertFalse(result["state_snapshot_copy"]["copied"])
        self.assertEqual(
            '{"snapshot": "old"}',
            _read_text(destination_state_path),
        )
        self.assertEqual(
            set([current_name, self.route.STATE_FILE_NAME]),
            set(os.listdir(destination_user)),
        )
        self.assertEqual([], self.shim.forms.alerts)

    def test_process_detection_failure_prevents_folder_wide_processing(self):
        destination_user = self._configure_resolved_route()
        self._set_cleanup_version(
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION
        )
        current_name = "current_telemetry.json"
        remaining_name = "remaining_telemetry.json"
        current_path = os.path.join(self.source_folder, current_name)
        remaining_path = os.path.join(self.source_folder, remaining_name)
        _write_text(current_path, '{"current": true}')
        _write_text(remaining_path, '{"remaining": true}')
        self.shim.telemetry.telemetry_file_path = current_path

        class FakeCurrentProcess(object):
            Id = 100
            SessionId = 10

        class FailingProcessApi(object):
            @staticmethod
            def GetCurrentProcess():
                return FakeCurrentProcess()

            @staticmethod
            def GetProcessesByName(process_name):
                raise RuntimeError("process enumeration failed")

        with mock.patch.object(
            self.startup,
            "Process",
            FailingProcessApi,
        ), mock.patch.object(
            self.startup,
            "_local_telemetry_files",
            side_effect=AssertionError("Folder enumeration is forbidden"),
        ) as enumeration_mock, mock.patch.object(
            self.startup.getpass,
            "getuser",
            return_value="TestUser",
        ):
            result = self.startup._on_app_closing(None, None)

        enumeration_mock.assert_not_called()
        self.assertTrue(result["current_file"]["copied"])
        self.assertFalse(os.path.exists(current_path))
        self.assertTrue(os.path.isfile(remaining_path))
        self.assertTrue(
            os.path.isfile(os.path.join(destination_user, current_name))
        )
        self.assertFalse(
            os.path.exists(os.path.join(destination_user, remaining_name))
        )
        self.assertFalse(result["is_last_revit_process"])
        self.assertEqual("inspection_failed", result["process_check_status"])
        self.assertTrue(result["cleanup_deferred"])
        self.assertTrue(result["post_migration_sweep"]["deferred"])

    def test_current_path_lookup_failure_defers_folder_wide_processing(self):
        self._configure_resolved_route()
        legacy_path = os.path.join(
            self.source_folder,
            "legacy_telemetry.json",
        )
        _write_text(legacy_path, '{"legacy": true}')

        with mock.patch.object(
            self.shim.telemetry,
            "get_telemetry_file_path",
            side_effect=RuntimeError("telemetry path unavailable"),
        ), mock.patch.object(
            self.startup,
            "_local_telemetry_files",
            side_effect=AssertionError("Folder enumeration is forbidden"),
        ) as enumeration_mock:
            result = self._run_shutdown(is_last_process=True)

        enumeration_mock.assert_not_called()
        self.assertTrue(os.path.isfile(legacy_path))
        self.assertFalse(result["current_file_handled"])
        self.assertTrue(result["cleanup_deferred"])
        self.assertEqual(
            "current_file_lookup_failed",
            result["cleanup_defer_reason"],
        )
        self.assertTrue(result["legacy_cleanup"]["deferred"])
        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        self.assertLess(
            state_payload.get("local_telemetry_cleanup_version", 0),
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION,
        )

    def test_unresolved_destination_retains_current_and_defers_cleanup(self):
        self.route.resolve_usage_route = lambda **kwargs: {
            "status": "not_found",
            "reason": "no_viable_candidates",
            "resolved_root": "",
            "candidate_count": 0,
        }
        current_path = os.path.join(
            self.source_folder,
            "current_telemetry.json",
        )
        legacy_path = os.path.join(
            self.source_folder,
            "legacy_telemetry.json",
        )
        _write_text(current_path, '{"current": true}')
        _write_text(legacy_path, '{"legacy": true}')
        self.shim.telemetry.telemetry_file_path = current_path

        with mock.patch.object(
            self.startup,
            "_local_telemetry_files",
            side_effect=AssertionError("Folder enumeration is forbidden"),
        ) as enumeration_mock:
            result = self._run_shutdown(is_last_process=True)

        enumeration_mock.assert_not_called()
        self.assertTrue(os.path.isfile(current_path))
        self.assertTrue(os.path.isfile(legacy_path))
        self.assertTrue(result["current_file"]["found"])
        self.assertTrue(result["current_file"]["copy_failed"])
        self.assertFalse(result["current_file"]["copy_attempted"])
        self.assertFalse(result["current_file"]["delete_attempted"])
        self.assertFalse(result["current_file"]["deleted"])
        self.assertFalse(result["current_file_handled"])
        self.assertTrue(result["cleanup_deferred"])
        self.assertEqual(
            "destination_unavailable",
            result["cleanup_defer_reason"],
        )
        self.assertTrue(result["legacy_cleanup"]["deferred"])
        state_payload = self.route.load_state(
            source_folder=self.source_folder,
        )
        self.assertLess(
            state_payload.get("local_telemetry_cleanup_version", 0),
            self.startup.LOCAL_TELEMETRY_CLEANUP_VERSION,
        )
        self.assertEqual(
            self.startup.PYTOOLS_BUILD_VERSION,
            state_payload["pytools_build_version"],
        )
        snapshot_result = state_payload["last_state_snapshot_copy"]
        self.assertEqual("destination_unavailable", snapshot_result["status"])
        self.assertFalse(snapshot_result["copy_attempted"])
        self.assertFalse(snapshot_result["copied"])
        self.assertTrue(snapshot_result["source_retained"])

    def test_shutdown_does_not_recover_when_route_unresolved(self):
        transfer_calls = []
        recovery_calls = []
        state_name = self.startup.telemetry_route.STATE_FILE_NAME
        route_stub = types.SimpleNamespace(
            STATE_FILE_NAME=state_name,
            telemetry_source_folder=lambda: self.source_folder,
            resolve_usage_route=lambda **kwargs: {
                "status": "not_found",
                "reason": "no_viable_candidates",
                "resolved_root": "",
                "candidate_count": 1,
            },
            recover_stale_usage_jsons=lambda *args, **kwargs: recovery_calls.append(
                (args, kwargs)
            ),
            record_recovery_state=lambda *args, **kwargs: None,
            record_transfer_state=lambda **kwargs: transfer_calls.append(
                dict(kwargs)
            ),
        )
        self.startup.telemetry_route = route_stub

        with mock.patch.object(
            self.startup.getpass,
            "getuser",
            return_value="TestUser",
        ):
            self.startup._on_app_closing(None, None)

        self.assertEqual([], recovery_calls)
        self.assertEqual(1, len(transfer_calls))
        self.assertEqual("route_unresolved", transfer_calls[0]["status"])

    def test_shutdown_does_not_recover_when_usage_base_missing(self):
        acc_root = _make_dir(
            os.path.join(
                self.temp_dir,
                "ACC",
                "ACCDocs",
                "CoolSys",
                "CED Content Collection",
            )
        )

        transfer_calls = []
        recovery_calls = []
        state_name = self.startup.telemetry_route.STATE_FILE_NAME
        route_stub = types.SimpleNamespace(
            STATE_FILE_NAME=state_name,
            telemetry_source_folder=lambda: self.source_folder,
            resolve_usage_route=lambda **kwargs: {
                "status": "resolved",
                "reason": "test_route",
                "resolved_root": acc_root,
                "candidate_count": 1,
            },
            recover_stale_usage_jsons=lambda *args, **kwargs: recovery_calls.append(
                (args, kwargs)
            ),
            record_recovery_state=lambda *args, **kwargs: None,
            record_transfer_state=lambda **kwargs: transfer_calls.append(
                dict(kwargs)
            ),
        )
        self.startup.telemetry_route = route_stub

        with mock.patch.object(
            self.startup.getpass,
            "getuser",
            return_value="TestUser",
        ):
            self.startup._on_app_closing(None, None)

        self.assertEqual([], recovery_calls)
        self.assertEqual(1, len(transfer_calls))
        self.assertEqual("usage_base_missing", transfer_calls[0]["status"])


class DiagnosticsButtonTests(unittest.TestCase):
    def test_shift_click_report_runs_without_revit(self):
        with tempfile.TemporaryDirectory(
            prefix="ced_adc_diagnostics_"
        ) as temp_dir:
            shim = _PyRevitShim(temp_dir)
            candidate = {
                "score": 123,
                "root": os.path.join(temp_dir, "ACC"),
                "root_exists": True,
                "usage_base_exists": True,
                "usage_subfolder_count": 4,
                "project_files_subfolder_count": 6,
                "project_files_only_automations": False,
                "usage_only_current_user": False,
                "key_info": {"matches_expected_route": True},
            }
            shim.route.get_username = lambda: "TestUser"
            shim.route.load_state = lambda: {
                "approved_root": candidate["root"],
                "last_resolution": {"status": "resolved"},
            }
            shim.route.resolve_usage_route = lambda **kwargs: {
                "status": "resolved",
                "reason": "single_viable_candidate",
                "resolved_root": candidate["root"],
                "best_score": 123,
                "margin": 123,
                "candidate_count": 1,
                "viable_count": 1,
                "scored_candidates": [candidate],
                "state_file": os.path.join(temp_dir, "state.json"),
            }

            with mock.patch.dict(sys.modules, shim.modules(), clear=False):
                runpy.run_path(
                    DIAGNOSTICS_PATH,
                    init_globals={"__shiftclick__": True},
                )

            report = "\n".join(shim.output.lines)
            self.assertIn("# CED Telemetry Route Diagnostics", report)
            self.assertIn("single_viable_candidate", report)
            self.assertIn("| 123 |", report)
            table_blocks = [
                line for line in shim.output.lines
                if "| Score | Root | Exists |" in line
            ]
            self.assertEqual(1, len(table_blocks))
            self.assertIn("| ---: | --- | :---: |", table_blocks[0])
            self.assertIn("| 123 |", table_blocks[0])
            self.assertIn("YES", report)
            self.assertEqual([], shim.forms.alerts)

    def test_manual_resolver_allows_single_unresolved_candidate(self):
        with tempfile.TemporaryDirectory(
            prefix="ced_adc_resolver_"
        ) as temp_dir:
            shim = _PyRevitShim(temp_dir)
            selected_root = os.path.join(
                temp_dir,
                "ACC",
                "ACCDocs",
                "CoolSys",
                "CED Content Collection",
            )
            alert_messages = []
            pick_calls = []
            manual_calls = []
            resolve_calls = {"count": 0}

            def alert(message, **kwargs):
                alert_messages.append(str(message))
                if kwargs.get("yes") and kwargs.get("no"):
                    return True
                return False

            def pick_folder(**kwargs):
                pick_calls.append(dict(kwargs))
                return selected_root

            def resolve_usage_route(**kwargs):
                resolve_calls["count"] += 1
                if resolve_calls["count"] == 1:
                    return {
                        "status": "not_found",
                        "reason": "no_viable_candidates",
                        "resolved_root": "",
                        "candidate_count": 1,
                    }
                return {
                    "status": "resolved",
                    "reason": "manual_approval_saved",
                    "resolved_root": selected_root,
                    "candidate_count": 1,
                    "state_file": os.path.join(temp_dir, "state.json"),
                }

            def set_manual_approved_root(root, username=None):
                manual_calls.append((root, username))
                return {"success": True, "inspected": {}}

            shim.forms.alert = alert
            shim.forms.pick_folder = pick_folder
            shim.route.get_username = lambda: "TestUser"
            shim.route.resolve_usage_route = resolve_usage_route
            shim.route.set_manual_approved_root = set_manual_approved_root
            shim.route.ensure_user_folder = lambda root, username=None: {
                "ok": True,
                "created": True,
                "path": os.path.join(root, "Project Files", "03 Automations", "Usage", username),
            }
            shim.route.recover_stale_usage_jsons = lambda *args, **kwargs: {
                "status": "partial_success",
                "files_found": 6,
                "files_moved": 3,
                "files_failed": 1,
                "files_skipped_existing": 2,
            }
            shim.route.record_recovery_state = lambda *args, **kwargs: None

            with mock.patch.dict(sys.modules, shim.modules(), clear=False):
                runpy.run_path(
                    DIAGNOSTICS_PATH,
                    init_globals={"__shiftclick__": False},
                )

            self.assertEqual(1, len(pick_calls))
            self.assertEqual([(selected_root, "TestUser")], manual_calls)
            completion_message = alert_messages[-1]
            self.assertIn(
                "Stale telemetry files moved: 3",
                completion_message,
            )
            self.assertIn(
                "Stale telemetry files failed: 1",
                completion_message,
            )
            self.assertIn(
                "Stale telemetry files skipped-existing: 2",
                completion_message,
            )
            self.assertFalse(
                any(
                    "Manual resolver is only enabled" in message
                    for message in alert_messages
                )
            )

    def test_manual_resolver_allows_resolved_path_override(self):
        with tempfile.TemporaryDirectory(
            prefix="ced_adc_resolver_override_"
        ) as temp_dir:
            shim = _PyRevitShim(temp_dir)
            old_root = os.path.join(
                temp_dir,
                "Old",
                "ACCDocs",
                "CoolSys",
                "CED Content Collection",
            )
            selected_root = os.path.join(
                temp_dir,
                "New",
                "ACCDocs",
                "CoolSys",
                "CED Content Collection",
            )
            selected_project_files = os.path.join(selected_root, "Project Files")
            alert_messages = []
            pick_calls = []
            manual_calls = []
            resolve_calls = {"count": 0}

            def alert(message, **kwargs):
                alert_messages.append(str(message))
                if kwargs.get("yes") and kwargs.get("no"):
                    return True
                return False

            def pick_folder(**kwargs):
                pick_calls.append(dict(kwargs))
                return selected_project_files

            def resolve_usage_route(**kwargs):
                resolve_calls["count"] += 1
                if resolve_calls["count"] == 1:
                    return {
                        "status": "resolved",
                        "reason": "single_viable_candidate",
                        "resolved_root": old_root,
                        "candidate_count": 1,
                    }
                return {
                    "status": "resolved",
                    "reason": "manual_approval_saved",
                    "resolved_root": selected_root,
                    "candidate_count": 1,
                    "state_file": os.path.join(temp_dir, "state.json"),
                }

            def set_manual_approved_root(root, username=None):
                manual_calls.append((root, username))
                return {"success": True, "inspected": {}}

            shim.forms.alert = alert
            shim.forms.pick_folder = pick_folder
            shim.route.get_username = lambda: "TestUser"
            shim.route.resolve_usage_route = resolve_usage_route
            shim.route.set_manual_approved_root = set_manual_approved_root
            shim.route.ensure_user_folder = lambda root, username=None: {
                "ok": True,
                "created": True,
                "path": os.path.join(root, "Project Files", "03 Automations", "Usage", username),
            }
            shim.route.recover_stale_usage_jsons = lambda *args, **kwargs: {
                "status": "no_stale_jsons",
                "files_found": 0,
                "files_moved": 0,
                "files_failed": 0,
                "files_skipped_existing": 0,
            }
            shim.route.record_recovery_state = lambda *args, **kwargs: None

            with mock.patch.dict(sys.modules, shim.modules(), clear=False):
                runpy.run_path(
                    DIAGNOSTICS_PATH,
                    init_globals={"__shiftclick__": False},
                )

            self.assertEqual(1, len(pick_calls))
            self.assertEqual([(selected_root, "TestUser")], manual_calls)
            self.assertTrue(
                any(old_root in message for message in alert_messages)
            )
            self.assertTrue(
                any("If this path is wrong" in message for message in alert_messages)
            )


if __name__ == "__main__":
    if "--live" in sys.argv:
        print_live_diagnostics()
    else:
        unittest.main(verbosity=2)
