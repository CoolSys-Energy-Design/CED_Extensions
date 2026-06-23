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
    def __init__(self):
        types.ModuleType.__init__(self, "pyrevit.telemetry")
        self.calls = []
        self.setup_count = 0

    def _record(self, name, value):
        self.calls.append((name, value))

    def set_telemetry_utc_timestamp(self, value):
        self._record("utc_timestamps", value)

    def set_telemetry_state(self, value):
        self._record("active", value)

    def set_telemetry_file_dir(self, value):
        self._record("telemetry_file_dir", value)

    def set_telemetry_server_url(self, value):
        self._record("telemetry_server_url", value)

    def set_telemetry_include_hooks(self, value):
        self._record("include_hooks", value)

    def set_apptelemetry_state(self, value):
        self._record("active_app", value)

    def set_apptelemetry_server_url(self, value):
        self._record("apptelemetry_server_url", value)

    def set_apptelemetry_event_flags(self, value):
        self._record("apptelemetry_event_flags", value)

    def setup_telemetry(self):
        self.setup_count += 1


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
        self.telemetry = _FakeTelemetry()
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

    def tearDown(self):
        self.temp_dir_obj.cleanup()

    def test_startup_path_and_flag_normalizers(self):
        self.assertEqual(
            os.path.normcase(os.path.normpath("A/../B")),
            self.startup._normalize_path("A/../B"),
        )
        self.assertEqual(16, self.startup._event_flags_to_int("0x10"))
        self.assertEqual(12, self.startup._event_flags_to_int("12"))
        self.assertEqual(0, self.startup._event_flags_to_int("invalid"))

    def test_configure_telemetry_uses_public_api_for_mismatches(self):
        self.shim.config.options = {
            "apptelemetry_event_flags": "0x2",
        }
        self.shim.telemetry.calls = []
        self.shim.telemetry.setup_count = 0

        self.startup._configure_pyrevit_telemetry()

        changed = dict(self.shim.telemetry.calls)
        self.assertEqual(True, changed["utc_timestamps"])
        self.assertEqual(True, changed["active"])
        self.assertEqual(self.source_folder, changed["telemetry_file_dir"])
        self.assertEqual(0, changed["apptelemetry_event_flags"])
        self.assertEqual(1, self.shim.telemetry.setup_count)

    def test_check_acc_sync_does_not_open_ui_when_candidates_exist(self):
        self.startup.telemetry_route.resolve_usage_route = lambda **kwargs: {
            "status": "ambiguous",
            "candidate_count": 2,
        }

        self.startup._check_acc_sync()

    def test_shutdown_copies_files_and_retains_local_telemetry(self):
        acc_root = _make_candidate(
            self.temp_dir,
            "ACC",
            ["01 Projects", "03 Automations", "04 Resources"],
            ["ExistingUser"],
            include_route_key=True,
        )
        usage_base = os.path.join(
            acc_root,
            "Project Files",
            "03 Automations",
            "Usage",
        )
        destination_user = _make_dir(
            os.path.join(usage_base, "TestUser")
        )

        telemetry_name = "session.json"
        state_name = self.startup.telemetry_route.STATE_FILE_NAME
        source_telemetry = os.path.join(self.source_folder, telemetry_name)
        source_state = os.path.join(self.source_folder, state_name)
        _write_text(source_telemetry, '{"event": 1}')
        _write_text(source_state, '{"local": true}')
        _make_dir(os.path.join(self.source_folder, "ignored_folder"))
        _write_text(
            os.path.join(destination_user, telemetry_name),
            '{"existing": true}',
        )

        transfer_calls = []
        route_stub = types.SimpleNamespace(
            STATE_FILE_NAME=state_name,
            telemetry_source_folder=lambda: self.source_folder,
            resolve_usage_route=lambda **kwargs: {
                "status": "resolved",
                "reason": "test_route",
                "resolved_root": acc_root,
            },
            record_transfer_state=lambda **kwargs: transfer_calls.append(
                dict(kwargs)
            ),
        )
        self.startup.telemetry_route = route_stub

        with mock.patch.object(
            self.startup.getpass,
            "getuser",
            return_value="TestUser",
        ), mock.patch.object(
            self.startup.time,
            "time",
            return_value=1234567890,
        ):
            self.startup._on_app_closing(None, None)

        copied_path = os.path.join(
            destination_user,
            "session_1234567890.json",
        )
        self.assertTrue(os.path.isfile(source_telemetry))
        self.assertTrue(os.path.isfile(source_state))
        self.assertTrue(os.path.isfile(copied_path))
        self.assertFalse(os.path.isfile(os.path.join(destination_user, state_name)))
        self.assertEqual(1, len(transfer_calls))
        self.assertEqual("success", transfer_calls[0]["status"])
        self.assertEqual(1, transfer_calls[0]["files_copied"])
        self.assertEqual(0, transfer_calls[0]["files_failed"])


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
            self.assertIn("YES", report)
            self.assertEqual([], shim.forms.alerts)


if __name__ == "__main__":
    if "--live" in sys.argv:
        print_live_diagnostics()
    else:
        unittest.main(verbosity=2)
