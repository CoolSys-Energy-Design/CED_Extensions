# -*- coding: utf-8 -*-
__title__ = "ACC Path\nResolver"
__doc__ = "Click to resolve telemetry ACC path manually. Shift+Click for diagnostics."

import os
import sys
import traceback

from pyrevit import forms, script

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
EXT_ROOT = os.path.abspath(os.path.join(THIS_DIR, os.pardir, os.pardir, os.pardir))
if EXT_ROOT not in sys.path:
    sys.path.append(EXT_ROOT)

import telemetry_route


output = script.get_output()
logger = script.get_logger()


def _bool_text(flag):
    return "YES" if bool(flag) else "NO"


def _md_cell(value):
    text = str(value)
    return text.replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def _selected_project_root(selected_root):
    selected_root = os.path.normpath(selected_root or "")
    if not selected_root:
        return ""

    if os.path.basename(selected_root).lower() == "project files":
        return os.path.dirname(selected_root)

    return selected_root


def _print_diagnostics():
    username = telemetry_route.get_username()
    result = telemetry_route.resolve_usage_route(username=username, persist=False)
    state_payload = telemetry_route.load_state()

    output.print_md("# CED Telemetry Route Diagnostics")
    output.print_md("- Shift+Click mode is read-only.")
    output.print_md("- No files are created, moved, copied, or deleted in this mode.")
    output.print_md("")
    output.print_md("## Route State")
    output.print_md("- **State file:** `{}`".format(result.get("state_file", "")))
    output.print_md("- **Approved root:** `{}`".format(state_payload.get("approved_root", "")))
    output.print_md("- **Last status:** `{}`".format((state_payload.get("last_resolution") or {}).get("status", "")))
    output.print_md("")
    last_recovery = state_payload.get("last_stale_recovery") or {}
    output.print_md("## Last Stale Telemetry Recovery")
    output.print_md("- **status:** `{}`".format(last_recovery.get("status", "")))
    output.print_md("- **files found:** `{}`".format(last_recovery.get("files_found", 0)))
    output.print_md("- **files moved:** `{}`".format(last_recovery.get("files_moved", 0)))
    output.print_md("- **files failed:** `{}`".format(last_recovery.get("files_failed", 0)))
    output.print_md("- **files skipped-existing:** `{}`".format(last_recovery.get("files_skipped_existing", 0)))
    output.print_md("")
    output.print_md("## Resolution Summary")
    output.print_md("- **status:** `{}`".format(result.get("status", "")))
    output.print_md("- **reason:** `{}`".format(result.get("reason", "")))
    output.print_md("- **resolved_root:** `{}`".format(result.get("resolved_root", "")))
    output.print_md("- **score / margin:** `{}` / `{}`".format(result.get("best_score", 0), result.get("margin", 0)))
    output.print_md("- **candidates / viable:** `{}` / `{}`".format(result.get("candidate_count", 0), result.get("viable_count", 0)))
    output.print_md("")

    scored = result.get("scored_candidates", [])
    if not scored:
        output.print_md("## Candidates")
        output.print_md("- No candidates found.")
        return

    output.print_md("## Candidates")
    table_lines = [
        "| Score | Root | Exists | Usage Base | Usage Subfolders | PF Subfolders | Only Automations | Only Current User | Key Match |"
    ]
    table_lines.append("| ---: | --- | :---: | :---: | ---: | ---: | :---: | :---: | :---: |")

    for item in scored:
        table_lines.append(
            "| {} | `{}` | {} | {} | {} | {} | {} | {} | {} |".format(
                item.get("score", 0),
                _md_cell(item.get("root", "")),
                _bool_text(item.get("root_exists")),
                _bool_text(item.get("usage_base_exists")),
                item.get("usage_subfolder_count", 0),
                item.get("project_files_subfolder_count", 0),
                _bool_text(item.get("project_files_only_automations")),
                _bool_text(item.get("usage_only_current_user")),
                _bool_text((item.get("key_info") or {}).get("matches_expected_route")),
            )
        )
    output.print_md("\n".join(table_lines))


def _manual_resolve():
    username = telemetry_route.get_username()
    result = telemetry_route.resolve_usage_route(username=username, persist=True)
    candidate_count = int(result.get("candidate_count", 0) or 0)
    resolved_root = result.get("resolved_root", "")

    prompt = (
        "Detected candidate count: {}\n"
        "Current status: {}\n"
        "Current resolved root:\n{}\n\n"
        "If this path is wrong, pick the correct ACC CED Content Collection root now.\n\n"
        "Continue?"
    ).format(candidate_count, result.get("status", ""), resolved_root or "(none)")
    if not forms.alert(prompt, title="ACC Path Resolver", yes=True, no=True):
        return

    selected_root = forms.pick_folder(title="Select CED Content Collection root folder")
    if not selected_root:
        return

    selected_root = _selected_project_root(selected_root)
    save_result = telemetry_route.set_manual_approved_root(selected_root, username=username)
    if not save_result.get("success"):
        inspected = save_result.get("inspected", {})
        forms.alert(
            "Selected folder is not a valid CED Content Collection root.\n\n"
            "Root exists: {}\n"
            "Project Files exists: {}\n"
            "Usage base exists: {}\n\n"
            "Selected:\n{}".format(
                inspected.get("root_exists"),
                inspected.get("project_files_exists"),
                inspected.get("usage_base_exists"),
                selected_root,
            ),
            title="ACC Path Resolver",
            ok=True,
        )
        return

    user_folder_result = telemetry_route.ensure_user_folder(selected_root, username=username)
    recovery_result = {
        "status": "not_available",
        "files_found": 0,
        "files_moved": 0,
        "files_failed": 0,
        "files_skipped_existing": 0,
    }
    if hasattr(telemetry_route, "recover_stale_usage_jsons"):
        try:
            recovery_result = telemetry_route.recover_stale_usage_jsons(
                selected_root,
                username=username,
            )
            if hasattr(telemetry_route, "record_recovery_state"):
                telemetry_route.record_recovery_state(recovery_result)
        except Exception as ex:
            recovery_result = {
                "status": "error",
                "files_found": 0,
                "files_moved": 0,
                "files_failed": 0,
                "files_skipped_existing": 0,
                "error": str(ex),
            }
    refreshed = telemetry_route.resolve_usage_route(username=username, persist=True)

    forms.alert(
        "Manual path saved.\n\n"
        "Approved root:\n{}\n\n"
        "Resolved status: {}\n"
        "Resolved root:\n{}\n\n"
        "User folder created: {}\n"
        "Stale recovery status: {}\n"
        "Stale telemetry files found: {}\n"
        "Stale telemetry files moved: {}\n"
        "Stale telemetry files failed: {}\n"
        "Stale telemetry files skipped-existing: {}\n\n"
        "State file:\n{}".format(
            selected_root,
            refreshed.get("status", ""),
            refreshed.get("resolved_root", ""),
            user_folder_result.get("created", False),
            recovery_result.get("status", ""),
            recovery_result.get("files_found", 0),
            recovery_result.get("files_moved", 0),
            recovery_result.get("files_failed", 0),
            recovery_result.get("files_skipped_existing", 0),
            refreshed.get("state_file", ""),
        ),
        title="ACC Path Resolver",
        ok=True,
    )


def main():

    if __shiftclick__:
        _print_diagnostics()
        return
    _manual_resolve()


try:
    main()
except Exception as ex:
    forms.alert("ACC Path Resolver failed:\n{}".format(ex), title="ACC Path Resolver", ok=True)
    output.print_md("## ERROR")
    output.print_md("`{}`".format(ex))
    output.print_md("```text\n{}\n```".format(traceback.format_exc()))
    logger.exception("ACC Path Resolver failed")
