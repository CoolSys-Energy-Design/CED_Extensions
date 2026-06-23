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
    output.print_md(
        "| Score | Root | Exists | Usage Base | Usage Subfolders | PF Subfolders | Only Automations | Only Current User | Key Match |"
    )
    output.print_md("| ---: | --- | :---: | :---: | ---: | ---: | :---: | :---: | :---: |")

    for item in scored:
        output.print_md(
            "| {} | `{}` | {} | {} | {} | {} | {} | {} | {} |".format(
                item.get("score", 0),
                item.get("root", ""),
                _bool_text(item.get("root_exists")),
                _bool_text(item.get("usage_base_exists")),
                item.get("usage_subfolder_count", 0),
                item.get("project_files_subfolder_count", 0),
                _bool_text(item.get("project_files_only_automations")),
                _bool_text(item.get("usage_only_current_user")),
                _bool_text((item.get("key_info") or {}).get("matches_expected_route")),
            )
        )


def _manual_resolve():
    username = telemetry_route.get_username()
    result = telemetry_route.resolve_usage_route(username=username, persist=True)
    candidate_count = int(result.get("candidate_count", 0) or 0)
    resolved_root = result.get("resolved_root", "")

    if candidate_count == 1:
        forms.alert(
            "Manual resolver is only enabled when candidate count is 0 or more than 1.\n\n"
            "Current resolved root:\n{}".format(resolved_root or "(none)"),
            title="ACC Path Resolver",
            ok=True,
        )
        return

    prompt = (
        "Detected candidate count: {}\n"
        "Current status: {}\n\n"
        "Pick the CED Content Collection root folder now?"
    ).format(candidate_count, result.get("status", ""))
    if not forms.alert(prompt, title="ACC Path Resolver", yes=True, no=True):
        return

    selected_root = forms.pick_folder(title="Select CED Content Collection root folder")
    if not selected_root:
        return

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
    cleanup_result = telemetry_route.cleanup_stale_user_folders(selected_root, username=username)
    refreshed = telemetry_route.resolve_usage_route(username=username, persist=True)

    forms.alert(
        "Manual path saved.\n\n"
        "Approved root:\n{}\n\n"
        "Resolved status: {}\n"
        "Resolved root:\n{}\n\n"
        "User folder created: {}\n"
        "Cleanup removed empty stale folders: {}\n"
        "Cleanup skipped folders: {}\n\n"
        "State file:\n{}".format(
            selected_root,
            refreshed.get("status", ""),
            refreshed.get("resolved_root", ""),
            user_folder_result.get("created", False),
            len(cleanup_result.get("removed", [])),
            len(cleanup_result.get("skipped", [])),
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
