#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Manage Space Profiles"""

import os
import sys

_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "..", "lib")
)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import _dev_reload
_dev_reload.purge()

from pyrevit import revit, script

import forms_compat as forms
import active_yaml
import manage_space_profiles_window

TITLE = "Manage Space Profiles (MEPRFP 2.0)"


def _profile_save_summary(profile_data):
    """Diagnostic — describe what's about to be saved.

    Lists every space_profile with its bucket_id and a per-LED summary
    of placement-rule.kind, so we can see in the output panel whether
    the in-memory dict actually carries the user's bucket and anchor
    selections.
    """
    profiles = profile_data.get("space_profiles") or []
    if not profiles:
        return "(no space_profiles in payload)"
    lines = []
    for p in profiles:
        if not isinstance(p, dict):
            continue
        pid = p.get("id") or "?"
        name = p.get("name") or "?"
        bucket = p.get("bucket_id") or "(none)"
        led_kinds = []
        for s in p.get("linked_sets") or ():
            if not isinstance(s, dict):
                continue
            for led in s.get("linked_element_definitions") or ():
                if not isinstance(led, dict):
                    continue
                rule = led.get("placement_rule") or {}
                led_kinds.append(rule.get("kind") or "?")
        lines.append("- `{}` *{}* bucket=`{}` kinds=`{}`".format(
            pid, name, bucket, led_kinds,
        ))
    return "\n".join(lines)


def _save_dirty_edits(doc, profile_data, output, action):
    summary = _profile_save_summary(profile_data)
    output.print_md(
        "**Saving space-profile edits**\n\n"
        "Snapshot of the in-memory payload about to be persisted:\n\n"
        "{}\n".format(summary)
    )
    try:
        with revit.Transaction(action, doc=doc):
            active_yaml.save_active_data(doc, profile_data, action=action)
    except Exception as exc:
        output.print_md(
            "**Save FAILED**\n\n"
            "- Action: `{}`\n"
            "- Error type: `{}`\n"
            "- Error: {}\n\n"
            "Your edits are still in memory in this Revit session — try Save "
            "again from the editor before closing the project. If the same "
            "error repeats, paste the message above so we can diagnose."
            .format(action, type(exc).__name__, exc)
        )
        raise
    output.print_md("**Space-profile edits saved.**")


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc) or {}

    if not isinstance(profile_data.get("space_profiles"), list):
        profile_data["space_profiles"] = []
    if not isinstance(profile_data.get("space_buckets"), list):
        profile_data["space_buckets"] = []

    if not profile_data["space_buckets"]:
        forms.alert(
            "No space_buckets are defined in the active YAML.\n\n"
            "You can still create profiles, but they need a bucket "
            "reference to apply at placement time. Add at least one "
            "bucket by hand-editing the YAML and re-importing, or open "
            "Manage Space Profiles after importing a starter YAML.",
            title=TITLE,
        )

    controller = manage_space_profiles_window.ManageSpaceProfilesController(
        profile_data=profile_data, doc=doc,
    )
    controller.show()

    if controller.dirty:
        _save_dirty_edits(doc, profile_data, output, action="Manage Space Profiles edit")


if __name__ == "__main__":
    main()
