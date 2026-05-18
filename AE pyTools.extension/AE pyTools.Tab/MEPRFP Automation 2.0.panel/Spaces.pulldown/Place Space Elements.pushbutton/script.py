#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Place Space Elements"""

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
import place_space_elements_window
import space_door_picker

TITLE = "Place Space Elements (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    profile_data = active_yaml.load_active_data(doc) or {}
    if not profile_data.get("space_profiles"):
        forms.alert(
            "No space_profiles in the active YAML. "
            "Use Manage Space Profiles to define some first.",
            title=TITLE,
        )
        return

    # Selection.PickObject can't run while a modal WPF window is up,
    # so we ask the user to click reference doors NOW (before opening
    # the placement preview) for any space that has more than one
    # door AND at least one door-dependent LED. The chosen anchors
    # ride into the placement run via ``door_choices``. The output
    # panel gets a per-space breakdown so when the prompt doesn't
    # appear (zero detected doors, etc.) the reason is visible.
    try:
        door_choices = space_door_picker.pre_pick_doors(
            uidoc, doc, profile_data, output=output,
        )
    except TypeError as exc:
        # Stale cached module without the new ``output`` kwarg —
        # CPython engine hasn't picked up the latest source. Fall back
        # so the placement still works, but warn so the user knows
        # they're missing the diagnostic output until Revit restarts.
        if "output" in str(exc):
            output.print_md(
                "**Note:** pyRevit's CPython engine has a stale "
                "cached version of `space_door_picker` (no `output` "
                "kwarg). Falling back to no-diagnostics mode for "
                "this run. **Restart Revit** to pick up the latest "
                "version and see the per-space pick analysis."
            )
            door_choices = space_door_picker.pre_pick_doors(
                uidoc, doc, profile_data,
            )
        else:
            raise

    controller = place_space_elements_window.show_modal(
        doc=doc,
        profile_data=profile_data,
        door_choices=door_choices,
    )

    # Surface the placement result + any warnings (parameter writes
    # that didn't apply, exceptions during placement, etc.) so the
    # user can see WHY a captured "Elevation from Level" / "Mark"
    # / etc. didn't land on the placed instance.
    result = getattr(controller, "last_result", None)
    if controller is not None and getattr(controller, "committed", False) and result is not None:
        output.print_md(
            "**Place Space Elements complete**\n\n"
            "- Placed: `{}`\n"
            "- Failed: `{}`\n"
            "- Warnings: `{}`\n".format(
                result.n_placed, result.n_failed, len(result.warnings),
            )
        )
        if result.warnings:
            output.print_md(
                "\n**Warnings:**\n\n"
                + "\n".join("- {}".format(w) for w in result.warnings)
            )


if __name__ == "__main__":
    main()
