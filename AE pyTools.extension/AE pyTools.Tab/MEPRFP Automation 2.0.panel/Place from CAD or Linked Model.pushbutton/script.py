#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Place from CAD or Linked Model"""

import os
import sys

_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "lib")
)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import _dev_reload
_dev_reload.purge()

from pyrevit import revit, script

import forms_compat as forms
import active_yaml
import placement_window
import shared_params

TITLE = "Place from CAD or Linked Model (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    if not shared_params.prompt_and_bind(
        doc, forms, TITLE,
        reason="required for placement to write back to placed elements",
    ):
        return

    profile_data = active_yaml.load_active_data(doc)
    if not profile_data.get("equipment_definitions"):
        forms.alert(
            "No profiles in the active store. "
            "Use 'New Profile' or 'Import YAML File' first.",
            title=TITLE,
        )
        return

    # Modeless: the dialog runs the placement through an ExternalEvent so
    # the active-view switch that binds the Level on workplane-based
    # families is legal (see placement_apply). show_modeless returns
    # immediately; the run report is printed to ``output`` by the
    # controller's on-complete callback after the user clicks Place.
    placement_window.show_modeless(
        doc, profile_data, uidoc=revit.uidoc, output=output,
    )


if __name__ == "__main__":
    main()
