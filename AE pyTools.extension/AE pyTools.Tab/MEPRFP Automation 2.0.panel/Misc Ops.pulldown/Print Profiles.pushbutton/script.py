#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Print Profiles

Write every profile name in the active Extensible-Storage YAML to a
plain text file, one name per row.
"""

import io
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

TITLE = "Print Profiles (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()

    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    data = active_yaml.load_active_data(doc)
    eqdefs = (data or {}).get("equipment_definitions") or []
    if not eqdefs:
        forms.alert("No profiles in the active store.", title=TITLE)
        return

    names = []
    for profile in eqdefs:
        if not isinstance(profile, dict):
            continue
        name = profile.get("name")
        names.append(name if name not in (None, "") else "<unnamed>")

    save_path = forms.save_file(
        file_ext="txt",
        title="Save profile names",
        default_name="profile_names",
    )
    if not save_path:
        return

    with io.open(save_path, "w", encoding="utf-8") as f:
        f.write("\n".join(names))
        f.write("\n")

    output.print_md(
        "**Print Profiles complete** — wrote {} profile name(s) to:\n\n`{}`".format(
            len(names), save_path
        )
    )


if __name__ == "__main__":
    main()
