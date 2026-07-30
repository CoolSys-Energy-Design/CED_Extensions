#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Combine Yaml

Merge two equipment-definition YAML files. File A keeps its numbering;
File B's profiles are appended with every EQ / SET / LED / annotation id
(and the references that point at them) renumbered so nothing collides
with A. The result is written to a new file (default name "combined").
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

from pyrevit import script

import forms_compat as forms
import yaml_io
import combine_yaml

TITLE = "Combine Yaml (MEPRFP 2.0)"


def _read(path):
    with io.open(path, "r", encoding="utf-8") as f:
        return f.read()


def _load_defs(path, label):
    """Parse a file and return its dict, or (None, error_message)."""
    try:
        data = yaml_io.parse(_read(path))
    except Exception as ex:
        return None, "Failed to read/parse {} ({}):\n{}".format(label, path, ex)
    if not isinstance(data, dict):
        return None, "{} is not a YAML mapping: {}".format(label, path)
    if not data.get("equipment_definitions"):
        return None, "{} has no equipment_definitions: {}".format(label, path)
    return data, None


def main():
    output = script.get_output()
    output.close_others()

    path_a = forms.pick_file(
        file_ext="yaml", title="Combine Yaml - pick File A (keeps its numbering)"
    )
    if not path_a:
        return
    path_b = forms.pick_file(
        file_ext="yaml", title="Combine Yaml - pick File B (appended & renumbered)"
    )
    if not path_b:
        return

    if os.path.normcase(os.path.abspath(path_a)) == os.path.normcase(
        os.path.abspath(path_b)
    ):
        if not forms.confirm(
            "File A and File B are the same file. Combine it with a copy of "
            "itself anyway?",
            title=TITLE,
        ):
            return

    data_a, err = _load_defs(path_a, "File A")
    if err:
        forms.alert(err, title=TITLE)
        return
    data_b, err = _load_defs(path_b, "File B")
    if err:
        forms.alert(err, title=TITLE)
        return

    try:
        combined, summary = combine_yaml.combine(data_a, data_b)
        text = yaml_io.dump(combined)
    except Exception as ex:
        forms.alert("Combine failed:\n{}".format(ex), title=TITLE)
        return

    save_path = forms.save_file(
        file_ext="yaml",
        title="Save combined YAML",
        default_name="combined",
    )
    if not save_path:
        return

    with io.open(save_path, "w", encoding="utf-8") as f:
        f.write(text)

    output.print_md(
        "**Combine Yaml complete**\n\n"
        "- File A: `{}` ({} profiles, unchanged)\n"
        "- File B: `{}` ({} profiles, renumbered)\n"
        "- Combined: **{} profiles** -> `{}`\n"
        "- Renumbered {} EQ id(s) and {} SET id(s) from File B".format(
            os.path.basename(path_a),
            summary["a_profile_count"],
            os.path.basename(path_b),
            summary["b_profile_count"],
            summary["combined_profile_count"],
            save_path,
            len(summary["eq_map"]),
            len(summary["set_map"]),
        )
    )


if __name__ == "__main__":
    main()
