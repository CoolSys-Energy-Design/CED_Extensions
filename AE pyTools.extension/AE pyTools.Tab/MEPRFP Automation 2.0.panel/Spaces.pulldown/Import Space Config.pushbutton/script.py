#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Import Space Config"""

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
import schema as _schema
import yaml_io

TITLE = "Import Space Config (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    picked = forms.pick_file(
        file_ext="yaml",
        title="Select space-config YAML",
    )
    if not picked:
        return

    try:
        result = active_yaml.import_space_config_file(doc, picked)
    except _schema.SchemaVersionError as exc:
        forms.alert(
            "Schema version check failed:\n\n{}".format(exc), title=TITLE,
        )
        return
    except yaml_io.YamlError as exc:
        forms.alert("Failed to parse YAML:\n\n{}".format(exc), title=TITLE)
        return
    except (IOError, OSError) as exc:
        forms.alert("Failed to read file:\n\n{}".format(exc), title=TITLE)
        return
    except ValueError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    except Exception as exc:
        forms.alert(
            "Unexpected error during import:\n\n{}".format(exc),
            title=TITLE,
            exitscript=False,
        )
        raise

    migrated = result["input_schema_version"] != result["stored_schema_version"]
    version_line = "`{}`".format(result["stored_schema_version"])
    if migrated:
        version_line = "`{}` (migrated from `{}`)".format(
            result["stored_schema_version"], result["input_schema_version"],
        )

    blank_note = ""
    if result.get("blank"):
        blank_note = (
            "\n\n*Blank file imported — empty space_buckets and "
            "space_profiles lists were created. Use Manage Space Buckets / "
            "Manage Space Profiles to populate them.*"
        )

    output.print_md(
        "**Import succeeded**\n\n"
        "- Source: `{}`\n"
        "- Schema version: {}\n"
        "- space_buckets imported: `{}` (replaced `{}`)\n"
        "- space_profiles imported: `{}` (replaced `{}`)\n\n"
        "Per-project classifications were left untouched. If a saved "
        "classification points at a bucket-id that doesn't exist in the "
        "imported config, the placement engine will skip it; re-run "
        "Classify Spaces to refresh assignments.{}".format(
            result["source_path"],
            version_line,
            result["buckets_imported"],
            result["buckets_replaced"],
            result["profiles_imported"],
            result["profiles_replaced"],
            blank_note,
        )
    )


if __name__ == "__main__":
    main()
