#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Export Space Config"""

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
import storage

TITLE = "Export Space Config (MEPRFP 2.0)"


def main():
    output = script.get_output()
    output.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    save_path = forms.save_file(
        file_ext="yaml",
        title=TITLE,
        default_name="space_config.yaml",
    )
    if not save_path:
        return

    try:
        result = active_yaml.export_space_config_file(doc, save_path)
    except storage.StorageError as exc:
        forms.alert(str(exc), title=TITLE)
        return
    except (IOError, OSError) as exc:
        forms.alert("Failed to write file:\n\n{}".format(exc), title=TITLE)
        return
    except Exception as exc:
        forms.alert(
            "Unexpected error during export:\n\n{}".format(exc),
            title=TITLE,
            exitscript=False,
        )
        raise

    output.print_md(
        "**Export succeeded**\n\n"
        "- Wrote: `{}`\n"
        "- Bytes: `{}`\n"
        "- Schema version: `{}`\n"
        "- space_buckets exported: `{}`\n"
        "- space_profiles exported: `{}`\n".format(
            result["save_path"],
            result["byte_count"],
            result.get("schema_version"),
            result.get("buckets_exported"),
            result.get("profiles_exported"),
        )
    )


if __name__ == "__main__":
    main()
