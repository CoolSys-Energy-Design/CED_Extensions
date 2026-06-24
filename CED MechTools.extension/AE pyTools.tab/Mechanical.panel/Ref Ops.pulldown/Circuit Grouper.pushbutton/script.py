# -*- coding: utf-8 -*-
__title__ = "Circuit Grouper"
__doc__ = ("Group circuitable devices (case controllers) by CKT_Circuit Number_CEDT, "
           "validate voltage/pole compatibility, regroup, then create one native "
           "Revit circuit per group.")

import os
import sys

from pyrevit import revit, forms, script

# make the sibling cg_* modules importable
THIS_DIR = os.path.dirname(__file__)
if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

import cg_core
import cg_collect
import cg_window
import cg_apply

logger = script.get_logger()
output = script.get_output()
TITLE = "Circuit Grouper"


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    rows_data = cg_collect.collect_devices(doc, cg_core.DEFAULT_SPEC_KEY)
    if not rows_data:
        forms.alert(
            "No devices found with BOTH CKT_Circuit Number_CEDT and Identity Mark populated.\n"
            "(Run Place Identity Mark first to stamp the case controllers.)",
            title=TITLE,
        )
        return

    panel_names, name_to_id = cg_collect.collect_panels(doc)

    plans, name_to_id = cg_window.show_window(
        rows_data, panel_names, name_to_id, cg_core.RATING_OPTIONS
    )

    if not plans:
        return  # cancelled or nothing valid

    # Pre-flight: warn about panel/circuit incompatibilities (distribution
    # system, voltage, poles) BEFORE creating anything, and let the user decide.
    panel_issues = cg_apply.preflight_panel_check(doc, plans, name_to_id)
    if panel_issues:
        lines = [
            "The selected panel does not match the circuit for the following "
            "group(s) (distribution system / voltage / poles):",
            "",
        ]
        for issue in panel_issues:
            lines.append(u"  • {}  →  panel '{}'   ({})".format(
                issue["group_key"], issue["panel"], issue["detail"]))
        lines.append("")
        lines.append(
            "If you proceed, these circuits are still created but their panel "
            "is left UNASSIGNED. Proceed anyway?")
        proceed = forms.alert("\n".join(lines), title=TITLE, yes=True, no=True)
        if not proceed:
            forms.alert(
                "Circuiting cancelled - nothing was changed. Fix the panel "
                "selection (or regroup) and run again.",
                title=TITLE,
            )
            return

    report = cg_apply.run(doc, plans, name_to_id, logger)

    # -- report --------------------------------------------------------
    output.print_md("## Circuit Grouper - Results")
    output.print_md("**Circuits created:** {}  |  **Members circuited:** {}".format(
        report["created"], report["members_circuited"]))
    if report["removed_from_existing"]:
        output.print_md("Removed {} member(s) from prior circuits.".format(
            report["removed_from_existing"]))

    for line in report["lines"]:
        output.print_md("- {}".format(line))

    if report["panel_warnings"]:
        output.print_md("### Panel warnings")
        for key, msg in report["panel_warnings"]:
            output.print_md("- `{}`: {}".format(key, msg))

    if report["skipped_no_connector"]:
        output.print_md("### Skipped (no power connector)")
        output.print_md("These element ids had no power connector, so no native "
                        "circuit could include them (CKT_* params were still set):")
        output.print_md("`{}`".format(
            ", ".join(str(i) for i in report["skipped_no_connector"])))

    if report["errors"]:
        output.print_md("### Errors")
        for key, msg in report["errors"]:
            output.print_md("- `{}`: {}".format(key, msg))


main()
