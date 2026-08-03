# -*- coding: utf-8 -*-
__title__ = "Circuit Grouper"
__doc__ = ("Gather every circuitable element (any family instance with a power "
           "connector) - from the current selection if one exists, else the "
           "whole model - group them by a parameter you choose, validate "
           "voltage/pole compatibility, regroup, then create one native Revit "
           "circuit per group.")

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

    # If the user has a selection, operate only on those elements; otherwise
    # scan the whole model. Either way only circuitable elements are kept.
    sel_ids = None
    scope_label = "model"
    try:
        picked = list(revit.uidoc.Selection.GetElementIds())
    except Exception:
        picked = []
    if picked:
        sel_ids = [cg_collect.element_id_value(eid) for eid in picked]
        scope_label = "selection"
    else:
        # No selection -> whole-model scan. Warn before the (potentially slow)
        # collection and let the user back out and select first.
        proceed = forms.alert(
            "Nothing is selected, so Circuit Grouper will scan the ENTIRE "
            "model for circuitable elements.\n\n"
            "Run time may drastically increase - especially as you change the "
            "parameter to group by, since every element is regrouped.\n\n"
            "Select the elements you want to circuit first for a faster, more "
            "focused run.\n\n"
            "Scan the whole model anyway?",
            title=TITLE, yes=True, no=True,
        )
        if not proceed:
            return

    rows_data = cg_collect.collect_devices(doc, sel_ids)
    if not rows_data:
        if sel_ids is not None:
            forms.alert(
                "None of the selected elements are circuitable "
                "(no electrical power connector).",
                title=TITLE,
            )
        else:
            forms.alert(
                "No circuitable elements found in the model "
                "(nothing has an electrical power connector).",
                title=TITLE,
            )
        return

    group_param_options = cg_core.common_group_params(rows_data)
    default_group_param = cg_core.default_group_param(group_param_options)
    # name-by offers the same common-parameter list, with its own default
    default_name_param = cg_core.default_name_param(group_param_options)

    logger.debug("Circuit Grouper scope=%s, %d circuitable element(s)",
                 scope_label, len(rows_data))

    panel_names, name_to_id, panel_info = cg_collect.collect_panels(doc)

    plans, name_to_id = cg_window.show_window(
        rows_data, panel_names, name_to_id, cg_core.RATING_OPTIONS,
        group_param_options, default_group_param,
        panel_info=panel_info,
        name_param_options=group_param_options,
        default_name_param=default_name_param,
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
                        "circuit could include them (their CKT_ Panel/Rating "
                        "params were still set):")
        output.print_md("`{}`".format(
            ", ".join(str(i) for i in report["skipped_no_connector"])))

    if report["errors"]:
        output.print_md("### Errors")
        for key, msg in report["errors"]:
            output.print_md("- `{}`: {}".format(key, msg))


main()
