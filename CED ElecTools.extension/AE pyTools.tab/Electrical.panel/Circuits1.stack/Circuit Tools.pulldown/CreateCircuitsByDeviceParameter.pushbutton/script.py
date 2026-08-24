# -*- coding: utf-8 -*-
__title__ = "Create Circuits by Device Parameter"
__doc__ = ("Gather every circuitable element (any family instance with a power "
           "connector) - from the current selection if one exists, else the "
           "whole model - group them by a parameter you choose, validate "
           "voltage/pole compatibility, regroup, then create one native Revit "
           "circuit per group. Group names containing DEDICATED create one "
           "native circuit per effective member.")

import os
import sys

from pyrevit import revit, forms, script, DB
from Snippets import revit_helpers

# make the sibling cg_* modules importable
THIS_DIR = os.path.dirname(__file__)
if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

import cg_core
import cg_window
import cg_collect
import cg_apply

logger = script.get_logger()
output = script.get_output()
TITLE = "Create Circuits by Device Parameter"


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    # If the user has a selection, operate only on those elements; otherwise
    # scan the whole model. Either way only circuitable elements are kept.
    sel_ids = None
    scope_label = "model"
    picked = list(revit.uidoc.Selection.GetElementIds())
    if picked:
        # Keep the Revit ElementId objects intact through collection. The
        # collector only normalizes ids when it builds the UI data model.
        sel_ids = picked
        scope_label = "selection"
    else:
        # No selection -> whole-model scan. Warn before the (potentially slow)
        # collection and let the user back out and select first.
        proceed = forms.alert(
            "Nothing is selected, so Create Circuits by Device Parameter will scan the ENTIRE "
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

    parameter_options = cg_core.common_group_params(rows_data)
    # Keep the two display/location groupings at the top of the Group By
    # picker. The raw Identity Mark parameter remains available to Name By,
    # while Group By uses the same concatenated identity shown in the grid.
    group_param_options = [
        cg_core.SPACE_GROUP_OPTION,
        cg_core.IDENTITY_GROUP_OPTION,
    ] + [p for p in parameter_options if p != "Identity Mark"]
    default_group_param = cg_core.default_group_param(group_param_options)
    # name-by offers the same common-parameter list, with its own default
    default_name_param = cg_core.default_name_param(parameter_options)

    logger.debug("Create Circuits by Device Parameter scope=%s, %d circuitable element(s)",
                 scope_label, len(rows_data))

    panel_names, name_to_id, panel_info = cg_collect.collect_panels(doc)

    plans, name_to_id = cg_window.show_window(
        rows_data, panel_names, name_to_id, cg_core.RATING_OPTIONS,
        group_param_options, default_group_param,
        panel_info=panel_info,
        name_param_options=parameter_options,
        default_name_param=default_name_param,
    )

    if not plans:
        return  # cancelled or nothing valid

    # Create the native circuits first, then pass the created circuits through
    # the same Move Selected Circuits service used by the ribbon tool. That
    # service owns the complete capacity workflow: fit_without/fit_with,
    # removable default SPARE/SPACE evaluation, confirmation, temporary
    # removal, assignment, and restoration/backfill.
    workflow_group = DB.TransactionGroup(doc, "Create Circuits by Device Parameter")
    workflow_group.Start()
    try:
        report = cg_apply.run(doc, plans, name_to_id, logger)
        assignment = cg_apply.assign_created_circuits_to_panels(
            doc, report.get("created_circuit_ids_by_panel", {}), name_to_id, logger)
        report["panel_assignment"] = assignment
        # Assignment failures are intentionally reported by Create Circuits by Device Parameter
        # while preserving the created circuits. The outer group therefore
        # assimilates the successful creation and any successful panel moves
        # into one undo item.
        workflow_group.Assimilate()
    except Exception:
        try:
            workflow_group.RollBack()
        except Exception:
            pass
        raise

    for event in list(assignment.get("buffered_events", []) or []):
        if not event:
            continue
        if event[0] == "md":
            output.print_md(event[1])
        elif event[0] == "table":
            output.print_table(event[1], event[2])

    # -- report --------------------------------------------------------
    output.print_md("## Create Circuits by Device Parameter - Results")
    output.print_md("**Circuits created:** {}  |  **Members circuited:** {}".format(
        report["created"], report["members_circuited"]))
    if assignment.get("moved"):
        output.print_md("Assigned {} created circuit(s) to their selected panels.".format(
            assignment["moved"]))
    if assignment.get("fallback_used"):
        output.print_md(
            "Move Selected Circuits used its default SPARE/SPACE replacement workflow.")
    status_rows = list(assignment.get("circuit_status") or [])
    if status_rows:
        def _linkify_created_circuit(row):
            label = row.get("circuit", "Created circuit")
            try:
                element_id = revit_helpers.elementid_from_value(
                    int(row.get("element_id")))
                return output.linkify(element_id, label)
            except Exception:
                return label

        def _assignment_row_values(row):
            return [
                _linkify_created_circuit(row),
                row.get("element_id", "-"),
                row.get("target_panel", "-"),
                row.get("actual_panel", "-"),
                row.get("status", "-"),
            ]

        verified_count = sum(
            1 for row in status_rows if row.get("status") == "ASSIGNED")
        output.print_md("### Actual circuit assignment")
        output.print_md(
            "Verified on selected panel: {} of {}.".format(
                verified_count, len(status_rows)))
        output.print_table(
            [_assignment_row_values(row) for row in status_rows],
            ["Circuit", "Element ID", "Target panel", "Actual panel", "Result"],
        )
        unresolved_rows = [
            row for row in status_rows if row.get("status") != "ASSIGNED"]
        if unresolved_rows:
            output.print_md("### Created circuits requiring manual resolution")
            output.print_md(
                "Select a linked circuit below to resolve its panel assignment manually.")
            output.print_table(
                [_assignment_row_values(row) for row in unresolved_rows],
                ["Circuit", "Element ID", "Target panel", "Actual panel", "Result"],
            )
    if assignment.get("errors"):
        output.print_md("### Panel assignment issues")
        for panel, message in assignment["errors"]:
            output.print_md("- `{}`: {}".format(panel, message))

    if assignment.get("not_on_target") or assignment.get("errors"):
        alert_lines = [
            "Circuits were created, but some did not land on their selected panel.",
            "",
        ]
        if assignment.get("not_on_target"):
            alert_lines.append(
                "{} circuit(s) did not land on the selected panel.".format(
                    assignment["not_on_target"]))
        if assignment.get("unassigned"):
            alert_lines.append(
                "{} of those circuit(s) have no panel assignment.".format(
                    assignment["unassigned"]))
        if assignment.get("errors"):
            alert_lines.append("Panel assignment issues:")
            for panel, message in assignment["errors"]:
                alert_lines.append("- {}: {}".format(panel, message))
        alert_lines.append("")
        alert_lines.append(
            "Created circuits were retained. See the actual assignment table in the pyRevit output.")
        forms.alert("\n".join(alert_lines), title=TITLE)

    if report["lines"]:
        output.print_md("### Circuit creation")
    for line in report["lines"]:
        output.print_md("- {}".format(line))

    if report["skipped_no_connector"]:
        output.print_md("### Skipped (no primary power connector)")
        output.print_md("These element ids had no primary power connector, so "
                        "the tool left them unchanged:")
        output.print_md("`{}`".format(
            ", ".join(str(i) for i in report["skipped_no_connector"])))

    if report["skipped_unavailable_primary"]:
        output.print_md("### Skipped (primary power connector already in use)")
        output.print_md("These element ids already had a system on their primary "
                        "power connector, so the tool left them unchanged:")
        output.print_md("`{}`".format(
            ", ".join(
                str(i) for i in report["skipped_unavailable_primary"])))

    if report["errors"]:
        output.print_md("### Errors")
        for key, msg in report["errors"]:
            output.print_md("- `{}`: {}".format(key, msg))


main()
