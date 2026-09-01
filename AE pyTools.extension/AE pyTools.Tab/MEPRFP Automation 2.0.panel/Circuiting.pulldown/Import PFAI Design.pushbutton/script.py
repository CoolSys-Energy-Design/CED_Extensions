#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Import PFAI Design"""

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
import pfai_import
import pfai_xlsx

TITLE = "Import PFAI Design (MEPRFP 2.0)"

# Types the CoolSys template does not ship but a PFAI plan asks for, and the
# type each is duplicated from. Duplicating inherits the source amperage and
# device configuration, so anything created this way is reported, not silent.
FALLBACK_TYPES = {
    "Fused - 60A": "Fused - 200A",
    "Duplex Wall - USB": "Duplex Wall",
}


def _md_table(out, columns, rows):
    """Render a table as markdown.

    Not `output.print_table` - that calls `itertools.izip_longest`, which is
    Python 2 only and does not exist on the CPython engine this script runs on.
    The other python3 buttons in this tab build their tables the same way.
    """
    def cell(v):
        return str(v).replace("|", "\\|").replace("\n", " ")
    out.print_md("| " + " | ".join(cell(c) for c in columns) + " |")
    out.print_md("|" + "|".join(["---"] * len(columns)) + "|")
    for r in rows:
        out.print_md("| " + " | ".join(cell(v) for v in r) + " |")


def _report(out, plan, result, path):
    out.print_md("# Import PFAI Design")
    out.print_md("**Workbook:** `%s`" % os.path.basename(path))
    c = plan.counts
    out.print_md(
        "Read %d panels, %d devices, %d circuits, %d keynotes."
        % (c["panels"], c["devices"], c["circuits"], c["keynotes"]))
    if result.swept:
        out.print_md("Cleared **%d** elements from a previous PFAI import."
                     % result.swept)

    rows = [
        ["Panels named / distribution systems set", result.panels_named,
         len(plan.panels)],
        ["Devices placed", result.placed, c["devices"]],
        ["Circuits assigned to a panel", result.circuits_ok, c["circuits"]],
        ["Keynotes placed", result.keynotes, c["keynotes"]],
        ["Keynotes showing their number on the sheet",
         result.keynotes_printed, c["keynotes"]],
        ["Circuit tags placed", result.tags, result.placed],
    ]
    _md_table(out, ["Pass", "Done", "Expected"], rows)

    if result.types_created:
        out.print_md("### Fixture types created")
        out.print_md(
            "These were not in the document. Each was duplicated from the type "
            "shown, so it inherits that type's amperage and device settings - "
            "check them before the sheet is issued.")
        _md_table(out, ["Created", "Duplicated from"],
                  [[t, s] for t, s in result.types_created])

    if result.problems:
        out.print_md("### %d item(s) need a look" % len(result.problems))
        _md_table(out, ["Pass", "Subject", "Reason"], result.problems)
        out.print_md(
            "_A circuit that reports **out of slots** or **panel and circuit do "
            "not match** is a design problem, not an import failure: the panel "
            "is full, or the load's connector voltage and poles disagree with "
            "the panel's distribution system._")
    else:
        out.print_md("### No problems reported.")

    out.print_md(
        "Everything created is stamped `Comments = \"PFAI\"` (keynotes use "
        "`Keynote Category_CEDT`). The whole import is one undo step, and "
        "re-running with **Replace** clears it first.")


def main():
    out = script.get_output()
    out.close_others()
    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return
    view = revit.active_view
    if getattr(view, "GenLevel", None) is None:
        forms.alert(
            "Open the E101 power plan floor plan first - the active view has no "
            "level to place devices on.", title=TITLE)
        return

    path = forms.pick_file(file_ext="xlsx",
                           title="Pick the PFAI design workbook")
    if not path:
        return

    try:
        plan = pfai_import.read_plan(path)
    except pfai_xlsx.XlsxError as exc:
        forms.alert(str(exc), title=TITLE)
        return

    if not plan.devices:
        forms.alert("The Devices sheet is empty - nothing to import.",
                    title=TITLE)
        return

    c = plan.counts
    sweep = forms.confirm(
        "Import %d devices, %d circuits and %d keynotes into '%s'?\n\n"
        "Yes  - replace any previous PFAI import first (recommended)\n"
        "No   - add alongside whatever is already there"
        % (c["devices"], c["circuits"], c["keynotes"],
           getattr(view, "Name", "the active view")),
        title=TITLE)

    result = pfai_import.run_import(
        doc, view, plan, sweep=bool(sweep), fallback_types=FALLBACK_TYPES)
    _report(out, plan, result, path)


if __name__ == "__main__":
    main()
