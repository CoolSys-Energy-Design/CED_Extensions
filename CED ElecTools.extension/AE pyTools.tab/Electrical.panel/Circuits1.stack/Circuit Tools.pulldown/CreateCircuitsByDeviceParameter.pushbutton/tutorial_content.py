# -*- coding: utf-8 -*-
"""Narration for the Create Circuits by Device Parameter guided tour.

The tool groups circuitable devices by a parameter they already carry, lets you
review and correct the grouping, then creates one native Revit circuit per
group. The tour follows that order: scope, group, review, fix, run.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Create Circuits by Device Parameter - Show Me Around"
TOUR_KEY = "create_circuits_by_device_parameter"

HOST = {
    "kind": "window",
    "title": "Create Circuits by Device Parameter",
}

TARGET_LABELS = {
    "ActionStatusText": "the action status line",
    "BulkSetButton": "the Bulk Set button",
    "DedicatedGroupButton": "the Dedicated Circuit button",
    "DocumentText": "the document line",
    "ExpandToggle": "the expand/collapse toggle",
    "Grid": "the circuit grid",
    "GroupByCombo": "the Group By dropdown",
    "GroupHeaderBorder": "a circuit group header",
    "NameByCombo": "the Name By dropdown",
    "NewGroupButton": "the New Group button",
    "RefreshButton": "the Refresh button",
    "RunButton": "the Run button",
    "ScopeText": "the scope line",
    "SelectDeviceButton": "the Select Device button",
    "SelectGroupButton": "the Select Group button",
    "ShowDeviceButton": "the Show Device button",
    "SummaryText": "the summary line",
    "ToolInfoButton": "the tool info button",
    "ValidationText": "the validation line",
}

STEPS = [
    {
        "chapter": "Getting oriented",
        "title": "One circuit per group, from data you already have",
        "target": None,
        "body": (
            "This tool reads a parameter off your devices, buckets the devices that share "
            "a value, and creates one real Revit circuit per bucket. If your case "
            "controllers already carry a circuit number, the circuiting is already "
            "described - this just builds it."
        ),
        "bullets": [
            "It creates native Revit power circuits, not placeholders.",
            "The window is modeless - keep working in the model while it is open.",
            "Nothing is created until you press Run.",
            "It follows the active document and re-binds when you switch models.",
        ],
    },
    {
        "chapter": "Getting oriented",
        "title": "Select your devices first",
        "target": "ScopeText",
        "body": (
            "The tool works on a selection, not the whole model. Select the circuitable "
            "devices before you launch it - this line then tells you the scope it is "
            "holding and how many circuitable elements it found."
        ),
        "bullets": [
            "Launch it with nothing selected and it tells you to select devices and run "
            "again.",
            "Whole-model scope is deliberately disabled - this is a scoped tool.",
            "Only circuitable devices count; anything else in the selection is ignored.",
        ],
    },
    {
        "chapter": "Grouping", "title": "Group By is the whole idea",
        "target": "GroupByCombo",
        "body": (
            "Pick the device parameter that decides which devices belong on the same "
            "circuit. Devices sharing a value become one circuit."
        ),
        "bullets": [
            "It opens on the best available option - circuit number first, then a "
            "concatenated identity, then whatever else is there.",
            "CKT_Circuit Number_CEDT is the usual choice on a CED job.",
            "The parameter you group by is also projected into its own column so you can "
            "see the values you are grouping on.",
        ],
    },
    {
        "chapter": "Grouping", "title": "Name By names the circuits",
        "target": "NameByCombo",
        "body": (
            "Separately from grouping, this chooses which parameter seeds each circuit's "
            "Load Name. It defaults to the existing load name, then Identity Mark."
        ),
        "bullets": [
            "Careful: changing this re-seeds every circuit's Load Name and replaces any "
            "edits you typed by hand - you asked for a new naming source, so it gives you "
            "one.",
            "Set it before you start hand-editing names, not after.",
        ],
    },
    {
        "chapter": "Reviewing", "title": "Read the grid: headers are circuits",
        "target": "Grid",
        "body": (
            "The grid is grouped, and the grouping is the point. Each group header IS a "
            "circuit-to-be; the rows underneath are the devices that will go on it."
        ),
        "bullets": [
            "The header carries that circuit's Load Name, Panel and Breaker controls.",
            "Member rows are read-only information plus an include checkbox and a status.",
            "Untick a member to leave it off the circuit without deselecting it in Revit.",
            "Only rows that are included AND not already on a circuit get circuited.",
            "Any column header opens a show/hide menu if you want fewer columns.",
        ],
    },
    {
        "chapter": "Reviewing", "title": "Panel and breaker per circuit",
        "target": "GroupHeaderBorder",
        "body": (
            "Each circuit needs a panel and a rating. The tool guesses the panel by "
            "proximity - the closest panel with room - and you correct it where the guess "
            "is wrong."
        ),
        "bullets": [
            "'No listed panel has room - pick a panel manually' means the nearby panels "
            "are full.",
            "'None of the listed panels exist in the model' means pick one yourself.",
            "A non-standard breaker size is flagged as a warning, not blocked - verify it "
            "is intended before circuiting.",
        ],
    },
    {
        "chapter": "Reviewing", "title": "Expand or collapse everything",
        "target": "ExpandToggle",
        "body": (
            "Toggles every group open or shut in one pass. Collapsed is the right way to "
            "review panels and ratings across many circuits; expanded is for checking "
            "membership."
        ),
        "bullets": [],
    },
    {
        "chapter": "Fixing the grouping", "title": "Regroup by hand where the data is wrong",
        "target": "NewGroupButton",
        "body": (
            "The parameter data will not always be right. Rather than making you go and "
            "fix parameters first, you can move devices into a new group here and carry "
            "on."
        ),
        "bullets": [
            "New Group takes the selected member rows into a fresh circuit.",
            "Dedicated Circuit gives the selected device a circuit of its own - the "
            "one-device-one-breaker case.",
            "This changes the plan, not the model; the underlying parameters are "
            "untouched.",
        ],
    },
    {
        "chapter": "Fixing the grouping", "title": "Set many circuits at once",
        "target": "BulkSetButton",
        "body": (
            "Select several circuit rows and Bulk Set applies the same panel, rating, "
            "load name or schedule notes to all of them, instead of editing each header."
        ),
        "bullets": [
            "The optional increment box makes the load name count up across the selected "
            "circuits rather than repeating.",
            "Leave a field blank in the bulk editor to leave that value alone.",
        ],
    },
    {
        "chapter": "Fixing the grouping", "title": "Verify against the model",
        "target": "SelectDeviceButton",
        "body": (
            "Before you create anything, prove the grouping matches reality. These select "
            "and zoom to the devices and groups in Revit so you can look at them."
        ),
        "bullets": [
            "Select Device / Show Device work on the highlighted member row.",
            "Select Group / Show Group work on the whole circuit.",
            "Use these on any group whose panel guess surprised you - usually the device "
            "is somewhere you did not expect.",
        ],
    },
    {
        "chapter": "Running it", "title": "Validation blocks the real mistakes",
        "target": "ValidationText",
        "body": (
            "Run will not proceed while a circuit is invalid, and this line names what is "
            "wrong. These are the checks worth having - they are the ones Revit would "
            "otherwise fail on mid-transaction."
        ),
        "bullets": [
            "'Panel is not selected' - a circuit has no destination.",
            "'Invalid breaker rating' - the rating is not usable.",
            "'Pole mismatch' - the devices in one group disagree on pole count.",
            "'Voltage mismatch' - they disagree on voltage.",
            "A pole or voltage mismatch usually means the grouping parameter lumped "
            "together things that are not actually one circuit.",
        ],
    },
    {
        "chapter": "Running it", "title": "Run, and what it writes",
        "target": "RunButton",
        "body": (
            "Run creates one native Revit circuit per group, assigns the panel, sets the "
            "rating and load name, and mirrors the panel and rating back onto each "
            "member's CKT_ shared parameters so the CED data model stays in step."
        ),
        "bullets": [
            "Load Name is set on the circuit itself, not on each device.",
            "Revit warnings raised by circuit creation are swallowed; errors are not.",
            "Panel assignment is a separate step after creation, so you may be asked "
            "about capacity - including assigning to an unscheduled panel, or rolling "
            "back.",
            "The status line reports progress: creating circuits, assigning panels, "
            "committed.",
        ],
    },
    {
        "chapter": "Running it", "title": "Refresh, and where to read more",
        "target": "RefreshButton",
        "body": (
            "Refresh re-reads the Revit-backed fields while keeping your groups and your "
            "staged values, so you do not lose work after editing the model."
        ),
        "bullets": [
            "Use it after fixing parameters or adding a panel, not to start over.",
            "The info button in the window header holds the tool's own short reference.",
            "Today the tool knows about Case Controllers; the device registry is built to "
            "be extended to other equipment.",
        ],
    },
]
