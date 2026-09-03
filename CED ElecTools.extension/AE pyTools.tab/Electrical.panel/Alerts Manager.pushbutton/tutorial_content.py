# -*- coding: utf-8 -*-
"""Narration for the Alerts Manager guided tour.

Alerts Manager is the triage bench for saved circuit alerts. The tour covers both
displays, the model-selection buttons, hiding alert types, and why Refresh sometimes
greys out.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Alerts Manager - Show Me Around"
TOUR_KEY = "alerts_manager"

HOST = {
    "kind": "window",
    "tag": "_ae_alerts_browser_window_persistent_v1",
    "title": "Alerts Manager",
}

TARGET_LABELS = {
    "ActiveAlertsList": "the Active alerts grid",
    "AlertsTabs": "the Active / Hidden tabs",
    "AllAlertsList": "the All Alerts grid",
    "CircuitCountText": "the circuit count",
    "CircuitList": "the circuit list",
    "DocumentText": "the document line",
    "HiddenAlertsList": "the Hidden alerts grid",
    "HideUnhideButton": "the Hide Type / Unhide Type button",
    "RefreshButton": "the Refresh button",
    "SelectedCircuitText": "the selected circuit line",
    "ShowHiddenToggle": "the Show Hidden switch",
    "StatusText": "the status line",
    "ToggleViewButton": "the Grid Display toggle",
}

STEPS = [
    {
        "chapter": "Getting oriented",
        "title": "What this window is for",
        "target": None,
        "body": (
            "Alerts Manager collects the alerts that circuit calculation saved onto "
            "your circuits and puts them in one modeless window, so you can work down "
            "them instead of rediscovering them one circuit at a time."
        ),
        "bullets": [
            "Modeless - leave it open and fix things in the model while it sits "
            "there.",
            "It reads alert payload off the circuits; it does not invent new checks.",
            "Circuit Manager and Circuit Browser can both launch it.",
        ],
    },
    {
        "chapter": "Getting oriented",
        "title": "Which model, and how many circuits",
        "target": "DocumentText",
        "body": (
            "The header names the active document and, next to it, how many circuits "
            "currently carry alert payload. That count is your workload for this "
            "session."
        ),
        "bullets": [
            "The header summary shows total Alerts and Hidden counts independently, "
            "so Show Hidden does not change the numbers you are reporting against.",
        ],
    },
    {
        "chapter": "Two ways to look at it",
        "title": "By circuit, or one flat grid",
        "target": "ToggleViewButton",
        "body": (
            "Grid Display switches between the two displays. By-circuit is for "
            "working a single circuit properly; All Alerts is for scanning and "
            "sorting the whole set."
        ),
        "bullets": [
            "By circuit - the left picker plus the Active / Hidden tabs on the right.",
            "All Alerts - one row per alert, no circuit picker, with Panel, Circuit, "
            "Load Name, Severity, Group, Alert ID and Message columns.",
            "Click any column header in All Alerts to sort by it.",
            "Sort by Severity to work the worst first; sort by Group to fix one class "
            "of problem at a time.",
        ],
    },
    {
        "chapter": "Two ways to look at it",
        "title": "The circuit list",
        "target": "CircuitList",
        "body": (
            "Single-select list of every circuit carrying alerts. Each row shows the "
            "panel and circuit, the load name, and the alert counts - total, active "
            "and hidden."
        ),
        "bullets": [
            "Selecting a circuit here drives everything on the right-hand side.",
        ],
    },
    {
        "chapter": "Reading the alerts",
        "title": "Active and Hidden tabs",
        "target": "AlertsTabs",
        "body": (
            "Active is what is wrong now. Hidden is what somebody deliberately "
            "suppressed for this circuit. Both grids show Severity, Group, Alert ID "
            "and Message."
        ),
        "bullets": [
            "Severity tells you whether it blocks an issue or is just worth knowing.",
            "Alert ID is the stable name - the thing to quote when you ask about one.",
            "Hidden rows appear in subdued italic text when Show Hidden is on.",
        ],
    },
    {
        "chapter": "Reading the alerts",
        "title": "Hiding an alert type you have accepted",
        "target": "HideUnhideButton",
        "body": (
            "Some alerts are a deliberate engineering decision, not a mistake. Select "
            "the alert type and use Hide Type so it stops competing for attention on "
            "that circuit. Unhide Type puts it back."
        ),
        "bullets": [
            "Hiding is per circuit, not global - another circuit still reports it.",
            "Only mapped, approved alert definitions can be hidden.",
            "Hideable types include Design.NonStandardOCPRating, "
            "Design.BreakerLugSizeLimitOverride, "
            "Design.BreakerLugQuantityLimitOverride, Design.CircuitLoadsNull, "
            "Calculations.BreakerLugSizeLimit and "
            "Calculations.BreakerLugQuantityLimit.",
            "Hide the type when you have decided it is acceptable - not to make a "
            "report look clean.",
        ],
    },
    {
        "chapter": "Fixing things",
        "title": "Find the element in the model",
        "target": "SelectedCircuitText",
        "body": (
            "The bottom command row selects things in Revit for the circuit you have "
            "chosen, so you can go and look at the actual problem."
        ),
        "bullets": [
            "Panel - selects the upstream panel.",
            "Circuit - selects the circuit element itself.",
            "Device - selects the downstream connected devices.",
            "X - clears the Revit selection between passes.",
        ],
    },
    {
        "chapter": "Fixing things",
        "title": "Refresh recalculates just that circuit",
        "target": "RefreshButton",
        "body": (
            "After you fix something in the model, Refresh recalculates the selected "
            "circuit and rebuilds its alert state. That is how you confirm an alert "
            "is actually gone rather than merely addressed."
        ),
        "bullets": [
            "It recalculates one circuit, not the model - it is safe to use "
            "repeatedly.",
            "Load- and sizing-related alerts always need a recalculation before they "
            "clear.",
        ],
    },
    {
        "chapter": "When it will not play",
        "title": "Why Refresh is greyed out",
        "target": None,
        "body": (
            "Three reasons, and the disabled button's tooltip tells you which."
        ),
        "bullets": [
            "No circuit is selected.",
            "An operation is already running.",
            "Recalculation is blocked by ownership or worksharing - someone else owns "
            "the circuit or a downstream element.",
            "Hover the disabled button to read the actual reason.",
        ],
    },
    {
        "chapter": "When it will not play",
        "title": "Working it efficiently",
        "target": None,
        "body": (
            "Filter in Circuit Manager, triage here. The two windows are meant to be "
            "used together."
        ),
        "bullets": [
            "Use Circuit Manager's 'Circuits With Alerts' filter to build the set, "
            "then work it here.",
            "Start with high-severity alerts; they often clear several lesser ones.",
            "Re-run calculation after any corrective action that changes load or "
            "sizing.",
        ],
    },
]
