# -*- coding: utf-8 -*-
"""Narration for the WM Tools dropdown guided tour.

What each tool in the WM Tools pulldown does. These are the Whole Foods /
refrigeration power-connector tools, and most of them key off the Refrigeration
Circuit Number_CEDT parameter.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "WM Tools - What's In Here"
TOUR_KEY = "wm_tools_dropdown"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "Overview",
        "title": "The refrigeration power toolkit",
        "target": None,
        "body": (
            "WM Tools automates the connector-and-circuit busywork on refrigeration "
            "jobs. Nearly everything here is driven by one parameter: Refrigeration "
            "Circuit Number_CEDT."
        ),
        "bullets": [
            "If a tool ignores an element, the usual reason is that parameter is "
            "blank.",
            "Run Load Electrical Content first on a new project - the rest depends on "
            "it.",
        ],
    },
    {
        "chapter": "Power groups",
        "title": "Place Power Groups",
        "target": None,
        "body": (
            "Places power groups on top of other elements. With nothing selected it "
            "targets every instance of the EMS case tag in the active view."
        ),
        "bullets": [
            "Shift+Click the button to place groups at an offset distance instead of "
            "on top.",
            "It is not restricted to EMS tags - you can place groups on almost any "
            "element.",
            "If the reference element has no 'Circuit #' parameter you get a warning "
            "and the group's Refrigeration Circuit Number_CEDT reads 'Circuit # Not "
            "Found'.",
        ],
    },
    {
        "chapter": "Power groups",
        "title": "Ungroup Power Groups",
        "target": None,
        "body": (
            "Ungroups the selected power groups and pushes the group's Refrigeration "
            "Circuit Number_CEDT value down onto the elements that were inside it."
        ),
        "bullets": [
            "The point is the parameter hand-off - the value survives the ungroup.",
        ],
    },
    {
        "chapter": "From Excel",
        "title": "Create Panels From Excel",
        "target": None,
        "body": (
            "Creates panel and switchboard instances in Revit from the Panel Creation "
            "workbook."
        ),
        "bullets": [
            "Alt+Click the button to get a copy of the Panel Creation workbook.",
            "Copy that workbook into the project's 03 Calculations / Electrical "
            "folder and rename it for the project before you fill it in.",
        ],
    },
    {
        "chapter": "From Excel",
        "title": "Create Circuits From Excel",
        "target": None,
        "body": (
            "The companion to Create Panels From Excel: reads the workbook and "
            "creates the circuits in the panels."
        ),
        "bullets": [
            "Run it after the panels exist.",
            "Its tooltip has no description written yet - the workbook is the "
            "documentation.",
        ],
    },
    {
        "chapter": "Connectors and circuits",
        "title": "Select Connectors By Case",
        "target": None,
        "body": (
            "Selects the connector elements for refrigeration cases, filtered by "
            "Refrigeration Circuit Number_CEDT and by connector type such as FANS or "
            "LIGHTS."
        ),
        "bullets": [
            "Use it to set case data or build circuits across many cases at once.",
            "This is the selection tool the other WM tools expect you to use first.",
        ],
    },
    {
        "chapter": "Connectors and circuits",
        "title": "Replace Existing Circuit",
        "target": None,
        "body": (
            "Swaps an existing placeholder circuit for real connector elements. "
            "Select a placeholder in a panel schedule, choose matching connector "
            "types, and they are added to the circuit while the placeholder is "
            "deleted."
        ),
        "bullets": [
            "Filters out anything that does not match the circuit's voltage and "
            "poles.",
            "Works with any CED Standard electrical fixture, not just refrigeration "
            "connectors - the only requirement is a value in Refrigeration Circuit "
            "Number_CEDT.",
            "So to swap in a receptacle or a disconnect, give it any value in that "
            "parameter and it appears as an option.",
        ],
    },
    {
        "chapter": "Display and setup",
        "title": "Toggle Connector Symbols",
        "target": None,
        "body": (
            "Flips the Symbol Visible_CED type parameter on every type of the EF- "
            "U_Refrig Power Connector-Balanced_CED-WM family."
        ),
        "bullets": [
            "Use it to show the connector symbols while you work and hide them for "
            "issue.",
        ],
    },
    {
        "chapter": "Display and setup",
        "title": "Load Electrical Content",
        "target": None,
        "body": (
            "Loads all the families and content WM Tools needs to work."
        ),
        "bullets": [
            "Run this first on a new project.",
            "If a WM tool cannot find a family it expects, run this before debugging "
            "anything else.",
        ],
    },
]
