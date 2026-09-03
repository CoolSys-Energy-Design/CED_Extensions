# -*- coding: utf-8 -*-
"""Narration for the Circuit Tools dropdown guided tour.

A one-line-per-tool map of the Circuit Tools pulldown, in the order the menu lists
them, so you can tell at a glance which button you want.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Circuit Tools - What's In Here"
TOUR_KEY = "circuit_tools_dropdown"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "Overview",
        "title": "Eleven tools, four jobs",
        "target": None,
        "body": (
            "Circuit Tools is the drawer for everything you do TO circuits that "
            "already exist, plus the handful of tools that create them. The menu is "
            "grouped: edit and find, move and rename, create, then project setup."
        ),
        "bullets": [
            "Most of these work off your current selection - devices or circuits.",
            "Anything that writes to the model shows you a review table first.",
            "Shift+Click any individual button that has its own tour for more detail.",
        ],
    },
    {
        "chapter": "Edit and find",
        "title": "Edit Circuit Properties",
        "target": None,
        "body": (
            "Opens the circuit property editor for the circuits you selected. With "
            "nothing selected it asks you to pick circuits from a list first."
        ),
        "bullets": [
            "Stage edits per circuit, preview the recalculated values, then apply all "
            "at once.",
            "This is the safe way to change several circuits without opening each "
            "panel schedule.",
        ],
    },
    {
        "chapter": "Edit and find",
        "title": "Circuited Device Finder",
        "target": None,
        "body": (
            "Give it circuits and it finds the devices connected to them, then opens "
            "a view and selects those devices in the model."
        ),
        "bullets": [
            "With no circuits preselected it prompts you with the circuit picker.",
            "Use it to answer 'what is actually on this breaker?' without tracing "
            "wires.",
        ],
    },
    {
        "chapter": "Move and rename",
        "title": "Batch Swap Circuits",
        "target": None,
        "body": (
            "Stage and apply many slot-level circuit transfers between panel "
            "schedules in one pass, instead of dragging circuits one at a time."
        ),
        "bullets": [
            "Choose source and target slots, review the sequence order, then apply.",
            "Handles circuits, default SPARE/SPACE entries, or a mix of both in one "
            "run.",
        ],
    },
    {
        "chapter": "Move and rename",
        "title": "Move Selected Circuits",
        "target": None,
        "body": (
            "Moves the circuits you selected to a target panel. It takes the "
            "selection from the model, a panel schedule, or the Circuit Manager "
            "workflow."
        ),
        "bullets": [
            "Works out slot availability for you.",
            "Restores default SPARE/SPACE entries after the transfer where it can.",
        ],
    },
    {
        "chapter": "Move and rename",
        "title": "Merge Circuits",
        "target": None,
        "body": (
            "Pick a main circuit, then pick one or more circuits to fold into it. The "
            "elements move onto the main circuit and the now-empty circuits are "
            "deleted."
        ),
        "bullets": [
            "Circuits have to be compatible - same system type, voltage and pole "
            "count.",
            "Spares and spaces are excluded from merging.",
        ],
    },
    {
        "chapter": "Move and rename",
        "title": "Rename Circuits by Device Parameter",
        "target": None,
        "body": (
            "Builds circuit names from the devices on them. Start from a selection of "
            "devices (or circuits) and it gives you a per-circuit string builder."
        ),
        "bullets": [
            "Compose names from parameter tokens, custom text and separators.",
            "Apply the active builder to the selected circuits or to all of them.",
            "The preview is editable per circuit and survives switching between "
            "circuits.",
            "Save and reload builder templates from your own pyRevit config.",
            "Optional duplicate handling appends ' #2', ' #3' and so on.",
            "If devices on one circuit disagree, the most common value wins; ties go "
            "to the lowest element id.",
        ],
    },
    {
        "chapter": "Create circuits",
        "title": "Create Circuits by Device Parameter",
        "target": None,
        "body": (
            "Collects circuitable devices, groups them by a device parameter such as "
            "CKT_Circuit Number_CEDT, and creates one native Revit circuit per group."
        ),
        "bullets": [
            "Review and regroup the results before anything is created.",
            "Choose the panel, breaker rating and load values for the new circuits.",
            "Select devices first for a focused run; with no selection it can scan "
            "the whole model after you confirm.",
        ],
    },
    {
        "chapter": "Create circuits",
        "title": "Create Dedicated Circuits",
        "target": None,
        "body": (
            "One dedicated circuit per selected element, all assigned to a panel you "
            "pick."
        ),
        "bullets": [
            "The results table lets you review and edit the key circuit values "
            "afterwards.",
            "This is the fast path for equipment that each needs its own breaker.",
        ],
    },
    {
        "chapter": "Create circuits",
        "title": "Add Spares/Spaces",
        "target": None,
        "body": (
            "Batch add or remove default SPARE and SPACE entries across the panel "
            "schedules you choose."
        ),
        "bullets": [
            "Quick Spare / Quick Space give you fill-all or remove-all behaviour.",
            "Or use the main table to review panels, pick a mode per panel, and apply "
            "staged actions.",
        ],
    },
    {
        "chapter": "Setup",
        "title": "Load Electrical Parameters",
        "target": None,
        "body": (
            "Loads the shared parameters the Calculate Circuits workflow depends on. "
            "Run it once per project, early."
        ),
        "bullets": [
            "If circuit calculation or the Circuit Manager behaves as though "
            "parameters are missing, this is the first thing to run.",
        ],
    },
]
