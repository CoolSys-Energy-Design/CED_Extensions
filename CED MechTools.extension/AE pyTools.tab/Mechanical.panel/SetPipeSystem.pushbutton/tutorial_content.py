# -*- coding: utf-8 -*-
"""Narration for the Set Piping System guided tour.

Set Piping System forces a system type onto segments you select, including segments
that have no system yet. The tour explains the two very different things it does
depending on whether the network is already defined.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Set Piping System - Show Me Around"
TOUR_KEY = "set_piping_system"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "Force a system onto what you selected",
        "target": None,
        "body": (
            "Revit decides system membership by topology, which makes it awkward to "
            "say 'these segments belong to this system'. This tool does that: it "
            "forces a piping or duct system onto the segments you pick, and creates "
            "systems on segments that do not have one yet."
        ),
        "bullets": [
            "Works on pipes, fittings and fixtures - or on ducts.",
            "It edits topology, so it is not a parameter-only change. Read the "
            "output.",
            "Everything runs inside transactions you can undo as one step.",
        ],
    },
    {
        "chapter": "Using it",
        "title": "Select first, then run",
        "target": None,
        "body": (
            "Select the pipes, fittings and fixtures (or the ducts) you want in the "
            "system, then run the tool. It reports the counts it found before doing "
            "anything."
        ),
        "bullets": [
            "With nothing selected it exits and says so.",
            "Mixed selections are rejected - pipes or ducts, one or the other, not "
            "both.",
            "It prints the selected element, pipe and duct counts so you can confirm "
            "it read your selection the way you meant it.",
        ],
    },
    {
        "chapter": "Using it",
        "title": "Pick the system type",
        "target": None,
        "body": (
            "You are then asked for the target system type - Select Piping System "
            "Type, or Select Duct System Type depending on what you selected."
        ),
        "bullets": [
            "This is the classification the segments end up carrying.",
        ],
    },
    {
        "chapter": "What happens next",
        "title": "Ducts take the simple path",
        "target": None,
        "body": (
            "For ducts the tool sets the duct system type parameter directly and "
            "reports how many succeeded and failed. No topology surgery."
        ),
        "bullets": [
            "Straightforward and quick - the interesting behaviour is on the piping "
            "side.",
        ],
    },
    {
        "chapter": "What happens next",
        "title": "Undefined piping: create and populate",
        "target": None,
        "body": (
            "Where the selected pipes have no system at all, the tool creates a "
            "piping system and adds the eligible connectors to it one at a time."
        ),
        "bullets": [
            "Connectors are added individually so one bad connector - already used, "
            "or wrongly classified - does not fail the whole run.",
            "If your target classification is too strict for the first connector, it "
            "falls back through other seed types until one is accepted.",
            "Where Revit can propagate on its own (through a tee or a fixture) it "
            "sets the pipe parameter and lets Revit spread it, the way Systemizer "
            "does.",
            "You may see 'Undefined system creation skipped: no eligible connectors "
            "found' - that means nothing in the selection could seed a system.",
        ],
    },
    {
        "chapter": "What happens next",
        "title": "Defined piping: cut, divide, reconnect",
        "target": None,
        "body": (
            "Where the pipes already belong to a system, the tool has to separate "
            "your selection from the rest of the network. It does that in four "
            "visible stages, and it puts the connections back afterwards."
        ),
        "bullets": [
            "Collecting boundary connections - finds where your selection meets "
            "everything else.",
            "Disconnecting boundary pairs - breaks those joints, preferring the "
            "MEPCurve side so endpoint and system ownership survive.",
            "Dividing affected systems - Revit re-splits the systems that were cut.",
            "Reconnecting boundary pairs - the joints go back together.",
            "The connection points are saved before the cut and used to find the "
            "connectors again afterwards, because element state changes in between.",
        ],
    },
    {
        "chapter": "Afterwards",
        "title": "Read the evaluation block",
        "target": None,
        "body": (
            "The run ends with an evaluation section in the output panel. That is "
            "where the tool tells you what it decided and what actually took."
        ),
        "bullets": [
            "'Could not create piping system' means the selection could not seed one "
            "- usually no open, compatible connector.",
            "Check the resulting system membership in a schedule or the system "
            "browser before you move on.",
            "If the result is wrong, undo is one step - then reselect and try a "
            "different system type.",
        ],
    },
]
