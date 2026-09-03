# -*- coding: utf-8 -*-
"""Narration for the System Tagger guided tour.

System Tagger walks a pasted list of system IDs and lets you click the refrigerated
cases belonging to each one. It can place tags, write Identity Mark, or both - and
it remembers where you stopped.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "System Tagger - Show Me Around"
TOUR_KEY = "system_tagger"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "One ID at a time, you do the pointing",
        "target": None,
        "body": (
            "System Tagger takes a list of system IDs from Excel and steps through "
            "them. For each ID it asks you to click the cases that belong to it, then "
            "labels them. You supply the knowledge; it supplies the bookkeeping."
        ),
        "bullets": [
            "Built for refrigerated cases, but it picks any Mechanical Equipment.",
            "It is a long job by nature - which is why it can be paused and resumed.",
            "The active view cannot be a template; open a working view first.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "Paste the SYS NO. column",
        "target": None,
        "body": (
            "It starts by asking you to paste the SYS NO. values from Excel. You can "
            "paste the whole block - it pulls the IDs out."
        ),
        "bullets": [
            "'No System IDs found in the pasted text' means it could not read any - "
            "check you copied the right column.",
            "The order you paste is the order it walks.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "Choose what it should write",
        "target": None,
        "body": (
            "You then choose what happens to each element you pick. This is the "
            "decision that changes what the run actually produces."
        ),
        "bullets": [
            "Tag + Identity Mark - sets Identity Mark and places a tag.",
            "Identity Mark only - writes the parameter, places no tag.",
            "Tag only - places a tag and leaves Identity Mark alone.",
            "For anything that places a tag you also pick the tag type; M_Mechanical "
            "Eqmt Tag is the usual one.",
        ],
    },
    {
        "chapter": "The picking loop",
        "title": "Click to select, click again to deselect",
        "target": None,
        "body": (
            "For each ID the status bar reads 'System ID <x>: pick Mechanical "
            "Equipment (ESC to finish)'. Click the cases for that ID. Clicking an "
            "already-picked case removes it."
        ),
        "bullets": [
            "Picked elements are highlighted so you can see your set building up.",
            "Press ESC to finish that ID and move to the next one.",
            "It confirms with 'System ID X: N case(s) selected.'",
        ],
    },
    {
        "chapter": "The picking loop",
        "title": "Multiple cases get letters",
        "target": None,
        "body": (
            "If one system ID covers several cases, the label gets a letter suffix - "
            "A, B, C and so on - so each case ends up uniquely identified."
        ),
        "bullets": [
            "A single case for an ID gets the plain ID with no letter.",
            "That suffix behaviour is why you should pick all the cases for an ID "
            "before pressing ESC.",
        ],
    },
    {
        "chapter": "Pausing and resuming",
        "title": "You can stop and come back",
        "target": None,
        "body": (
            "Because these runs are long, the tool saves its position in the model "
            "and offers to resume next time."
        ),
        "bullets": [
            "Stopping mid-run leaves 'System Tagger paused. Run again to resume.'",
            "Next run it asks 'Resume System Tagger at <ID>? (n of N)' - or Start "
            "Over.",
            "Resume state lives in a DataStorage element in the model, so it survives "
            "closing Revit. Older runs stored it on Project Information and are "
            "migrated automatically.",
        ],
    },
    {
        "chapter": "Afterwards",
        "title": "Check the failures",
        "target": None,
        "body": (
            "Individual elements can fail without stopping the run, and those are "
            "reported rather than swallowed."
        ),
        "bullets": [
            "'Failed to set Identity Mark for X' - usually the parameter is missing "
            "or read-only on that family.",
            "'Failed to place tag for X' - usually the tag type is not loaded or not "
            "valid for that category.",
            "A preview highlight pass runs before committing so you can see what was "
            "picked.",
        ],
    },
]
