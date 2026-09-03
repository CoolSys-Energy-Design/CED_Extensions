# -*- coding: utf-8 -*-
"""Narration for the ConnectTo guided tour.

ConnectTo joins one MEP element to another by moving the first to meet the second.
Two picks, and where you click matters.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "ConnectTo - Show Me Around"
TOUR_KEY = "connect_to"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "Two picks and they are joined",
        "target": None,
        "body": (
            "ConnectTo moves one element so its connector meets another element's "
            "connector, and joins them. It saves the align-rotate-nudge dance you "
            "would otherwise do by hand."
        ),
        "bullets": [
            "The first element moves. The second stays put.",
            "There is a video walkthrough on the button's help link (F1 while "
            "hovering).",
        ],
    },
    {
        "chapter": "Using it",
        "title": "Click near the connector you mean",
        "target": None,
        "body": (
            "This is the whole trick, and it is easy to miss: the tool picks the "
            "connector CLOSEST TO WHERE YOU CLICK. So click near the end you want "
            "joined, not in the middle of the element."
        ),
        "bullets": [
            "First prompt: 'Pick element to move' - click close to the connector that "
            "should do the joining.",
            "Second prompt: 'Pick element to be connected to' - again, click close to "
            "the target connector.",
            "Clicking mid-span gives you whichever connector happens to be nearest, "
            "which is rarely what you wanted.",
        ],
    },
    {
        "chapter": "When it complains",
        "title": "The four messages you will see",
        "target": None,
        "body": (
            "The errors are specific, and each one tells you exactly what to change."
        ),
        "bullets": [
            "'You've selected the same object twice' - you picked one element for "
            "both prompts.",
            "'You picked 2 connectors of different domain' - piping to ducting, or "
            "similar. Connectors have to be the same domain.",
            "'One of the objects have no unused connector' - every connector on that "
            "element is already joined. Disconnect something first.",
            "'You chose an invalid object' - the thing you picked has no connectors "
            "the tool can use.",
        ],
    },
    {
        "chapter": "Working with it",
        "title": "Where it fits",
        "target": None,
        "body": (
            "Use ConnectTo for the one-off joins, and Disconnect as its opposite "
            "number when you need to take a network apart."
        ),
        "bullets": [
            "If a join keeps failing on 'no unused connector', run Disconnect on the "
            "offending element first, then ConnectTo.",
            "It moves geometry, so check the result before moving on - undo is one "
            "step.",
        ],
    },
]
