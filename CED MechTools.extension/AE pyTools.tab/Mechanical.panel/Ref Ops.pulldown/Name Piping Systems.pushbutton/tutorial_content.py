# -*- coding: utf-8 -*-
"""Narration for the Name Piping Systems guided tour.

Name Piping Systems labels refrigeration pipe by walking outward from a root pipe
you pick, following the connected network and applying branch rules.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Name Piping Systems - Show Me Around"
TOUR_KEY = "name_piping_systems"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "Label a pipe tree from its root",
        "target": None,
        "body": (
            "You pick the pipe at the rack where a system starts, and the tool "
            "traverses the connected pipe network outward from there, labelling as it "
            "goes. Branch rules decide what each downstream run is called."
        ),
        "bullets": [
            "It can write Identity Mark, place tags, place text notes, or a "
            "combination.",
            "The active view cannot be a template - open a working view first.",
            "It understands connector nodes, accessory branches and ball valves as it "
            "traverses.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "The identity mode",
        "target": None,
        "body": (
            "First decide how it should treat Identity Mark on pipes that already "
            "have one. This is the setting that determines whether the run is "
            "additive or destructive."
        ),
        "bullets": [
            "Do Not Update Identity Marks - leave the parameter alone entirely.",
            "Fill Blank Identity Marks Only - additive and safe; only touches empty "
            "ones.",
            "Rewrite All Identity Marks - overwrites whatever is there.",
            "Use Fill Blank Only when extending an existing labelled system.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "The tagging mode",
        "target": None,
        "body": (
            "Then decide the same question for tags: add only where missing, or clear "
            "and redo."
        ),
        "bullets": [
            "Do Not Add Tags - label without annotating.",
            "Tag Untagged Pipes Only - additive.",
            "Remove Matching Tags and Retag - clears the tool's own tags and places "
            "fresh ones.",
            "You also pick the pipe tag family type; 'No pipe tag types found in this "
            "project' means load one first.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "Number, suffix and placement",
        "target": None,
        "body": (
            "The remaining inputs build the actual label text and decide where it "
            "goes."
        ),
        "bullets": [
            "Pipe Number - must be greater than zero.",
            "Suffix Letter - exactly one letter, A to Z.",
            "Label Placement - toggles for where labels sit, including System First "
            "ordering. You must select at least one placement option.",
            "It validates all of these before doing anything, with a specific message "
            "for each.",
        ],
    },
    {
        "chapter": "Running it",
        "title": "Pick the root pipe",
        "target": None,
        "body": (
            "Then it asks you to select the root pipe for the tree. Everything "
            "follows from that pick, so pick the pipe at the rack start."
        ),
        "bullets": [
            "Cancelling gives 'Operation canceled.' and writes nothing.",
            "'No root pipe selected' means the pick did not land on a pipe.",
            "'No pipes were labeled from the selected root pipe' means the traversal "
            "found nothing connected - check the root really is connected to the "
            "network.",
        ],
    },
    {
        "chapter": "Afterwards",
        "title": "Check the connection list",
        "target": None,
        "body": (
            "The run prints the pipe connection list of everything it labelled. That "
            "is how you confirm the traversal followed the route you expected rather "
            "than wandering into a neighbouring system."
        ),
        "bullets": [
            "It reports separately for Text Notes, Tags and Identity Mark.",
            "'All candidate labels were removed by FF cleanup' means the cleanup "
            "rules rejected everything - worth investigating rather than re-running.",
            "Follow this with Print Pipe Data to total up what you have labelled.",
        ],
    },
]
