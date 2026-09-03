# -*- coding: utf-8 -*-
"""Narration for the Disconnect guided tour.

Disconnect separates selected pipe or duct segments from the rest of the network,
and cleans up the system membership that separation leaves behind.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Disconnect - Show Me Around"
TOUR_KEY = "disconnect"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "Cut a selection out of its network",
        "target": None,
        "body": (
            "Disconnect breaks the joints between the segments you selected and "
            "everything they were attached to. It is the opposite number to "
            "ConnectTo, and it does more than just unhook geometry."
        ),
        "bullets": [
            "Select the pipe or duct segments first; with nothing selected it just "
            "says 'Nothing selected.'",
            "It works through ConnectorSet removal, not by deleting and redrawing.",
        ],
    },
    {
        "chapter": "How it is careful",
        "title": "It disconnects from the right side",
        "target": None,
        "body": (
            "When two elements are joined there are two connectors involved, and "
            "which one you release from matters. The tool prefers to disconnect from "
            "the MEPCurve side so endpoint and system ownership stay intact."
        ),
        "bullets": [
            "That is why a disconnect here behaves better than doing it by hand from "
            "the fitting side.",
            "It reports each cut as 'Disconnected X from Y' so you can follow along.",
        ],
    },
    {
        "chapter": "How it is careful",
        "title": "It cleans up orphaned equipment",
        "target": None,
        "body": (
            "Disconnecting often leaves an endpoint element - a piece of equipment or "
            "a fixture - belonging to a system it is no longer physically part of. "
            "The tool removes those stale memberships."
        ),
        "bullets": [
            "You will see 'Removed orphaned <type> element <id> from system <name>'.",
            "This is the step that stops phantom system membership showing up in "
            "schedules later.",
        ],
    },
    {
        "chapter": "How it is careful",
        "title": "It re-divides the systems",
        "target": None,
        "body": (
            "After the cut, the original system may now be two separate networks. The "
            "tool asks Revit to divide it so each piece becomes its own system."
        ),
        "bullets": [
            "'Divided piping system X' means the split took.",
            "'System X is still single network' means the cut did not actually "
            "separate anything - the network is still continuous by some other route. "
            "That is usually the interesting finding.",
        ],
    },
    {
        "chapter": "Working with it",
        "title": "Read the log, then verify",
        "target": None,
        "body": (
            "Every action is reported. If something did not work, the message names "
            "the element and the reason rather than failing silently."
        ),
        "bullets": [
            "'No connector manager found for X' - that element has no connectors to "
            "work with.",
            "'Failed disconnect' or 'Failed dividing system' lines carry the "
            "underlying error.",
            "Set Piping System uses this same disconnect-and-divide machinery "
            "internally, so the messages will look familiar.",
        ],
    },
]
