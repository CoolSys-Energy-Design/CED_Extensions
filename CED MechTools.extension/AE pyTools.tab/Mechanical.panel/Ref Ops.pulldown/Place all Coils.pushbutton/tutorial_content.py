# -*- coding: utf-8 -*-
"""Narration for the Place all Coils guided tour.

Place all Coils reads an Excel list of descriptions, coil counts and model numbers,
matches the descriptions to Spaces or Rooms, and places coil families in them.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Place all Coils - Show Me Around"
TOUR_KEY = "place_all_coils"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "Excel in, coils in spaces out",
        "target": None,
        "body": (
            "This tool turns a coil schedule into placed families. It reads "
            "Description, Coil Count and Model # from Excel, finds the Space or Room "
            "whose name matches the description, and places that many coils in it."
        ),
        "bullets": [
            "Requires Excel to be installed - it says 'Excel is not registered on "
            "this machine' otherwise.",
            "Places Mechanical Equipment family types; there must be some in the "
            "project.",
            "Coils are offset 2 feet vertically from the space's placement point.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "Where do the spaces come from?",
        "target": None,
        "body": (
            "First it asks for the source of the Spaces and Rooms. On a typical job "
            "the architectural rooms live in a link, so you point it at the link "
            "rather than the host."
        ),
        "bullets": [
            "Host model (this project) - use the Spaces or Rooms in the model you are "
            "in.",
            "Or pick a loaded Revit link and use its rooms.",
            "Spaces with no location are flagged '[no location]' so you know they "
            "cannot receive anything.",
            "'No Spaces or Rooms found in the host model' means pick the link "
            "instead.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "Pick the placement level",
        "target": None,
        "body": (
            "Coils have to land on a level. It offers a level picker and pre-selects "
            "a 'Level 1'-ish level by name, falling back to the lowest elevation if "
            "no name matches."
        ),
        "bullets": [
            "Check the pre-selection rather than accepting it - the guess is name- "
            "based.",
            "Cancelling here cancels the run.",
        ],
    },
    {
        "chapter": "Setting up",
        "title": "Point at the Excel file",
        "target": None,
        "body": (
            "Then choose the coil placement workbook. The columns it cares about are "
            "Description, coil count and Model #."
        ),
        "bullets": [
            "Description is matched against space and room names.",
            "Model # drives which family type is chosen.",
            "It reports how many rows it read, and from which sheet.",
        ],
    },
    {
        "chapter": "How it matches",
        "title": "Description to space, Model # to type",
        "target": None,
        "body": (
            "Two independent matches have to succeed for a row to place anything, and "
            "both are fuzzy rather than exact."
        ),
        "bullets": [
            "Description is matched to the Space or Room name.",
            "Model # is matched to the closest Mechanical Equipment family type "
            "available.",
            "Rows can be filtered by manufacturer, and rows with a blank or "
            "mismatched manufacturer are skipped and counted.",
            "Set a row's space to <Do not place in any space> to exclude it "
            "deliberately.",
        ],
    },
    {
        "chapter": "When nothing happens",
        "title": "Read the diagnostic report",
        "target": None,
        "body": (
            "If it places nothing it does not just shrug - it prints a full "
            "diagnostic to the output panel telling you which stage lost your rows. "
            "That report is the tool's best feature."
        ),
        "bullets": [
            "It lists the Excel file and sheet, and how many rows were read.",
            "How many spaces the chosen source produced, and how many were '[no "
            "location]'.",
            "How many Mechanical Equipment types exist versus how many matched.",
            "The placement level, and how many rows passed every filter.",
            "Per-reason skip counts - manufacturer mismatch, blank manufacturer, no "
            "candidate symbols, and so on.",
            "Work down that list; the row counts show exactly where the data broke.",
        ],
    },
    {
        "chapter": "Afterwards",
        "title": "Then space them out",
        "target": None,
        "body": (
            "Place all Coils gets the right number of the right coils into the right "
            "rooms. It does not lay them out nicely."
        ),
        "bullets": [
            "Follow it with Space Coils to distribute them along a wall or on a grid.",
            "That two-step split is deliberate: placement is data-driven, layout is a "
            "judgement call.",
        ],
    },
]
