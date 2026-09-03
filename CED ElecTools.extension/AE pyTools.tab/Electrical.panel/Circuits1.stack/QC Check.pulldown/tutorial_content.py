# -*- coding: utf-8 -*-
"""Narration for the QC Check dropdown guided tour.

What each of the four QC Check tools does, and which one to reach for when you are
reviewing someone else's electrical model.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "QC Check - What's In Here"
TOUR_KEY = "qc_check_dropdown"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "Overview",
        "title": "Four ways to audit a model",
        "target": None,
        "body": (
            "QC Check holds two visual checks that colour the model, and two report "
            "generators that produce something you can read or hand over."
        ),
        "bullets": [
            "The colour tools are toggles - run them again to turn the override off.",
            "The report tools are read-only; they never change the model.",
        ],
    },
    {
        "chapter": "Visual checks",
        "title": "Color Circuits Check",
        "target": None,
        "body": (
            "A toggle that colours fixtures red or green depending on whether they "
            "are circuited. It is the fastest way to spot devices nobody wired."
        ),
        "bullets": [
            "First run creates the red/green view filters and enables them.",
            "Run it again to disable them, and again to re-enable - it cycles.",
            "It uses real view filters, not temporary hide/isolate, so it does not "
            "disturb your view state or view templates.",
        ],
    },
    {
        "chapter": "Visual checks",
        "title": "Color Circuits by Panel",
        "target": None,
        "body": (
            "Colours the model by which panel feeds each element, and builds a "
            "matching legend so the drawing explains itself."
        ),
        "bullets": [
            "Creates or updates one multi-category filter per panel.",
            "Included panels get a unique colour; everything else goes grey and "
            "halftone.",
            "Also creates or updates a 'PanelChecker Legend' drafting view with "
            "swatches.",
            "It never creates or edits view templates.",
        ],
    },
    {
        "chapter": "Reports",
        "title": "Electrical System Check",
        "target": None,
        "body": (
            "Sweeps equipment, devices and circuits for the mistakes that matter - "
            "panel capacity, device loads, circuit connections - and writes a report "
            "you can work down."
        ),
        "bullets": [
            "This is the one to run before issuing a set.",
            "Read-only: it tells you what is wrong, it does not fix anything.",
        ],
    },
    {
        "chapter": "Reports",
        "title": "Panel Report",
        "target": None,
        "body": (
            "Generates a detailed report of every panel and its breakers, formatted "
            "similarly to a manufacturer submittal."
        ),
        "bullets": [
            "Intended to make submittal review easier - compare the report against "
            "what the manufacturer sent back.",
        ],
    },
]
