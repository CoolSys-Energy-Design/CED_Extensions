# -*- coding: utf-8 -*-
"""Narration for the Print Pipe Data guided tour.

Print Pipe Data totals pipe quantities per system ID and exports the result to
Excel. It is the takeoff at the end of the refrigeration piping workflow.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Print Pipe Data - Show Me Around"
TOUR_KEY = "print_pipe_data"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "Totals per system ID, straight to Excel",
        "target": None,
        "body": (
            "This is the takeoff step. It gathers pipe, groups it by system ID, "
            "totals the lengths and capacities, and writes a spreadsheet."
        ),
        "bullets": [
            "Read-only on the model - it measures, it does not modify.",
            "System ID comes from the Identity Mark parameter on the pipes.",
            "Run it after Name Piping Systems has populated those Identity Marks.",
        ],
    },
    {
        "chapter": "Using it",
        "title": "Select the system types",
        "target": None,
        "body": (
            "It asks which pipe system types to include, built from the system types "
            "it actually found on the pipes in the model. That is your scope control."
        ),
        "bullets": [
            "The button tooltip still says 'select worksets' - that is out of date; "
            "it asks for system types.",
            "'No pipes were found in this model' or 'No system types were found on "
            "the pipe elements' means there is nothing to total.",
            "Cancelling the picker ends the run with 'No system types were selected.'",
        ],
    },
    {
        "chapter": "What you get",
        "title": "One row per system ID",
        "target": None,
        "body": (
            "The output is a Pipe Totals sheet keyed on system ID, with the "
            "quantities summarised per PipeSegment."
        ),
        "bullets": [
            "Element IDs - what went into each total, so you can trace a number back.",
            "Vertical Length Total and Horizontal Length Total - kept separate, "
            "because they cost differently.",
            "Evaporation Capacity Total.",
            "Pipe with no Identity Mark is collected into an '<Unassigned ID>' bucket "
            "rather than dropped - check that bucket, it is usually a labelling gap.",
            "System types the pipes do not identify show as '<Unknown System Type>'.",
        ],
    },
    {
        "chapter": "What you get",
        "title": "Saving the file",
        "target": None,
        "body": (
            "It prompts for a save location and writes Excel. If Excel is not "
            "available it falls back to CSV rather than failing."
        ),
        "bullets": [
            "'Excel export failed. Falling back to CSV' tells you which format you "
            "got.",
            "The completion message lists system types, pipe count, direct identity "
            "values found, system IDs exported, assigned IDs, whether the unassigned "
            "bucket is present, and the file path.",
            "Read the assigned-versus-unassigned counts before you send the file on - "
            "that is your data-quality check.",
        ],
    },
]
