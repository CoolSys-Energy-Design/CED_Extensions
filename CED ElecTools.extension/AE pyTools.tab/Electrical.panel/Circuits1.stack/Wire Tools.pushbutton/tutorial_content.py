# -*- coding: utf-8 -*-
"""Narration for the Wire Tools guided tour.

Wire Tools draws and tags wire for circuits that already exist. The tour walks the
window top to bottom: pick a scheme, build a selection, set the geometry, then Run.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Wire Tools - Show Me Around"
TOUR_KEY = "wire_tools"

HOST = {
    "kind": "window",
    "title": "Wire Tools",
}

TARGET_LABELS = {
    "BendOffsetText": "the bend offset box",
    "BranchArcRadio": "the branch bend style options",
    "DeviceCountText": "the device count",
    "ExistingTagBehaviorCombo": "the existing-tag behaviour dropdown",
    "HomerunLengthText": "the homerun length box",
    "HomerunPanelDirectionRadio": "the homerun direction options",
    "HomerunTaggingToggle": "the homerun tagging switch",
    "InterconnectScopePanel": "the interconnect scope options",
    "NodeGroup": "the node selection group",
    "RedrawExistingToggle": "the redraw-existing switch",
    "RunButton": "the Run button",
    "SchemeCombo": "the wiring scheme dropdown",
    "SchemeInstructionsText": "the scheme description line",
    "SelectDevicesButton": "the Select Devices button",
    "SelectHomerunsButton": "the Select Homeruns button",
    "SelectNodeButton": "the Select Node button",
    "SkipSingleDeviceToggle": "the skip-single-device switch",
    "StatusText": "the status line",
    "SystemTypeCombo": "the system type dropdown",
    "TagHomerunsButton": "the Tag Homeruns button",
    "TagTypeCombo": "the tag type dropdown",
    "TargetViewText": "the target view line",
    "UseCurrentButton": "the Use Current view button",
    "WireTypeCombo": "the wire type dropdown",
}

STEPS = [
    {
        "chapter": "Getting oriented",
        "title": "Wire follows circuits, not the other way round",
        "target": None,
        "body": (
            "Wire Tools draws wire between things that are already circuited. It does "
            "not create circuits - do that first in Circuit Manager or Circuit Tools, "
            "then come here to draw and tag the wire."
        ),
        "bullets": [
            "The window is modeless: leave it open and keep working in the model.",
            "It needs an active project document; it will tell you if there isn't "
            "one.",
            "Work down the window - scheme, selection, geometry, then Run.",
        ],
    },
    {
        "chapter": "Choosing a scheme",
        "title": "Four wiring schemes",
        "target": "SchemeCombo",
        "body": (
            "The scheme is the single most important choice in this window - it "
            "decides what gets connected to what. Everything below the dropdown "
            "reshapes itself to suit the scheme you pick."
        ),
        "bullets": [
            "Wire by Circuit - connect selected circuits or devices, plus every other "
            "device sharing their underlying circuits.",
            "Interconnect - connect the selected devices in one continuous sequence, "
            "even across several circuits.",
            "Individual Homeruns - one homerun from each selected device, off its "
            "primary connector.",
            "Wire to Node - connect each selected device straight to one compatible "
            "electrical node.",
        ],
    },
    {
        "chapter": "Choosing a scheme",
        "title": "Read the description under the dropdown",
        "target": "SchemeInstructionsText",
        "body": (
            "This line restates in plain language what the current scheme will do. If "
            "it does not describe what you want, change the scheme rather than "
            "fighting the options below it."
        ),
        "bullets": [],
    },
    {
        "chapter": "Building a selection",
        "title": "Select the devices",
        "target": "SelectDevicesButton",
        "body": (
            "Devices are the main input for every scheme. Click Select Devices, pick "
            "in the model, and the count next to it confirms what the tool is "
            "holding."
        ),
        "bullets": [
            "The status text beside the count tells you whether the selection is "
            "valid for the current scheme.",
            "Clear beside it empties the selection without closing the window.",
        ],
    },
    {
        "chapter": "Building a selection",
        "title": "Homeruns and nodes",
        "target": "SelectNodeButton",
        "body": (
            "Some schemes need a second input. Wire to Node needs the node; homerun "
            "work needs the existing homeruns. These controls only appear for the "
            "schemes that use them."
        ),
        "bullets": [
            "Select Node - pick the one compatible electrical node to wire into.",
            "Select Homeruns - pick existing homerun wires to work with or retag.",
            "Interconnect adds a scope option so you can control how far the chain "
            "runs.",
        ],
    },
    {
        "chapter": "Wire and system",
        "title": "Wire type and system type",
        "target": "WireTypeCombo",
        "body": (
            "Pick the wire type the new wire is created with, and where relevant the "
            "system type it belongs to. These are the same types Revit offers - the "
            "tool just saves you the trip to the properties palette."
        ),
        "bullets": [
            "Set these before you Run; changing them afterwards means redrawing.",
        ],
    },
    {
        "chapter": "Geometry",
        "title": "Bend style and offset",
        "target": "BranchArcRadio",
        "body": (
            "Branches and homeruns each get their own bend style - arc or chamfer - "
            "and a bend offset controls how far the bend sits from the device."
        ),
        "bullets": [
            "The small info button beside these opens a diagram of what each option "
            "looks like.",
            "Offset is the usual knob for making wire read cleanly at plan scale.",
        ],
    },
    {
        "chapter": "Geometry",
        "title": "Which way a homerun leaves",
        "target": "HomerunPanelDirectionRadio",
        "body": (
            "A homerun has to point somewhere. Choose whether it leaves in the "
            "direction the device faces, or aims toward the panel, then set its "
            "length."
        ),
        "bullets": [
            "Device facing - follows the device's own orientation.",
            "Panel direction - aims the tail toward the supplying panel.",
            "Homerun length controls how long that tail is drawn.",
            "The info button here also opens a diagram.",
        ],
    },
    {
        "chapter": "Tagging",
        "title": "Tag homeruns in the same pass",
        "target": "HomerunTaggingToggle",
        "body": (
            "Turn this on and Wire Tools tags the homeruns it draws, so you are not "
            "making a second pass with the tag tool."
        ),
        "bullets": [
            "Tag type picks the annotation family used.",
            "Add leaders controls whether tags get a leader.",
            "The existing-tag behaviour dropdown decides what happens where a tag is "
            "already present - leave it, or replace it.",
            "Tag Homeruns runs the tagging on its own, without redrawing wire.",
        ],
    },
    {
        "chapter": "Running it",
        "title": "The two switches that catch people out",
        "target": "RedrawExistingToggle",
        "body": (
            "Before you Run, check these two. They are the usual reason output does "
            "not match expectations."
        ),
        "bullets": [
            "Redraw existing - whether wire that is already drawn gets replaced or "
            "left alone.",
            "Skip single device - whether a circuit with only one device gets wire at "
            "all.",
        ],
    },
    {
        "chapter": "Running it",
        "title": "Confirm the target view",
        "target": "TargetViewText",
        "body": (
            "Wire is drawn into a specific view. This line tells you which one. Use "
            "Current re-points it at the view you are looking at now."
        ),
        "bullets": [
            "If wire appears 'nowhere', it usually went into the view named here.",
            "Set this deliberately when you have several plan views open.",
        ],
    },
    {
        "chapter": "Running it",
        "title": "Run, then read the status line",
        "target": "RunButton",
        "body": (
            "Run does the work in one transaction. The status line underneath reports "
            "what was drawn, tagged and skipped - read it rather than assuming."
        ),
        "bullets": [
            "If the window greys out with a message, the active context is wrong - "
            "usually the wrong kind of view.",
            "Nothing is written until you press Run; closing the window costs you "
            "nothing but the selection.",
        ],
    },
]
