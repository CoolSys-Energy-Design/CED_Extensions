# -*- coding: utf-8 -*-
"""Narration for the Tag by Example guided tour.

Tag by Example copies an existing tag's placement onto other instances. The tour
covers picking the example, choosing targets, the four options that decide how
faithful the copy is, and what happens to tags that already exist.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Tag by Example - Show Me Around"
TOUR_KEY = "tag_by_example"

HOST = {
    "kind": "window",
    "title": "Tag by Example",
}

TARGET_LABELS = {
    "CopyLeaderToggle": "the copy leader switch",
    "CreateButton": "the Create button",
    "EditTargetsButton": "the Edit Targets button",
    "ErrorBannerText": "the error banner",
    "ExampleTypeText": "the example tag type line",
    "HostCategoryText": "the host category line",
    "HostFamilyText": "the host family line",
    "HostTypeText": "the host type line",
    "IncludeNestedToggle": "the include nested switch",
    "LeaderText": "the leader line",
    "NewTargetsButton": "the New Targets button",
    "OwnerViewText": "the owner view line",
    "PickExampleButton": "the Pick Example button",
    "PreserveRotationToggle": "the preserve rotation switch",
    "PreviewSelectionButton": "the Preview Selection button",
    "ReferenceCountText": "the reference count",
    "SkipMatchingRadio": "the existing-tag options",
    "StatusText": "the status line",
    "TargetCountText": "the target count",
    "TargetModeCombo": "the target mode dropdown",
    "TargetViewText": "the target view line",
    "UseModelOrientationToggle": "the model orientation switch",
}

STEPS = [
    {
        "chapter": "Getting oriented",
        "title": "Tag one thing well, then repeat it",
        "target": None,
        "body": (
            "Tag by Example takes a tag you have already placed and positioned "
            "nicely, and reproduces that placement on every other comparable instance "
            "- keeping the offset relative to the host rather than in absolute space."
        ),
        "bullets": [
            "The point is host-relative placement: each new tag sits where the "
            "example sits relative to its own host.",
            "Needs an active project document and an active view.",
            "Modeless - it stays open while you pick in the model.",
            "If you had tags selected when you launched it, they are picked up as the "
            "starting example.",
        ],
    },
    {
        "chapter": "The example",
        "title": "Pick the tag to copy",
        "target": "PickExampleButton",
        "body": (
            "Click Pick Example, then select the tag in the model whose placement you "
            "want reproduced. Everything the tool does is derived from this one tag."
        ),
        "bullets": [
            "Position the example tag exactly how you want the others before picking "
            "it.",
            "You can re-pick at any time to change the pattern.",
        ],
    },
    {
        "chapter": "The example",
        "title": "Check what it actually read",
        "target": "HostCategoryText",
        "body": (
            "The panel then reports what it understood: the tag type, and the host's "
            "category, family and type. Read this before you create anything - it is "
            "your confirmation the tool is generalising from the right thing."
        ),
        "bullets": [
            "Owner view tells you which view the example tag lives in.",
            "The leader line reports whether the example has a leader and what shape.",
            "Reference count matters for multi-reference tags.",
        ],
    },
    {
        "chapter": "The targets",
        "title": "How wide to cast the net",
        "target": "TargetModeCombo",
        "body": (
            "Target mode decides what counts as 'comparable'. Narrow it to the same "
            "type, or widen it, depending on whether the family's other types want "
            "the same tag placement."
        ),
        "bullets": [
            "Same type is the safe default and what the tool starts on.",
            "Widen it when a family's types share geometry and should tag "
            "identically.",
        ],
    },
    {
        "chapter": "The targets",
        "title": "Choose the instances",
        "target": "NewTargetsButton",
        "body": (
            "New Targets starts a fresh pick; Edit Targets adjusts the set you "
            "already have. The count tells you how many instances will get a tag."
        ),
        "bullets": [
            "Preview Selection highlights the current target set in the model - use "
            "it before Create, every time.",
            "Include nested controls whether instances nested inside families are "
            "eligible.",
        ],
    },
    {
        "chapter": "Faithfulness",
        "title": "The three switches that decide how exact the copy is",
        "target": "PreserveRotationToggle",
        "body": (
            "These control how literally the example's placement is reproduced. They "
            "are the difference between tags that look deliberate and tags you have "
            "to nudge."
        ),
        "bullets": [
            "Preserve rotation - keep the example tag's rotation on every copy.",
            "Use model orientation - place relative to the host's own orientation "
            "rather than the view's, which is what you want for rotated equipment.",
            "Copy leader - reproduce the example's leader instead of placing bare "
            "tags.",
        ],
    },
    {
        "chapter": "Existing tags",
        "title": "What happens where a tag already is",
        "target": "SkipMatchingRadio",
        "body": (
            "Most of the time you are running this over a partly tagged view. These "
            "three options decide what happens to tags that are already there."
        ),
        "bullets": [
            "Skip matching - leave existing equivalent tags alone. The safe default.",
            "Replace matching - delete equivalent tags and place fresh ones.",
            "Replace all - clear any tag on the target and place yours.",
            "Replace all is the one to be careful with - it discards other people's "
            "tagging decisions.",
        ],
    },
    {
        "chapter": "Running it",
        "title": "Confirm the view, then Create",
        "target": "TargetViewText",
        "body": (
            "Tags land in the view named here. Create does the work and the status "
            "line reports what was placed and skipped."
        ),
        "bullets": [
            "If a red banner appears, read it and dismiss it - it names the specific "
            "problem.",
            "Nothing is written until Create; closing the window loses only the "
            "setup.",
            "Your option choices are remembered between sessions in your pyRevit "
            "config.",
        ],
    },
]
