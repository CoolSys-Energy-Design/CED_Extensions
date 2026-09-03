# -*- coding: utf-8 -*-
"""Narration for the Space Coils guided tour.

Space Coils takes coils you have already placed and distributes them evenly - either
along a wall direction or on a centred grid.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Space Coils - Show Me Around"
TOUR_KEY = "space_coils"

# Narration-only: no persistent window of its own to spotlight.
HOST = None

TARGET_LABELS = {}

STEPS = [
    {
        "chapter": "What it does",
        "title": "Layout, not placement",
        "target": None,
        "body": (
            "Space Coils is the second half of the coil workflow. The coils already "
            "exist; this tool arranges them evenly using the walls or the space "
            "around them."
        ),
        "bullets": [
            "Select the coil family instances before you run it - it insists on that.",
            "It works out the bounding walls and the Space or Room from your "
            "selection.",
            "It measures from visible geometry when a family's raw bounding box is "
            "misleadingly large, and tells you when it does.",
        ],
    },
    {
        "chapter": "Choosing a mode",
        "title": "Wall or centre",
        "target": None,
        "body": (
            "The first question is how the coils should be arranged. The two modes "
            "suit different rooms."
        ),
        "bullets": [
            "Wall distribution - line them along one wall direction, evenly, with "
            "equal gaps to the walls at each end.",
            "Center distribution - arrange them on a grid centred in the space.",
            "Pick wall distribution for a run of cases along a wall; centre for an "
            "open room.",
        ],
    },
    {
        "chapter": "Wall distribution",
        "title": "Which wall, and how far off it",
        "target": None,
        "body": (
            "For wall distribution it asks which wall direction to align to, then how "
            "far off the wall to sit."
        ),
        "bullets": [
            "Coils face opposite the wall direction, so they point into the room.",
            "The offset is in inches and 15 inches is the normal value.",
            "It must be zero or positive, and it must be a number - it will tell you "
            "off otherwise.",
            "Gaps to the left and right walls come out equal.",
        ],
    },
    {
        "chapter": "Center distribution",
        "title": "Set the grid, then the template",
        "target": None,
        "body": (
            "Centre distribution asks for a grid size as rows by columns, and then "
            "whether the coils in that grid are all the same size."
        ),
        "bullets": [
            "'Selected N coil(s), but grid only has M slots' means your grid is too "
            "small for the selection.",
            "If the grid mixes coil sizes, it asks you to arrange the selected coils "
            "into the layout you want first, then click Continue - that tells the "
            "tool which size belongs in which grid position.",
            "That template step is skipped when every coil is the same size.",
        ],
    },
    {
        "chapter": "When it refuses",
        "title": "Not enough room",
        "target": None,
        "body": (
            "Most failures are geometric, and the message says which constraint it "
            "could not satisfy."
        ),
        "bullets": [
            "'Not enough space between the walls to place the coils' - the run does "
            "not fit at the offset you asked for.",
            "'Not enough space to fit the grid with equal spacing' - same problem in "
            "grid mode.",
            "'Could not find bounding walls or a Space/Room for the selection' - "
            "select coils that sit inside a Space or Room, and make sure the nearby "
            "walls are visible in the view.",
            "That last one is usually a view visibility problem, not a model problem.",
        ],
    },
]
