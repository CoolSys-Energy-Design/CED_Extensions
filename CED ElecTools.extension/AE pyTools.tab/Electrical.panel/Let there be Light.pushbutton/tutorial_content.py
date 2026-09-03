# -*- coding: utf-8 -*-
"""Narration for the Let There Be Light guided tour.

Let There Be Light places lighting fixtures from a CAD CSV export or from a linked
Revit model. The tour follows the real order of work: choose a source, load it, map
fixture types and levels, validate, then place.

Pure data - see UIClasses/tutorial/__init__.py for the content contract.
Reword freely; the tour re-reads this file when it changes on disk.
"""

TOUR_TITLE = "Let There Be Light - Show Me Around"
TOUR_KEY = "let_there_be_light"

HOST = {
    "kind": "window",
    "title": "Let There Be Light",
}

TARGET_LABELS = {
    "ApplyBulkFixtureButton": "the Apply to selected rows button",
    "AutoMapButton": "the Auto-Match button",
    "BrowseCsvButton": "the Browse button",
    "BulkFixtureCombo": "the bulk mapping dropdown",
    "CsvLevelCombo": "the placement level dropdown",
    "CsvPathTextBox": "the CSV path box",
    "DivisorTextBox": "the unit adjustment box",
    "FixtureGrid": "the fixture mapping grid",
    "LevelGrid": "the level mapping grid",
    "LevelMappingSection": "the level mapping section",
    "LinkCombo": "the Revit link dropdown",
    "LinkSummaryText": "the link summary line",
    "LoadCsvButton": "the Load CSV button",
    "LoadLinkButton": "the Load Link Fixtures button",
    "LoadedCountText": "the loaded-fixtures count",
    "ModeWarningText": "the mode warning line",
    "PlaceButton": "the Place button",
    "SkipDuplicatesCheck": "the skip duplicates switch",
    "SourceModeCombo": "the data source dropdown",
    "ValidationText": "the validation line",
}

STEPS = [
    {
        "chapter": "Getting oriented",
        "title": "Two sources, one placement engine",
        "target": None,
        "body": (
            "This tool places lighting fixtures into the host model from somebody "
            "else's data - either a CSV exported from CAD, or the lighting fixtures "
            "inside a linked Revit model. The mapping and placement half is the same "
            "either way."
        ),
        "bullets": [
            "CSV mode - map CAD block names to host fixture types, then place on one "
            "level you choose.",
            "Link mode - collect the linked fixtures, map family/types AND levels, "
            "then place into the host.",
            "Nothing is placed until you press Place at the end.",
        ],
    },
    {
        "chapter": "Choose a source",
        "title": "Pick your source first",
        "target": "SourceModeCombo",
        "body": (
            "Set the source mode before anything else. Switching it later clears the "
            "loaded fixtures and every mapping you have made, and the tool will warn "
            "you before it does."
        ),
        "bullets": [
            "The warning is real - switching mid-way means redoing the mapping.",
            "The line under the header restates what the current mode will do.",
        ],
    },
    {
        "chapter": "Choose a source",
        "title": "CSV mode: path and units",
        "target": "CsvPathTextBox",
        "body": (
            "Browse to the CAD export, then set the unit adjustment factor. Getting "
            "the factor wrong is the single most common cause of fixtures landing in "
            "the wrong place."
        ),
        "bullets": [
            "Enter 1 if the CSV coordinates are already in feet.",
            "Enter 12 if they are in inches.",
            "Then Load CSV reads the file and reports how many fixtures and source "
            "types it found.",
        ],
    },
    {
        "chapter": "Choose a source",
        "title": "Link mode: pick the link and load",
        "target": "LinkCombo",
        "body": (
            "Choose a loaded Revit link, then Load Link Fixtures to collect its "
            "lighting fixtures."
        ),
        "bullets": [
            "The link has to be loaded - an unloaded or unavailable link is rejected "
            "with a message.",
            "The summary line reports what was collected.",
            "Link mode reads the source type from a parameter you select, rather than "
            "deriving it from text.",
        ],
    },
    {
        "chapter": "Mapping fixtures",
        "title": "Map source types to host types",
        "target": "FixtureGrid",
        "body": (
            "This grid is the heart of the tool: one row per source fixture type, and "
            "the host family and type it should become. Unmapped rows are skipped at "
            "placement, not guessed at."
        ),
        "bullets": [
            "Set a row to <Skip Mapping> to deliberately exclude that source type.",
            "The loaded count above tells you how many fixtures and source types are "
            "in play.",
        ],
    },
    {
        "chapter": "Mapping fixtures",
        "title": "Let it guess first",
        "target": "AutoMapButton",
        "body": (
            "Auto-Match opens a matching window that pairs source values with host "
            "types by parameter, so you are correcting a guess instead of filling in "
            "a blank grid."
        ),
        "bullets": [
            "In CSV mode the source value is derived from the Text Before and Text "
            "After boxes - useful when block names carry a prefix or suffix.",
            "In link mode it matches on a source type parameter you choose.",
            "Common target parameters are Identity Type Mark, Fixture Type_CEDT and "
            "CED-E-FIXTURE TYPE.",
            "Review the result in the grid; auto-match is a head start, not an "
            "answer.",
        ],
    },
    {
        "chapter": "Mapping fixtures",
        "title": "Fix many rows at once",
        "target": "BulkFixtureCombo",
        "body": (
            "Select several rows, pick a target family/type in the bulk dropdown, and "
            "apply it to all of them. Clear does the reverse."
        ),
        "bullets": [
            "It tells you off if you apply with no rows selected, or with no target "
            "chosen.",
            "Fastest way to handle a source that exported one fixture as a dozen "
            "block variants.",
        ],
    },
    {
        "chapter": "Mapping levels",
        "title": "CSV: one level. Link: a level map",
        "target": "LevelMappingSection",
        "body": (
            "Levels are where the two modes differ most. CSV data has no level, so "
            "you pick one. Linked fixtures each carry a level, so you map them."
        ),
        "bullets": [
            "CSV mode - choose the single placement level from the dropdown.",
            "Link mode - the level grid maps each link level to a host level.",
            "Linked fixtures with no level assigned show as <No Level Assigned in "
            "Link> and need a target like anything else.",
            "Unmapped levels are skipped, and reported in the result.",
        ],
    },
    {
        "chapter": "Placing",
        "title": "Validate before you commit",
        "target": "ValidationText",
        "body": (
            "The validation line is the tool telling you whether it is ready. Skip "
            "duplicates protects you from placing fixtures on top of ones already "
            "there - useful on a second pass."
        ),
        "bullets": [
            "Leave skip duplicates on unless you specifically want doubles.",
            "If validation complains, it names what is unmapped or invalid.",
        ],
    },
    {
        "chapter": "Placing",
        "title": "Place, then read the tally",
        "target": "PlaceButton",
        "body": (
            "Place writes the fixtures. The result is an itemised tally, and it is "
            "worth reading every time - the skip counts tell you what your mapping "
            "missed."
        ),
        "bullets": [
            "The tally reports placed, plus skipped for unmapped fixture, unmapped "
            "level, invalid point, duplicates and errors.",
            "A big 'unmapped fixture' number means go back to the grid, not that the "
            "tool failed.",
            "There is also a Normalize Lighting Fixture Elevations pass for cleaning "
            "up heights after placement.",
        ],
    },
]
