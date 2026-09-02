# -*- coding: utf-8 -*-
"""Narration for the Circuit Manager guided tour.

Pure data. Each step is a dict:

    chapter  - grouping label, drives the "Jump to" list
    title    - short heading in the speech bubble
    target   - Name of the element in CircuitBrowserPanel.xaml to spotlight, or
               None for a narration-only step. A target that is missing,
               collapsed or zero-sized degrades to narration-only.
    body     - one or two sentences of plain narration
    bullets  - optional detail lines rendered under the body

Keep this file free of WPF/Revit imports so it can be reviewed and edited by
anyone who knows the tool but not the codebase.
"""

TOUR_TITLE = "Circuit Manager - Show Me Around"

STEPS = [
    {
        "chapter": "Getting oriented",
        "title": "Evening. Let's walk the panel.",
        "target": None,
        "body": (
            "Circuit Manager is a dockable pane that lists every electrical circuit in the "
            "open model, so you can find, check, act on and recalculate circuits without "
            "opening a panel schedule for each one."
        ),
        "bullets": [
            "It docks like Project Browser - leave it open while you work.",
            "It reads the model on demand; nothing changes until you run an action.",
            "I will highlight each control as we go. Use Next and Back.",
        ],
    },
    {
        "chapter": "Getting oriented",
        "title": "Which model am I looking at?",
        "target": "DocumentNameText",
        "body": (
            "This line names the document the list is bound to. Circuit Manager follows the "
            "active document, so it re-binds when you switch models."
        ),
        "bullets": [
            "Switching documents or views triggers a rebind automatically.",
            "If this says something you did not expect, the list below is not your model.",
        ],
    },
    {
        "chapter": "Getting oriented",
        "title": "Refresh is your reset",
        "target": "RefreshButton",
        "body": (
            "Refresh rebuilds the list from current model state. The list is a snapshot, not a "
            "live link, so anything you change in Revit outside this pane needs a Refresh."
        ),
        "bullets": [
            "Run it after editing circuits, panels or connections in the model.",
            "Run it after someone relinquishes a workset you were blocked on.",
            "Run it when a row looks stale or a button is disabled for no obvious reason.",
        ],
    },
    {
        "chapter": "Getting oriented",
        "title": "The status line",
        "target": "StatusText",
        "body": (
            "Every load, filter and action reports here - counts, what was applied, and why "
            "something was skipped. It is the first place to look when a run does not do what "
            "you expected."
        ),
        "bullets": [],
    },
    {
        "chapter": "Finding circuits",
        "title": "Search across the whole row",
        "target": "SearchBox",
        "body": (
            "Type to filter. Search matches panel name, circuit number, load name and the rest "
            "of the row text, so a partial equipment tag or a panel name both work."
        ),
        "bullets": [
            "Filtering is live - no Enter needed.",
            "The X at the right clears the search and restores the full list.",
        ],
    },
    {
        "chapter": "Finding circuits",
        "title": "Filter by type, or by problem",
        "target": "FilterButton",
        "body": (
            "The filter menu has two halves. The top half toggles circuit types on and off. The "
            "bottom half holds the special filters, which are exclusive - turning one on shows "
            "all types so nothing hides behind a type toggle."
        ),
        "bullets": [
            "Types: Branch, Feeder, Space, Spare, XFMR PRI, XFMR SEC, Conduit Only, N/A.",
            "Circuits With Alerts - with a Show All / Active Alerts Only sub-choice.",
            "User Overrides - rows where someone has overridden a computed value.",
            "Failed Calculations - rows whose last calculation failed or was blocked.",
            "Checked Circuits Only - collapses the list to your working set.",
            "Reset Filters puts every type back and clears the special filters.",
        ],
    },
    {
        "chapter": "Finding circuits",
        "title": "The filter-is-on marker",
        "target": "FilterActiveMark",
        "body": (
            "This marker appears whenever the list is filtered. If a circuit you know exists is "
            "not in the list, check this before you go hunting in the model."
        ),
        "bullets": [],
    },
    {
        "chapter": "Finding circuits",
        "title": "List view or card view",
        "target": "ToggleViewButton",
        "body": (
            "Toggles between the compact list template and the card template. Compact fits far "
            "more rows in a narrow dock; cards show more per circuit."
        ),
        "bullets": [
            "The list is virtualized, so compact stays fast on large models.",
        ],
    },
    {
        "chapter": "Finding circuits",
        "title": "Options: theme, sort, density",
        "target": "BrowserOptionsButton",
        "body": (
            "The three-dot menu holds display preferences. Theme and accent are shared with the "
            "other AE pyTools windows through the AE-pyTools-Theme config, so setting them here "
            "sets them everywhere."
        ),
        "bullets": [
            "Theme: Light, Dark, Dark Alt.",
            "Accent Color: changes the highlight colour used across the tools.",
            "Display Mode: Compact or Card - same as the toggle we just looked at.",
            "Sort By: Circuit, Load Name, or Rating.",
            "List View: Show Circuit Type Badges, Compress Item Width.",
        ],
    },
    {
        "chapter": "Reading a row",
        "title": "What the badges are telling you",
        "target": "CircuitList",
        "body": (
            "Each row carries its state as badges, so you can triage without opening anything. "
            "Left to right: circuit type, then the conductor and condition markers."
        ),
        "bullets": [
            "Type badge - BRANCH, FEEDER, SPACE, SPARE, XFMR PRI, XFMR SEC, CONDUIT ONLY, N/A.",
            "N - a neutral is included on this circuit.",
            "IG - an isolated ground is included.",
            "Override icon - a user override is set on this circuit.",
            "Sync lock icon - writeback is blocked by ownership; the tooltip names the owner.",
            "Alert (!) - this circuit has saved alerts; click it to read them.",
        ],
    },
    {
        "chapter": "Reading a row",
        "title": "Right-click a row for row-specific work",
        "target": "CircuitList",
        "body": (
            "The row context menu does everything that only makes sense for the row under your "
            "cursor, without disturbing the checked set you built up in the toolbar."
        ),
        "bullets": [
            "Select in Model - Panel, Circuit, or Device.",
            "Show Devices in Model / Show Panel in Model.",
            "Edit Circuit Properties.",
            "Move Selected Circuits.",
            "Inject Into Detail Item.",
            "Show Circuit Type Badges / Compress Item Width density toggles.",
        ],
    },
    {
        "chapter": "Reading a row",
        "title": "Checked is not the same as selected",
        "target": "CheckAllButton",
        "body": (
            "This trips people up, so it is worth being precise. Checked circuits are the scope "
            "most batch tools act on. Selected circuits are just your current highlight. A row "
            "can be one, the other, or both."
        ),
        "bullets": [
            "Click selects a single row. Ctrl+Click adds or removes one row.",
            "Shift+Click selects the contiguous range from the last focused row.",
            "Check All / Uncheck All operate on what the current filter is showing.",
            "Checked Circuits Only in the filter menu collapses to your working set.",
        ],
    },
    {
        "chapter": "Selecting in the model",
        "title": "Jump to the supplying equipment",
        "target": "SelectEquipmentButton",
        "body": (
            "Selects the upstream panel or equipment feeding the checked or selected circuits, "
            "so you can confirm you are working on the right distribution before you change it."
        ),
        "bullets": [],
    },
    {
        "chapter": "Selecting in the model",
        "title": "Select the circuits themselves",
        "target": "SelectCircuitsButton",
        "body": (
            "Selects the electrical system elements for the checked or selected rows. Useful "
            "for driving Revit's own properties palette or a schedule from this list."
        ),
        "bullets": [],
    },
    {
        "chapter": "Selecting in the model",
        "title": "Select what is downstream",
        "target": "SelectDownstreamButton",
        "body": (
            "Selects the connected devices on those circuits. This is the quickest visual check "
            "that a circuit actually picks up the equipment you think it does."
        ),
        "bullets": [],
    },
    {
        "chapter": "Selecting in the model",
        "title": "Clear the selection",
        "target": "ClearSelectionButton",
        "body": (
            "Clears the Revit selection without touching your checked set here. Use it between "
            "verification passes so old highlights do not confuse the next check."
        ),
        "bullets": [],
    },
    {
        "chapter": "Acting on circuits",
        "title": "The Actions menu",
        "target": "ActionsButton",
        "body": (
            "This is where the batch work happens. Every action runs against your checked "
            "circuits, and each one opens its own review window first - nothing is written to "
            "the model straight from this menu."
        ),
        "bullets": [
            "Add/Remove Neutral.",
            "Add/Remove IG.",
            "Auto Size Breaker.",
            "Mark as New/Existing.",
            "Edit Circuit Properties.",
            "Move Selected Circuits.",
        ],
    },
    {
        "chapter": "Acting on circuits",
        "title": "How a review window behaves",
        "target": None,
        "body": (
            "The action windows all share the same shape: a row per affected circuit, per-row "
            "checkboxes, and Apply / Cancel. You stage what you want, then commit once."
        ),
        "bullets": [
            "Rows the action cannot support are hidden until you tick show-unsupported.",
            "Ownership-blocked rows are marked and excluded from the apply.",
            "Cancel writes nothing - closing the window is always safe.",
            "After apply, a run summary reports what succeeded and what was skipped.",
        ],
    },
    {
        "chapter": "Calculating",
        "title": "Calculate All or Calculate Selected",
        "target": "CalcAllButton",
        "body": (
            "All calculates everything the current context is showing. Selected calculates only "
            "your checked or selected rows - that is the one to use on a large model."
        ),
        "bullets": [
            "Calculation respects ownership: locked circuits are reported, not forced.",
            "Recalculate after any action that changes load, conductor or breaker sizing.",
        ],
    },
    {
        "chapter": "Calculating",
        "title": "Calculate settings",
        "target": "CalcSettingsButton",
        "body": (
            "The gear opens the calculation settings used by both Calculate buttons. Set these "
            "once for the project before you run a large recalculation."
        ),
        "bullets": [],
    },
    {
        "chapter": "Calculating",
        "title": "Preview before writeback",
        "target": "CalcPreviewToggle",
        "body": (
            "With preview On, a calculation run shows you the computed result and waits for your "
            "decision before writing anything back to the model. With it Off, results are "
            "written directly."
        ),
        "bullets": [
            "The On/Off state is remembered between sessions.",
            "Leave it On while you are still trusting a new set of settings.",
        ],
    },
    {
        "chapter": "The one line",
        "title": "See the distribution as a tree",
        "target": "ShowOneLineButton",
        "body": (
            "Opens an interactive one line diagram of the electrical distribution. It is the "
            "fastest way to understand what feeds what - and you can re-feed a panel by "
            "dragging its card onto a new source."
        ),
        "bullets": [
            "Drag an equipment card onto another to re-feed that panel.",
            "Zoom in / out / reset, and toggle card mode and loads mode.",
            "Build a single branch or build the whole tree at once.",
            "Refresh re-reads the distribution after model changes.",
            "The button is unavailable in drafting views - open a model view first.",
        ],
    },
    {
        "chapter": "When something is blocked",
        "title": "Why a button goes grey",
        "target": None,
        "body": (
            "Disabled controls almost always come down to one of three things, and all three "
            "are visible from the pane."
        ),
        "bullets": [
            "Nothing is checked or selected - most actions need a scope.",
            "Ownership: the circuit or a downstream element is owned by another user. The sync "
            "lock badge and its tooltip name the owner.",
            "A filter is hiding what you expected - watch for the filter-active marker.",
            "When in doubt: Refresh, then re-check your scope.",
        ],
    },
    {
        "chapter": "When something is blocked",
        "title": "That is the whole tour",
        "target": None,
        "body": (
            "Find with Search and Filter, build a scope with the checkboxes, verify it with "
            "Select in Model, run an Action, then Calculate and clear the alerts. That loop is "
            "the whole tool."
        ),
        "bullets": [
            "Alerts Manager and BatchSwap are separate buttons on this panel and pick up where "
            "this pane leaves off.",
            "Longer written guides live in README.v2.md next to this tool.",
            "Right-click the Circuit Manager button on the ribbon any time to bring me back.",
        ],
    },
]


def chapters():
    """Ordered chapter names paired with the index of their first step."""
    found = []
    for index, step in enumerate(STEPS):
        name = step.get("chapter") or ""
        if not found or found[-1][0] != name:
            found.append((name, index))
    return found
