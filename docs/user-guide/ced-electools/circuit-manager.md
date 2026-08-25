---
id: ced-electools-circuit-manager
doc_type: tool
title: Circuit Manager
summary: Opens the dockable circuit workspace for searching, filtering, selecting, editing, moving, and recalculating circuits.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Manager
navigation_path: []
status: production
audience: [electrical]
model_impact: Actions launched from the manager can edit, move, or recalculate circuits.
keywords: [electrical, circuits, manager, search, filter, dockable]
aliases: [Circuit Workspace]
last_verified: "2026-08-24"
---

# Circuit Manager

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need one dockable workspace for circuit review and actions. | Depends on the action you run from the manager. | An open electrical project. |

## What it does

Opens the dockable circuit manager, where you can search and filter circuits, select related model elements, and initiate supported move, alert, editing, or recalculation workflows.

## When to use it

Use it for ongoing circuit coordination rather than one-off selection or report commands.

## Before you start

- Refresh the manager after material model changes.
- Confirm the checked or selected circuits before running an editing action.

## Steps

1. Run **Circuit Manager**.
2. Use search and filters to find the target circuits.
3. Select or check the circuits, then choose the required supported action.
4. Refresh and verify the result.

## Results and verification

The window stays docked for continued use. Verify circuit membership, panel assignment, and recalculated values in the manager and model.

## Notes and limitations

- If the dockable pane is unavailable, the command reports that it cannot open it.
- Individual actions can have their own prerequisites and model impact.

## Related pages

- [Alerts Manager](alerts-manager.md)
- [Edit Circuit Properties](edit-circuit-properties.md)
- [Move Selected Circuits](move-selected-circuits.md)
