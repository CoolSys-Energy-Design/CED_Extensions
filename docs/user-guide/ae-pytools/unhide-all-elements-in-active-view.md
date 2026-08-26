---
id: ae-pytools-unhide-all-elements-in-active-view
doc_type: tool
title: Unhide All Elements in Active View
summary: Unhides manually hidden elements in the active view and selects the restored items.
extension: AE pyTools
ribbon_path: AE pyTools > Miscellaneous > Views > Unhide All Elements in Active View
navigation_path: [Miscellaneous]
status: production
audience: [all]
model_impact: modifies-active-view
keywords: [view, hidden elements, visibility, selection]
aliases: []
last_verified: "2026-08-24"
---

# Unhide All Elements in Active View

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Elements may have been manually hidden in the current view. | Restores manually hidden elements in the active view and selects them. | Open the view to repair. |

## What it does

Unhides elements that were manually hidden in the active view, then makes those restored elements the current selection.

## When to use it

Use it to recover from temporary Hide in View actions when the hidden items are unknown.

## Before you start

- Open the view whose manually hidden elements should be restored.

## Steps

1. On **AE pyTools > Miscellaneous > Views**, click **Unhide All Elements in Active View**.
2. Review the restored elements, which become selected.
3. Reapply intentional visibility control as needed.

## Results and verification

Manually hidden elements in the active view are restored and selected. The model elements themselves are not edited.

## Notes and limitations

- The command targets manual hiding in the active view; it does not override view templates, filters, worksets, or category visibility.
- If no manually hidden elements exist, the command reports that no items were found.

## Related pages

- [Toggle Grid Bubbles](toggle-grid-bubbles.md)
- [Miscellaneous Tools](miscellaneous.md)
