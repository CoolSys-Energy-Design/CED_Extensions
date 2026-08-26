---
id: ae-pytools-select-on-sheets-title-blocks
doc_type: tool
title: Select on Sheets: Title Blocks
summary: Selects title blocks associated with views or sheets from the active context.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Select
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, sheets, title blocks]
aliases: []
last_verified: 2026-08-24
---

# Select on Sheets: Title Blocks

## What it does

Finds the title block placed on each selected sheet and selects those title block instances.

## When to use it

Use it when you have selected sheets and need to inspect, edit, or report on their title blocks.

## Before you start

- Select one or more sheets in the Project Browser.

## Steps

1. Select the required sheets in the Project Browser.
2. On **AE pyTools > Selection > Select**, click **Select on Sheets: Title Blocks**.
3. Work with the title blocks that become selected.

## Results and model impact

The current selection is replaced with title blocks owned by the selected sheets. The model is unchanged.

## Notes and limitations

- Only selected sheet items are considered; other selected element types are ignored.
- If no selected sheets contain title blocks, the command does not create a new selection.

## Related pages

- [Query Selection](query-selection.md)
- [Selection Tools](selection.md)
