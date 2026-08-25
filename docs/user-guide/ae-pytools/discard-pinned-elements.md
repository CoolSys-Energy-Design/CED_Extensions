---
id: ae-pytools-discard-pinned-elements
doc_type: tool
title: Discard Pinned Elements
summary: Removes pinned elements from the active selection without changing the model.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Filter Selection
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, pinned, filter]
aliases: []
last_verified: 2026-08-24
---

# Discard Pinned Elements

## What it does

Removes pinned elements from the current Revit selection. It does not unpin, modify, or delete any element.

## When to use it

Use it before an operation that should apply only to editable, unpinned items in a mixed selection.

## Before you start

- Select the elements you want to review or edit.

## Steps

1. Select the required elements in Revit.
2. On **AE pyTools > Selection > Filter Selection**, click **Discard Pinned Elements**.
3. Continue working with the remaining selection.

## Results and model impact

Pinned items are removed from the selection; all other selected items remain selected. The model is unchanged.

## Notes and limitations

- An empty selection remains empty.
- This command does not report which elements were removed.

## Related pages

- [Filter Selection by Type](filter-selection-by-type.md)
- [Selection Tools](selection.md)
