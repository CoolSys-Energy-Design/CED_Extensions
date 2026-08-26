---
id: ae-pytools-filter-selection-by-type
doc_type: tool
title: Filter Selection by Type
summary: Replaces the active selection with chosen category, family, and type groups.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Filter Selection
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, filter, category, family, type]
aliases: []
last_verified: 2026-08-24
---

# Filter Selection by Type

## What it does

Groups the current selection by category, family, and type, then replaces the selection with the groups you choose.

## When to use it

Use it when a selection contains several kinds of elements and you need to keep only specific family types.

## Before you start

- Select one or more elements.

## Steps

1. Select the elements to refine.
2. On **AE pyTools > Selection > Filter Selection**, click **Filter Selection by Type**.
3. Choose one or more listed category, family, and type groups.
4. Click **OK**.

## Results and model impact

The chosen groups become the active selection. The model is unchanged.

## Notes and limitations

- The command requires a non-empty selection.
- Cancelling the picker leaves the original selection unchanged.
- Group labels include the number of matching selected elements.

## Related pages

- [Discard Pinned Elements](discard-pinned-elements.md)
- [Query Selection](query-selection.md)
- [Selection Tools](selection.md)
