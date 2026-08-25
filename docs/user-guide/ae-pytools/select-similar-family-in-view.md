---
id: ae-pytools-select-similar-family-in-view
doc_type: tool
title: Select Similar: Family in View
summary: Selects visible elements matching the chosen family.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Select
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, similar, family, view]
aliases: []
last_verified: 2026-08-24
---

# Select Similar: Family in View

## What it does

Replaces the selection with visible instances of the same family as a selected example element.

## When to use it

Use it to work on a single family in the current view without selecting its instances elsewhere in the project.

## Before you start

- Open the view you want to work in.
- Select a family instance to use as the example.

## Steps

1. Select one family instance.
2. On **AE pyTools > Selection > Select**, click **Select Similar: Family in View**.
3. Review the resulting selection before continuing.

## Results and model impact

Matching visible family instances become the active selection. The model is unchanged.

## Notes and limitations

- Elements outside the active view are not included.
- This includes family types in the same family; use [Filter Selection by Type](filter-selection-by-type.md) if you need a narrower result.

## Related pages

- [Select Similar: Family in Model](select-similar-family-in-model.md)
- [Filter Selection by Type](filter-selection-by-type.md)
- [Selection Tools](selection.md)
