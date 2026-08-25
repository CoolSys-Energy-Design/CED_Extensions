---
id: ae-pytools-select-similar-category-in-view
doc_type: tool
title: Select Similar: Category in View
summary: Selects visible elements matching the chosen element category.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Select
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, similar, category, view]
aliases: []
last_verified: 2026-08-24
---

# Select Similar: Category in View

## What it does

Replaces the selection with elements visible in the active view that share the category of the currently selected elements.

## When to use it

Use it for view-specific cleanup, tagging, or review when selecting the same category across the full model would be too broad.

## Before you start

- Open the view you want to work in.
- Select one or more example elements whose categories you want to find.

## Steps

1. Select the example element or elements.
2. On **AE pyTools > Selection > Select**, click **Select Similar: Category in View**.
3. Confirm the resulting selection before continuing.

## Results and model impact

Matching category elements in the active view become the active selection. The model is unchanged.

## Notes and limitations

- Elements outside the active view are not included.
- View visibility and filters affect which elements can be selected.

## Related pages

- [Select Similar: Category in Model](select-similar-category-in-model.md)
- [Select Similar: Family in View](select-similar-family-in-view.md)
- [Selection Tools](selection.md)
