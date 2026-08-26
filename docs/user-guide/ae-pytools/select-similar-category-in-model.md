---
id: ae-pytools-select-similar-category-in-model
doc_type: tool
title: Select Similar: Category in Model
summary: Selects model elements matching the chosen element category.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Select
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, similar, category, model]
aliases: []
last_verified: 2026-08-24
---

# Select Similar: Category in Model

## What it does

Replaces the selection with elements in the model that share the category of the currently selected elements.

## When to use it

Use it to gather all instances of a Revit category for a project-wide review or batch operation.

## Before you start

- Select one or more example elements whose categories you want to find.

## Steps

1. Select the example element or elements.
2. On **AE pyTools > Selection > Select**, click **Select Similar: Category in Model**.
3. Confirm the resulting selection before running a model-changing command.

## Results and model impact

Matching category elements become the active selection. The model is unchanged.

## Notes and limitations

- This searches the model rather than only the active view.
- A broad category can produce a very large selection; inspect it before continuing.

## Related pages

- [Select Similar: Category in View](select-similar-category-in-view.md)
- [Select Similar: Family in Model](select-similar-family-in-model.md)
- [Selection Tools](selection.md)
