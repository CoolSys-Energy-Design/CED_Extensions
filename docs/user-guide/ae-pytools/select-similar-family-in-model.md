---
id: ae-pytools-select-similar-family-in-model
doc_type: tool
title: Select Similar: Family in Model
summary: Selects model elements matching the chosen family.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Select
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, similar, family, model]
aliases: []
last_verified: 2026-08-24
---

# Select Similar: Family in Model

## What it does

Replaces the selection with all instances of the same family as a selected example element, across the model.

## When to use it

Use it when a family-wide update or review is needed across all views and types of a family.

## Before you start

- Select a family instance to use as the example.

## Steps

1. Select one family instance.
2. On **AE pyTools > Selection > Select**, click **Select Similar: Family in Model**.
3. Review the project-wide selection before making changes.

## Results and model impact

Matching family instances in the model become the active selection. The model is unchanged.

## Notes and limitations

- This is broader than a type-based selection: different types within the same family can be included.
- The command is intended for family instances; system-family workflows may not behave as expected.

## Related pages

- [Select Similar: Family in View](select-similar-family-in-view.md)
- [Select Similar: Category in Model](select-similar-category-in-model.md)
- [Selection Tools](selection.md)
