---
id: ae-pytools-quick-dimension
doc_type: tool
title: Quick Dimension
summary: Creates one dimension from selected elements using an inferred or picked linear reference direction.
extension: AE pyTools
ribbon_path: AE pyTools > Miscellaneous > Quick Dimension
navigation_path: [Miscellaneous]
status: production
audience: [all]
model_impact: Creates a dimension element in the active view.
keywords: [dimension, quick dimension, grids, pipes, ducts, reference planes]
aliases: []
last_verified: "2026-08-24"
---

# Quick Dimension

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need a dimension between several compatible linear or family references. | Creates one dimension in the active view. | At least two valid references and a linear direction. |

## What it does

Creates a dimension from the current selection or picked elements. It infers a direction from linear selected items when possible, otherwise asks you to pick a line, grid, detail line, or reference plane.

## When to use it

Use it for a quick dimension across linear model elements, grids, lines, or family instances with usable center reference planes.

## Before you start

- Select compatible elements, or be ready to pick them.
- Work in a view where the dimension can be created.
- For family instances, ensure usable center reference planes are defined.

## Steps

1. Select the elements, then run **Quick Dimension**.
2. Pick a lead direction if the command cannot infer one.
3. Pick the dimension-line location.
4. Review the created dimension.

## Results and verification

The new dimension is selected. Confirm witness lines, reference choices, and placement; the output lists any selected elements that could not be dimensioned.

## Notes and limitations

- Full support is for linear elements such as pipes, ducts, lines, and grids.
- At least two valid references are required.

## Related pages

- [Copy V/G Settings to View Templates](copy-vg-settings-to-view-templates.md)
