---
id: ae-pytools-zoom-to-selection
doc_type: tool
title: Zoom to Selection
summary: Zooms the active view to the bounding area of the selected elements.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Selection
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, zoom, view, navigation]
aliases: []
last_verified: 2026-08-24
---

# Zoom to Selection

## What it does

Centers the active Revit view on the current selection and zooms to its visible bounding area.

## When to use it

Use it when selected elements are difficult to locate in a large plan, section, elevation, or 3D view.

## Before you start

- Open the view you want to focus.
- Select one or more elements with visible bounding boxes in that view.

## Steps

1. Select the elements to locate.
2. On **AE pyTools > Selection**, click **Zoom to Selection**.
3. Continue reviewing or editing the focused selection.

## Results and model impact

The active view is centered and zoomed around the selection. The model and selection are unchanged.

## Notes and limitations

- Elements without a bounding box in the active view are ignored.
- The current command reads a saved zoom-factor setting but does not provide the PDF-era Shift-click settings prompt.
- Use it from an active view; sheet/viewport navigation is not the intended workflow.

## Related pages

- [Query Selection](query-selection.md)
- [Selection Tools](selection.md)
