---
id: ae-pytools-query-selection
doc_type: tool
title: Query Selection
summary: Reports detailed information about the active Revit selection.
extension: AE pyTools
ribbon_path: AE pyTools > Selection > Selection
navigation_path: [Selection]
status: production
audience: all
model_impact: none
keywords: [selection, parameters, report, query]
aliases: []
last_verified: 2026-08-24
---

# Query Selection

## What it does

Creates a pyRevit output table for the current selection, showing chosen instance or type parameter values. Element IDs in the table link back to the corresponding Revit elements.

## When to use it

Use it to compare parameter values across selected equipment, devices, or other elements without creating a schedule.

## Before you start

- Select the elements to inspect.
- For a mixed selection, be ready to choose the categories to include.

## Steps

1. Select the elements to inspect.
2. On **AE pyTools > Selection**, click **Query Selection**.
3. If prompted, select the categories to include.
4. Select the parameters to display.
5. Review the output table and use its Element ID links to navigate to individual elements.

## Results and model impact

The command opens a report in the pyRevit output window. It does not change the model or the selection.

## Notes and limitations

- Parameters are collected from both selected instances and their types.
- A cross mark in the report means that parameter does not exist for that element.
- Selecting many elements or parameters can take longer to process.

## Related pages

- [Filter Selection by Type](filter-selection-by-type.md)
- [Zoom to Selection](zoom-to-selection.md)
- [Selection Tools](selection.md)
