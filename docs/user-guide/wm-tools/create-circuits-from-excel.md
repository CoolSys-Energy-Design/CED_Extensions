---
id: wm-tools-create-circuits-from-excel
doc_type: tool
title: Create Circuits From Excel
summary: Places configured electrical symbols and creates circuits from selected workbook sheets.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Create Circuits From Excel
navigation_path: []
status: production
audience: [electrical, refrigeration]
model_impact: May load families, place instances, create circuits, assign panels, and remove qualifying space circuits.
keywords: [wm, circuits, excel, panels, connectors, refrigeration]
aliases: [Circuit Creation]
last_verified: "2026-08-24"
---

# Create Circuits From Excel

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Approved circuit data is ready in the WM workbook. | Creates and assigns model instances and electrical systems. | Valid workbook sheets, configured panels, and matching electrical content. |

## What it does

Reads selected workbook sheets, places the configured symbols, creates electrical circuits, and attempts to assign each circuit to its named panel.

## When to use it

Use it after panels and WM content are available and the source data has been coordinated.

## Before you start

- Verify workbook rows, panel names, family names, types, and circuit data.
- Ensure target panels have placeable faces and required family content is available.
- Save the project; this is a high-impact batch operation.

## Steps

1. Run **Create Circuits From Excel** and select the valid sheets.
2. Allow the command to resolve or load missing configured families.
3. Review placement, circuit, assignment, and cleanup messages in the pyRevit output.

## Results and verification

Check placed symbols, circuit membership, panel assignments, and circuit numbering against the workbook. Resolve reported missing symbols, panels, or failed assignments manually.

## Notes and limitations

- The command sets circuit sequencing to `OddThenEven` when needed.
- It can remove zero-rated space circuits whose load name contains `space`.
- Rows without a matching family/type, panel surface, or valid data are reported and skipped.

## Related pages

- [Load Electrical Content](load-electrical-content.md)
- [Create Panels From Excel](create-panels-from-excel.md)
- [Replace Existing Circuit](replace-existing-circuit.md)
