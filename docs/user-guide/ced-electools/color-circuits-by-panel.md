---
id: ced-electools-color-circuits-by-panel
doc_type: tool
title: Color Circuits by Panel
summary: Creates or updates panel-based graphic filters, a view template, and a legend for selected panels.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > QC Check > Color Circuits by Panel
navigation_path: [QC Check]
status: production
audience: [electrical]
model_impact: Creates or updates filters, a view template, and legend content, then activates temporary view-template mode.
keywords: [electrical, color, circuits, panels, filters, view template, legend]
aliases: [Panel Circuit Colors]
last_verified: "2026-08-24"
---

# Color Circuits by Panel


## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need a visual circuit-by-panel coordination view. | Creates or updates filters, a template, and legend content. | An active floor plan or RCP and electrical equipment with distribution systems. |

## What it does

Builds panel-specific circuit filters for selected panels, applies them through a view template, creates related legend content, and switches the active view into temporary template mode.

> [!TIP]
> Use this with [Move Selected Circuits](move-selected-circuits.md) to identify the panel supplying devices, then move their circuits to the correct target panel.

## When to use it

Use it for electrical coordination and QA when panel assignments need to be visually distinguished.

## Before you start

- Open a floor plan or reflected ceiling plan.
- Confirm electrical equipment has the intended panel names and distribution systems.
- Save the project because the command creates reusable model/view assets.

## Steps

1. Run **Color Circuits by Panel**.
2. Select the panels to include.
3. Let the command create or update filters, the template, and legend.
4. Review the temporary-template view result.

## Results and verification

Verify filter colors, the selected panel coverage, generated legend, and the active view's temporary view-template state.

## Notes and limitations

- The active view must be a floor plan or RCP.
- It stops when no panel names are available from electrical equipment with distribution systems.

## Related pages

- [Color Circuits Check](color-circuits-check.md)
- [Move Selected Circuits](move-selected-circuits.md)
- [Panel Report](panel-report.md)
