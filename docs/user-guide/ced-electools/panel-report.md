---
id: ced-electools-panel-report
doc_type: tool
title: Panel Report
summary: Produces a panel-focused review report of circuit and breaker information.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > QC Check > Panel Report
navigation_path: [QC Check]
status: production
audience: [electrical]
model_impact: report-only
keywords: [panels, breakers, circuits, report, quality control]
aliases: [PanelBreakerList]
last_verified: "2026-08-24"
---

# Panel Report

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Panel and breaker data need a focused review or comparison against submittals. | None - creates a review report only. | A project with panels and circuits to report. |

## What it does

Creates a panel-focused listing of circuit and breaker information for review.

## When to use it

Use it to check panel loading, breaker data, and schedule information before coordination or submittal review.

## Before you start

- Confirm current circuit and panel data has been saved.

## Steps

1. On **AE pyTools > Electrical > QC Check**, click **Panel Report**.
2. Select the report scope when prompted.
3. Review the generated report.

## Results and verification

The command produces a panel review report without editing the model. Compare report values to panel schedules and applicable manufacturer information.

## Notes and limitations

- Report values are only as current as the circuit and panel data in the model.

## Related pages

- [Electrical System Check](electrical-system-check.md)
- [Calculate Circuits](calculate-circuits.md)
- [CED ElecTools](index.md)
