---
id: ced-electools-electrical-system-check
doc_type: tool
title: Electrical System Check
summary: Runs electrical circuit and system validation checks for coordination and quality control.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > QC Check > Electrical System Check
navigation_path: [QC Check]
status: production
audience: [electrical]
model_impact: report-only
keywords: [electrical, quality control, circuits, panels, validation]
aliases: []
last_verified: "2026-08-24"
---

# Electrical System Check

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Electrical panels, circuits, and devices need a pre-coordination or pre-issue review. | None - reports validation findings. | A project with electrical systems to review. |

## What it does

Runs configured electrical-system validation checks and reports conditions that need review, such as incompatible or incomplete circuit and panel data.

## When to use it

Use it after significant circuit edits and before issuing electrical deliverables.

## Before you start

- Save or synchronize current electrical work.
- Run [Calculate Circuits](calculate-circuits.md) first when sizing-related results are required.

## Steps

1. On **AE pyTools > Electrical > QC Check**, click **Electrical System Check**.
2. Run the available checks.
3. Review each reported finding and correct the associated model condition.
4. Re-run the check to confirm resolution.

## Results and verification

The command reports validation findings without editing the model. Treat each result as a coordination item to investigate and resolve.

## Notes and limitations

- A clean result does not replace project-specific engineering review.
- Calculation-dependent findings can be stale until calculations are rerun.

## Related pages

- [Calculate Circuits](calculate-circuits.md)
- [Panel Report](panel-report.md)
- [CED ElecTools](index.md)
