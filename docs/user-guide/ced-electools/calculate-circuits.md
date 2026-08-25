---
id: ced-electools-calculate-circuits
doc_type: tool
title: Calculate Circuits
summary: Calculates electrical circuit sizing and related values, then writes applicable results to the model.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Calculate Circuits
navigation_path: []
status: production
audience: [electrical]
model_impact: updates-circuit-and-connected-element-data
keywords: [circuits, calculation, wire size, conduit size, voltage drop, breaker]
aliases: []
last_verified: "2026-08-24"
---

# Calculate Circuits

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Circuit sizing and output parameters need to reflect the current model. | Calculates and writes applicable circuit and connected-element values. | Electrical parameters must be loaded and circuit data must be ready for calculation. |

## What it does

Calculates circuit values such as breaker rating, conductor and conduit sizing, and voltage drop, then writes the applicable results back to circuits and connected elements.

## When to use it

Use it after circuiting or materially changing electrical loads, lengths, panels, or sizing inputs.

## Before you start

- Run [Load Electrical Parameters](load-electrical-parameters.md) for projects that were not created from the CED template.
- Review circuit inputs and resolve known warnings before calculating.

## Steps

1. On **AE pyTools > Electrical**, click **Calculate Circuits**.
2. Review the calculation settings and scope.
3. Run the calculation.
4. Review reported warnings and calculated values.

## Results and verification

Applicable calculated values are written to the model. Verify representative circuit, wire, conduit, breaker, and voltage-drop results before publishing schedules or one-line diagrams.

## Notes and limitations

> [!WARNING]
> This updates model data. Correct invalid circuit inputs and review warnings instead of treating a completed run as design approval.

## Related pages

- [Load Electrical Parameters](load-electrical-parameters.md)
- [Electrical System Check](electrical-system-check.md)
- [Sync One-Line Data](sync-one-line-data.md)
