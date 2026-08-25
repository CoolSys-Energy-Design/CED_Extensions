---
id: ae-pytools-parameter-monitor
doc_type: tool
title: Parameter Monitor
summary: Tracks selected host or linked-model parameters against an accepted project baseline in a modeless window.
extension: AE pyTools
ribbon_path: AE pyTools > CED Tools > Parameter Monitor
navigation_path: [CED Tools]
status: production
audience: [all]
model_impact: Does not edit monitored elements; writes only its own project storage for tracking data.
keywords: [parameters, monitor, baseline, linked model, changes, modeless]
aliases: [Parameter Tracking]
last_verified: "2026-08-24"
---

# Parameter Monitor

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need to track parameter changes against an accepted baseline. | Monitored elements are read-only; the tool saves its tracking data in project storage. | An active project and elements or linked-model data to monitor. |

## What it does

Opens a modeless monitoring window for creating tracking sets, reviewing monitored elements, comparing values to a stored baseline, filtering results, and navigating to elements.

## When to use it

Use it to monitor coordinated parameter data over time without changing the monitored model values.

## Before you start

- Decide which elements and parameters should define the tracking set.
- Review and explicitly accept the baseline before treating later differences as changes.

## Steps

1. Run **Parameter Monitor**.
2. Create or choose a tracking set, then select the target elements and parameters.
3. Accept the baseline when the values are correct.
4. Refresh later to review changed, missing, or matching values.

## Results and verification

Use the status, parameter-difference, missing-value, and location columns to review results. Select a row to inspect the associated element and properties.

## Notes and limitations

- The window is modeless; rerunning the command activates the existing window when present.
- It does not write values back to monitored host or linked elements.
- Tool storage is separate from the underlying model parameters being monitored.

## Related pages

- [Theme](theme.md)
