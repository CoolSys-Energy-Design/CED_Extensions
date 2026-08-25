---
id: ae-pytools-acc-path-resolver
doc_type: tool
title: ACC Path Resolver
summary: Manually approves the CED ACC content root for telemetry routing or prints read-only path diagnostics.
extension: AE pyTools
ribbon_path: AE pyTools > CED Tools > ACC Path Resolver
navigation_path: [CED Tools]
status: beta
audience: [administrators, support]
model_impact: Does not change Revit model elements; writes telemetry routing state and can recover stale telemetry files.
keywords: [acc, desktop connector, telemetry, path, diagnostics, routing]
aliases: [ADC Diagnostics]
last_verified: "2026-08-24"
---

# ACC Path Resolver

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| CED telemetry cannot resolve a valid ACC content root. | No model edits; changes local telemetry-routing state. | A synced CED Content Collection location. |

## What it does

Clicking the command lets you manually approve a valid CED Content Collection root. Shift+click prints read-only routing diagnostics, candidate scores, and state information to pyRevit output.

## When to use it

Use it only for telemetry-routing support when the automatically detected location is missing or ambiguous.

## Before you start

- Confirm that Desktop Connector has synced the correct CED Content Collection.
- Use Shift+click first when you need evidence before changing the saved route.

## Steps

1. Shift+click **ACC Path Resolver** to review diagnostics, if needed.
2. Click the command normally and confirm the manual-resolution prompt.
3. Browse to the CED Content Collection root and confirm the result message.

## Results and verification

Verify the saved root, resolved status, user folder, and stale-file recovery counts in the completion dialog or diagnostics output.

## Notes and limitations

- The selected root must contain the expected project-files and usage locations.
- This is a support configuration command; it is not a model-management workflow.

## Related pages

- [About](about.md)
