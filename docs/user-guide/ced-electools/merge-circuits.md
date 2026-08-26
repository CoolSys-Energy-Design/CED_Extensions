---
id: ced-electools-merge-circuits
doc_type: tool
title: Merge Circuits
summary: Moves elements from compatible source circuits into a chosen main circuit and deletes emptied source circuits.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Merge Circuits
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: Changes circuit membership and deletes source circuits that become empty.
keywords: [electrical, circuits, merge, panel, circuit membership]
aliases: []
last_verified: "2026-08-24"
---

# Merge Circuits

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Compatible circuits should become one circuit. | Reassigns elements and deletes empty source circuits. | A chosen main circuit and compatible source circuits. |

## What it does

Uses a selected main circuit as the destination, moves elements from other selected compatible source circuits into it, then deletes source circuits that are empty afterward.

## When to use it

Use it to consolidate multiple circuits of the same voltage and pole configuration into one target circuit.

## Steps

1. Select circuits or circuited elements that you wish to merge into one circuit.
2. Run **Merge Circuits** .
3. Choose the main circuit and select compatible source circuits.
4. Confirm the merge selection.
5. Review the pyRevit output report.

## Results and verification

Verify source-device membership on the main circuit and confirm that only intended empty source circuits were removed.

## Notes and limitations

- Spares and spaces cannot be used as the main circuit or sources.
- Incompatible circuits are excluded and reported.
- The workflow rolls back if the merge operation fails.

## Related pages

- [Edit Circuit Properties](edit-circuit-properties.md)
- [Move Selected Circuits](move-selected-circuits.md)
