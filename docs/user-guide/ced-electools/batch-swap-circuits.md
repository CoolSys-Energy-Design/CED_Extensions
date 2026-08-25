---
id: ced-electools-batch-swap-circuits
doc_type: tool
title: Batch Swap Circuits
summary: Stages and applies multiple slot-level circuit, spare, and space changes across panel schedules.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Batch Swap Circuits
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: modifies-circuits-and-panel-schedules
keywords: [circuits, panels, panel schedules, spare, space, transfer]
aliases: [BatchSwapCircuits]
last_verified: "2026-08-24"
---

# Batch Swap Circuits

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Several panel-schedule slot changes must be planned and applied together. | Moves circuits and can add, remove, or replace SPARE and SPACE entries. | Review source and target panel schedules before applying changes. |

## What it does

Lets you stage multiple slot-level circuit transfers and default SPARE/SPACE changes, review their sequence, and apply them as one workflow.

## When to use it

Use it to redistribute circuits within a panel or between compatible panels without editing each schedule slot manually.

## Before you start

- Confirm source and target panels are compatible and editable.
- Resolve worksharing ownership and review available poles or slots.

## Steps

1. On **AE pyTools > Electrical > Circuit Tools**, click **Batch Swap Circuits**.
2. Select source and target panel-schedule slots and stage the required operations.
3. Review the planned order and target conditions.
4. Apply the staged changes and inspect the affected schedules.

## Results and verification

The approved staged operations are applied to the panel schedules. Verify circuit destinations, slot positions, SPARE/SPACE entries, and downstream circuit data.

## Notes and limitations

> [!WARNING]
> This can change multiple circuit assignments and schedule slots in one operation. Review all staged operations before applying them.

## Related pages

- [Move Selected Circuits](move-selected-circuits.md)
- [Create Dedicated Circuits](create-dedicated-circuits.md)
- [Circuit Tools](circuit-tools.md)
