---
id: ced-electools-move-selected-circuits
doc_type: tool
title: Move Selected Circuits
summary: Transfers circuits associated with selected elements to another compatible panel.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Move Selected Circuits
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: modifies-circuit-panel-assignment
keywords: [circuits, panels, transfer, selected elements]
aliases: []
last_verified: "2026-08-24"
---

# Move Selected Circuits

## At a glance

| Use this when | Model impact                                                         | Required context                                                  |
|---|----------------------------------------------------------------------|-------------------------------------------------------------------|
| Circuits serving selected elements must move to a different panel. | Changes circuit(s) panel assignment. | Select elements on circuits and choose a compatible target panel. |

## What it does

Finds circuits associated with the selected elements and transfers them to a selected compatible panel.

> [!NOTE]
> You can select devices in plan or select circuits directly before running the tool. When devices are selected, the tool automatically extracts their circuits. It does NOT change what circuits the devices are a part of.


## When to use it

Use it when a layout or distribution change requires moving a focused group of loads to another panel.

> [!TIP]
> Use this with [Color Circuits by Panel](color-circuits-by-panel.md) to visualize which panels supply devices, then move the affected circuits to the correct target panel.

## Before you start

- Select elements that belong to the circuits to move.
- Confirm the target panel is compatible, has capacity, and is editable.

## Steps

1. Select the relevant circuited elements.
2. On **AE pyTools > Electrical > Circuit Tools**, click **Move Selected Circuits**.
3. Choose the destination panel and confirm the transfer.
4. Review the destination schedule and connected elements.

> [!TIP]
> Move Selected Circuits is also executable through the **Circuit Manager** interface. You can either select circuits in the manager and right click, or check circuits and execute from the **Actions** menu.

## Results and verification

Eligible circuits are reassigned to the destination panel. Verify circuit numbers, panel schedules, spare/space cleanup, and one-line information as applicable.

## Notes and limitations

- Incompatible panels, unavailable capacity, or worksharing ownership can block a move.
- The tool can replace default SPARE/SPACE entries when the target conditions allow it.

## Related pages

- [Batch Swap Circuits](batch-swap-circuits.md)
- [Color Circuits by Panel](color-circuits-by-panel.md)
- [Sync One-Line Data](sync-one-line-data.md)
- [Circuit Tools](circuit-tools.md)
