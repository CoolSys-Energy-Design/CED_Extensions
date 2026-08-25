---
id: ced-electools-create-dedicated-circuits
doc_type: tool
title: Create Dedicated Circuits
summary: Creates one native Revit circuit for each selected element and assigns it to a chosen panel.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Create Dedicated Circuits
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: creates-circuits
keywords: [circuits, dedicated circuits, panels, equipment]
aliases: []
last_verified: "2026-08-24"
---

# Create Dedicated Circuits

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Each selected device or piece of equipment needs its own circuit. | Creates circuits and assigns them to a panel. | Select circuitable elements and choose an appropriate panel. |

## What it does

Creates one dedicated native Revit circuit per selected element, assigns the circuits to a chosen panel, and provides a review table for key circuit values.

## When to use it

Use it for equipment that requires an individual branch circuit, rather than a shared circuit.

## Before you start

- Select compatible circuitable elements.
- Confirm the target panel has the required available capacity and distribution compatibility.

## Steps

1. Select the elements that need dedicated circuits.
2. On **AE pyTools > Electrical > Circuit Tools**, click **Create Dedicated Circuits**.
3. Choose the target panel and review the proposed results.
4. Create the circuits and review their key values.

## Results and verification

One circuit is created for each eligible selected element and assigned to the chosen panel. Confirm circuit numbers, panel assignment, loads, poles, and schedule placement.

## Notes and limitations

- Incompatible selected elements or insufficient panel capacity can prevent creation.
- Review the resulting circuit table before treating the operation as complete.

## Related pages

- [Batch Swap Circuits](batch-swap-circuits.md)
- [Load Electrical Parameters](load-electrical-parameters.md)
- [Circuit Tools](circuit-tools.md)
