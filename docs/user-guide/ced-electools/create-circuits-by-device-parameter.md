---
id: ced-electools-create-circuits-by-device-parameter
doc_type: tool
title: Create Circuits by Device Parameter
summary: Groups circuitable devices by a selected parameter and creates native Revit circuits from those groups.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Create Circuits by Device Parameter
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: Creates electrical circuits and can assign them to selected panels.
keywords: [electrical, circuits, device parameter, panel assignment, batch creation]
aliases: [Create Circuits by Parameter]
last_verified: "2026-08-24"
---

# Create Circuits by Device Parameter

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Devices share a parameter that defines their intended circuit grouping. | Creates circuits and may assign them to panels. | Selected circuitable devices, or confirmation to scan the model. |

## What it does

Groups circuitable devices by a chosen parameter, lets you review and regroup them, then creates one native Revit circuit per group with selected panel, breaker, and load values.

## When to use it

Use it to create many circuits from coordinated device data instead of manually circuiting each group.

## Before you start

- Populate and verify the grouping parameter, such as `CKT_Circuit Number_CEDT`.
- Select target devices to avoid scanning the entire model.
- Confirm target panels, breaker ratings, and load data.

## Steps

1. Select devices, then run **Create Circuits by Device Parameter**.
2. Choose the grouping parameter and review or adjust groups.
3. Choose panel assignment and circuit settings.
4. Create the circuits and review the pyRevit output.

## Results and verification

Verify created circuit count, device membership, panel assignments, and any unresolved assignment rows in the output.

## Notes and limitations

- With no selection, the command requires confirmation before scanning the full model.
- Devices without a primary power connector, or with an already-used primary connector, are skipped and reported.
- Created circuits can be retained even when some panel assignments need manual resolution.

## Related pages

- [Create Dedicated Circuits](create-dedicated-circuits.md)
- [Move Selected Circuits](move-selected-circuits.md)
