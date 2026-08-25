---
id: ced-mechtools-space-coils
doc_type: tool
title: Space Coils
summary: Evenly distributes selected coil families along a chosen wall direction.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Ref Ops > Space Coils
navigation_path: [Refrigeration]
status: beta
audience: [refrigeration]
model_impact: moves-and-orients-coil-instances
keywords: [refrigeration, coils, spacing, walls, orientation]
aliases: []
last_verified: "2026-08-24"
---

# Space Coils

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| A row of selected refrigeration coils needs consistent spacing along a wall. | Moves and orients selected coil family instances. | Select the coils to distribute and choose the wall direction. |

## What it does

Distributes selected coil families with equal gaps along the chosen wall direction, positions them 15 inches off the wall, and faces them opposite that direction.

## When to use it

Use it after coil placement when a line of cases or coils needs consistent layout spacing.

## Before you start

- Select only the coil families that belong in one spacing operation.
- Confirm the intended wall direction and required clearances.

## Steps

1. Select the coils to distribute.
2. On **AE pyTools > Mechanical > Ref Ops**, click **Space Coils**.
3. Choose the wall direction.
4. Review the resulting spacing and coil orientation.

## Results and verification

Selected coils are moved and oriented according to the chosen direction. Verify equal gaps, wall offset, facing direction, clearances, and connected piping.

## Notes and limitations

> [!WARNING]
> This beta command moves selected model elements. Use a focused selection and verify the layout before reconnecting or documenting the system.

## Related pages

- [Place All Coils](place-all-coils.md)
- [CED MechTools](index.md)
