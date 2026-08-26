---
id: ced-mechtools-three-d-rotate
doc_type: tool
title: 3D Rotate
summary: Rotates selected linear elements around their own axes or a user-selected axis in three dimensions.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Utilities > 3D Rotate
navigation_path: [Utilities]
status: beta
audience: [mechanical]
model_impact: rotates-selected-elements
keywords: [3D rotate, axis, piping, ductwork, rotation]
aliases: [Element 3D Rotation]
last_verified: "2026-08-24"
---

# 3D Rotate

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Linear elements need an axis-based rotation that cannot be achieved in a plan view. | Rotates selected elements in 3D. | Select eligible linear elements and confirm the intended rotation axis and angle. |

## What it does

Provides an options window for rotating selected eligible elements around their own axes or around a chosen linear axis.

## When to use it

Use it for controlled 3D alignment adjustments to linear mechanical elements.

## Before you start

- Select only the linear elements to rotate.
- Confirm worksharing ownership and dependent connections.

## Steps

1. Select the target elements.
2. On **AE pyTools > Mechanical > Utilities**, click **3D Rotate**.
3. Choose the rotation mode, axis, and angle in the options window.
4. Apply the rotation and inspect the result in 3D and plan views.

## Results and verification

Selected elements rotate about the chosen axis. Verify alignment, connector continuity, clearances, and associated fittings.

## Notes and limitations

> [!WARNING]
> This beta command changes 3D geometry. Test on a small selection and verify connected networks before wider use.

## Related pages

- [Make Parallel](make-parallel.md)
- [Transition](transition.md)
- [CED MechTools](index.md)
