---
id: ced-mechtools-make-parallel
doc_type: tool
title: Make Parallel
summary: Rotates a selected target element to run parallel to a selected reference element in the XY plane.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Utilities > Make Parallel
navigation_path: [Utilities]
status: beta
audience: [mechanical]
model_impact: rotates-selected-elements
keywords: [parallel, XY, rotation, piping, ductwork]
aliases: [Make parallel (XY)]
last_verified: "2026-08-24"
---

# Make Parallel

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| A linear element should match the plan direction of another element. | Rotates the target element in the XY plane. | Pick a reference linear element, then the target element. |

## What it does

Uses the first picked element as the reference direction and rotates the second picked element to be parallel to it in the XY plane.

## When to use it

Use it to align parallel runs without manually calculating the rotation angle.

## Before you start

- Identify the correct reference and target elements.
- Review connected fittings and clearances around the target element.

## Steps

1. On **AE pyTools > Mechanical > Utilities**, click **Make Parallel**.
2. Pick the reference element.
3. Pick the target element to rotate.
4. Inspect the target's new direction and connections.

## Results and verification

The target rotates to run parallel with the reference. Verify its orientation, position, and connected geometry.

## Notes and limitations

- This beta tool is intended for XY-plane alignment.
- The first selected element is the reference; the second is the element that changes.

## Related pages

- [3D Rotate](three-d-rotate.md)
- [Transition](transition.md)
- [CED MechTools](index.md)
