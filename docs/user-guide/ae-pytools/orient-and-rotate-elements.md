---
id: ae-pytools-orient-and-rotate-elements
doc_type: guide
title: Orient and Rotate Elements
summary: Set selected component families to a plan direction or rotate them by a quarter turn.
extension: AE pyTools
ribbon_path: AE pyTools > Orientation
navigation_path: [Orientation]
status: production
audience: [all]
model_impact: modifies-selected-elements
keywords: [orientation, rotation, north, south, east, west, clockwise, counterclockwise, tags]
aliases: [Orient Up, Orient Left, Orient Down, Orient Right, Rotate CW, Rotate CCW]
last_verified: "2026-08-24"
---

# Orient and Rotate Elements

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Component families need a standard plan direction or a 90-degree rotation. | Rotates eligible selected family instances; related tag behavior follows the saved Orientation settings. | Select one or more unpinned component family instances. |

## What it does

The six Orientation commands preserve the selected element locations while either setting a plan-facing direction or applying a 90-degree rotation. The rotation commands support eligible component families connected with wires.

## When to use it

Use these commands to standardize the direction of equipment, fixtures, and other placed component families without manually rotating each item.

## Before you start

- Select the component family instances to change.
- Review [Tag Behavior Settings](tag-behavior-settings.md) when selected elements have related tags.

## Commands

| Command | Result |
|---|---|
| **Orient Up** | Faces selected elements plan north. |
| **Orient Left** | Faces selected elements plan west. |
| **Orient Down** | Faces selected elements plan south. |
| **Orient Right** | Faces selected elements plan east. |
| **Rotate CCW** | Rotates selected elements 90 degrees counterclockwise. |
| **Rotate CW** | Rotates selected elements 90 degrees clockwise. |

## Steps

1. Select the required component family instances.
2. On **AE pyTools > Orientation**, choose the command that gives the required direction or rotation.
3. Review the rotated elements and related tags in the active view.

## Results and verification

Eligible selected instances change orientation and remain at their original locations. Verify their final direction, wire and host relationships, and any related tag placement.

## Notes and limitations

- Pinned elements, annotations, system families, and face-based elements hosted on vertical walls are ignored.
- Shift-click any Orientation command to temporarily leave related tags in place, regardless of the saved Orientation settings.

> [!WARNING]
> These commands modify the model. Check the result before continuing with tagging, wiring, or other dependent work.

## Related pages

- [Tag Behavior Settings](tag-behavior-settings.md)
- [Orientation Tools](orientation.md)
