---
id: ae-pytools-tag-behavior-settings
doc_type: guide
title: Tag Behavior Settings
summary: Control how related tag heads move, rotate, and retain Model orientation during normal Orientation commands.
extension: AE pyTools
ribbon_path: AE pyTools > Orientation
navigation_path: [Orientation]
status: production
audience: [all]
model_impact: saved-user-preference
keywords: [tags, position, rotation, model orientation, smart button]
aliases: [Tag Position, Tag Rotation, Model Orientation, Tag Model Orientation]
last_verified: "2026-08-24"
---

# Tag Behavior Settings

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Related tags need predictable behavior when their hosts are oriented or rotated. | Changes saved user preferences only; toggling a setting does not edit the model. | No selection is required. |

## What it does

The three smart buttons control related-tag behavior for normal clicks of the commands in [Orient and Rotate Elements](orient-and-rotate-elements.md).

## When to use it

Set the options before rotating tagged component families when tag placement or reading direction must be preserved.

## Before you start

- No selection is required to change a setting.
- Confirm the current smart-button icons before running an Orientation command.

## Settings

| Setting | When on | Dependency |
|---|---|---|
| **Tag Position** | Moves an eligible tag head to preserve its relative location to the host. | None. |
| **Tag Rotation** | Rotates an eligible tag head to preserve its relative angle to the host. | Requires Tag Position. |
| **Model Orientation** | Keeps eligible related tags in Revit's **Model** orientation. | Requires Tag Position. |

## Steps

1. On **AE pyTools > Orientation**, click a smart button to toggle the desired setting.
2. Confirm its icon shows the intended state.
3. Select the component families to change.
4. Run an Orientation command with a normal click.
5. Review tag placement and readability after the model change.

## Results and verification

The setting is saved for the user and its button icon updates. It affects later normal Orientation command clicks; verify tag placement after each model-changing command.

## Notes and limitations

- Tag Rotation and Model Orientation do nothing while Tag Position is off.
- Shift-clicking an Orientation command temporarily leaves related tags in place and ignores the saved settings for that run.

## Related pages

- [Orient and Rotate Elements](orient-and-rotate-elements.md)
- [Orientation Tools](orientation.md)
