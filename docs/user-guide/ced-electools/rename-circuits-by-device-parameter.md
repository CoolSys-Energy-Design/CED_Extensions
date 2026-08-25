---
id: ced-electools-rename-circuits-by-device-parameter
doc_type: tool
title: Rename Circuits by Device Parameter
summary: Builds circuit names from connected-device parameters, custom text, separators, and saved templates.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Rename Circuits by Device Parameter
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: Changes circuit names on selected circuits.
keywords: [electrical, circuits, rename, parameters, naming, templates]
aliases: [Circuit Name Builder]
last_verified: "2026-08-24"
---

# Rename Circuits by Device Parameter

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Circuit names must be consistently derived from their connected devices. | Writes new circuit names. | Selected devices or circuits with usable parameter data. |

## What it does

Finds connected circuits from the current selection and provides a string-builder UI for combining parameter tokens, custom text, and separators into names.

## When to use it

Use it to apply a repeatable circuit naming rule across a coordinated set of circuits.

## Before you start

- Select target devices or circuits.
- Check that the needed device parameters have the expected values.
- Decide whether duplicate incrementing should be used.

## Steps

1. Run **Rename Circuits by Device Parameter**.
2. Build a name using parameter tokens, custom fields, and separators.
3. Preview the result on selected or all listed circuits.
4. Apply and read the results table.

## Results and verification

Verify the previous and new name for every processed circuit in the output and confirm names in the model.

## Notes and limitations

- Builder templates are stored in the current user's pyRevit configuration.
- For tied parameter values on a circuit, the most common value is used; an Element ID tie uses the lowest ID.
- Circuits without usable source values can be skipped or remain unchanged.

## Related pages

- [Edit Circuit Properties](edit-circuit-properties.md)
- [Circuit Manager](circuit-manager.md)
