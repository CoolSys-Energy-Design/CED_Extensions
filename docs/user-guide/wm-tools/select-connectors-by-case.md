---
id: wm-tools-select-connectors-by-case
doc_type: tool
title: Select Connectors By Case
summary: Selects refrigeration connector instances by circuit number and connector type.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Select Connectors By Case
navigation_path: []
status: production
audience: [electrical, refrigeration]
model_impact: none
keywords: [wm, connectors, cases, refrigeration circuit number, selection]
aliases: [Select Case Connectors]
last_verified: "2026-08-24"
---

# Select Connectors By Case

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need the connector instances for one refrigeration circuit and load type. | Changes the current selection only. | Connector instances with `Refrigeration Circuit Number_CEDT` values. |

## What it does

Groups eligible connector instances by refrigeration circuit number, then lets you choose a circuit and one or more connector types to select.

## When to use it

Use it to prepare a focused connector selection before editing case data or creating circuits.

## Before you start

- Confirm target connector instances carry the expected refrigeration circuit number.
- Ensure connector family types use meaningful type names.

## Steps

1. Run **Select Connectors By Case**.
2. Choose the refrigeration circuit number.
3. Choose the connector types to include.

## Results and verification

The matching connector instances become the active Revit selection. Verify the count and selected types before performing another command.

## Notes and limitations

- The command does not modify the model.
- It only finds matching family instances with the expected circuit-number parameter.

## Related pages

- [Replace Existing Circuit](replace-existing-circuit.md)
- [Place Power Groups](place-power-groups.md)
