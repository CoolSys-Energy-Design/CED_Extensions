---
id: wm-tools-replace-existing-circuit
doc_type: tool
title: Replace Existing Circuit
summary: Replaces a selected panel-schedule placeholder circuit with compatible connector instances.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Replace Existing Circuit
navigation_path: []
status: production
audience: [electrical, refrigeration]
model_impact: Changes circuit membership and removes the selected placeholder family instance.
keywords: [wm, circuit, placeholder, panel schedule, connectors]
aliases: [Replace Placeholder Circuit]
last_verified: "2026-08-24"
---

# Replace Existing Circuit

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| A panel schedule contains a placeholder that must be replaced by actual connectors. | Adds selected connectors to a circuit and deletes its placeholder. | A panel schedule view with exactly one placeholder circuit selected. |

## What it does

Finds connector instances matching the selected circuit's voltage and poles, lets you choose connector types, adds them to the circuit, and removes the placeholder.

## When to use it

Use it once the real connectors are in the model and a placeholder circuit needs conversion.

## Before you start

- Open a panel schedule view.
- Select exactly one circuit cell containing a recognized placeholder family instance.
- Verify the target connectors have `Refrigeration Circuit Number_CEDT` values and compatible electrical data.

## Steps

1. Select the placeholder circuit in the panel schedule.
2. Run **Replace Existing Circuit**.
3. Choose the matching refrigeration circuit and connector types.
4. Review the output confirmation.

## Results and verification

Verify the intended connector instances are assigned to the circuit and that the placeholder is gone. Review the circuit's panel, number, load name, and connector count in the output.

## Notes and limitations

- The command works only from a panel schedule view.
- It filters candidates by voltage and poles.
- Cancel if the placeholder or candidate list is not the intended circuit.

## Related pages

- [Select Connectors By Case](select-connectors-by-case.md)
- [Create Circuits From Excel](create-circuits-from-excel.md)
