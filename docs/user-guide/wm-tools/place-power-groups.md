---
id: wm-tools-place-power-groups
doc_type: tool
title: Place Power Groups
summary: Places a selected power group on chosen references and writes available circuit information to it.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Place Power Groups
navigation_path: []
status: production
audience: [electrical, refrigeration]
model_impact: Places model and optional detail groups and writes group parameter values.
keywords: [wm, power groups, ems tags, circuit number, refrigeration]
aliases: [Place Case Power Groups]
last_verified: "2026-08-24"
---

# Place Power Groups

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Case power-group graphics must be placed against reference elements. | Places groups and writes circuit data. | Selected references, or EMS case tags in the active view. |

## What it does

Lets you choose a model group and optional attached detail group, then places it against selected references. Without a selection, it uses EMS case tags in the active view.

## When to use it

Use it after case tags or other reference elements are positioned and their circuit information is ready.

## Before you start

- Select reference elements, or make sure the active view contains the intended EMS case tags.
- Confirm the chosen groups are correct for the project.
- Hold Shift when starting the command if an offset placement is needed.

## Steps

1. Run **Place Power Groups**.
2. Choose the model group and, if available, the attached detail group.
3. Complete any offset choice and review the placed groups.

## Results and verification

Verify group placement, orientation, and the `Refrigeration Circuit Number_CEDT` value. References without a usable `Circuit #` parameter receive the fallback text `ckt # not found`.

## Notes and limitations

- The command can target many element types, not only EMS tags.
- It changes the model by placing group instances.
- Review any fallback circuit text before issuing drawings.

## Related pages

- [Ungroup Power Groups](ungroup-power-groups.md)
- [Select Connectors By Case](select-connectors-by-case.md)
