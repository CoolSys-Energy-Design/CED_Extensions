---
id: wm-tools-ungroup-power-groups
doc_type: tool
title: Ungroup Power Groups
summary: Ungroups qualifying power groups and propagates circuit and system information to their contents.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Ungroup Power Groups
navigation_path: []
status: production
audience: [electrical, refrigeration]
model_impact: Ungroups model groups and writes circuit and system parameters to resulting elements.
keywords: [wm, ungroup, power groups, refrigeration circuit number, system number]
aliases: [Explode Power Groups]
last_verified: "2026-08-24"
---

# Ungroup Power Groups

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Placed power groups must become independent elements while retaining their data. | Ungroups instances and propagates parameter values. | Selected groups, or qualifying groups visible in the active view. |

## What it does

Finds model groups with a populated `Refrigeration Circuit Number_CEDT` parameter, ungroups them, and propagates refrigeration circuit and system data to their elements.

## When to use it

Use it only when group instances are no longer needed as editable group instances.

## Before you start

- Select the intended power groups, or isolate the correct groups in the active view.
- Verify their refrigeration circuit number values.
- Save the project because ungrouping is not a presentation-only change.

## Steps

1. Select target groups if you do not want the command to scan the active view.
2. Run **Ungroup Power Groups**.
3. Review the resulting independent elements and their propagated data.

## Results and verification

Confirm the groups have been removed and resulting elements carry the correct `Refrigeration Circuit Number_CEDT` and `System Number_CEDT` values.

## Notes and limitations

- With no selection, the command scans model groups visible in the active view.
- Groups without a populated refrigeration circuit number are skipped.

## Related pages

- [Place Power Groups](place-power-groups.md)
