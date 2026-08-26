---
id: ae-pytools-navigate-tags-and-hosts
doc_type: guide
title: Navigate Tags and Hosts
summary: Switch the Revit selection between tags and their hosted elements in the active view.
extension: AE pyTools
ribbon_path: AE pyTools > Tags
navigation_path: [Tags]
status: production
audience: [all]
model_impact: selection-only
keywords: [tags, hosts, selection, navigation]
aliases: [Get Hosts From Tags, Get Tags From Hosts, Select Hosted Tags]
last_verified: "2026-08-24"
---

# Navigate Tags and Hosts

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need to move quickly between tags and the elements they reference. | Selection changes only. | Select tags or host elements in an active view. |

## What it does

**Get Hosts From Tags** replaces the selection with the hosts of selected tags. **Get Tags From Hosts** selects tags hosted by selected elements in the active view.

## When to use it

Use these commands to inspect model elements from their tags, or to adjust tags after selecting their hosts.

## Before you start

- Select valid tags for **Get Hosts From Tags**.
- Select host elements for **Get Tags From Hosts**.

## Steps

1. Make the starting selection.
2. On **AE pyTools > Tags**, click **Get Hosts From Tags** or **Get Tags From Hosts**.
3. Hold Shift to append the found elements to the selection instead of replacing it.

## Results and verification

The command changes the Revit selection only. Confirm that the resulting items are the intended tags or hosts.

## Notes and limitations

- **Get Tags From Hosts** searches the active view.
- Tags or hosts that cannot be resolved are not added to the result.

## Related pages

- [Clean Model Tags](clean-model-tags.md)
- [Tag Tools](tags.md)
