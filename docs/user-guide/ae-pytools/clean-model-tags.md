---
id: ae-pytools-clean-model-tags
doc_type: tool
title: Clean Model Tags
summary: Converts orthogonal model-oriented tags to standard horizontal or vertical tag orientation.
extension: AE pyTools
ribbon_path: AE pyTools > Tags > Clean Model Tags
navigation_path: [Tags]
status: production
audience: [all]
model_impact: modifies-tags
keywords: [tags, model orientation, horizontal, vertical]
aliases: []
last_verified: "2026-08-24"
---

# Clean Model Tags

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Orthogonal tags using Model orientation need to behave as horizontal or vertical tags. | Changes eligible tag orientation. | Select tags, or use the active view when no tags are selected. |

## What it does

Converts tags at orthogonal Model-orientation angles - 0, 90, 180, or 270 degrees - to Revit's horizontal or vertical orientation.

## When to use it

Use it to make tags easier to rotate and standardize after previous model-orientation workflows.

## Before you start

- Select only the tags to change when a targeted result is needed.

## Steps

1. Optionally select tags in the active view.
2. On **AE pyTools > Tags**, click **Clean Model Tags**.
3. Review the changed tag directions.

## Results and verification

Eligible tags are converted to horizontal or vertical orientation. Verify tag locations and readability in the active view.

## Notes and limitations

- With no selection, the command processes tags in the active view.
- Tag-family geometry can make the result appear different than expected; use the command selectively.

## Related pages

- [Navigate Tags and Hosts](navigate-tags-and-hosts.md)
- [Tag Tools](tags.md)
