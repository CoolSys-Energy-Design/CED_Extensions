---
id: ced-electools-find-circuited-elements
doc_type: tool
title: Find Circuited Elements
summary: Selects electrical circuits, their connected elements, or both for model navigation and review.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Find Circuited Elements
navigation_path: []
status: production
audience: [electrical]
model_impact: selection-only
keywords: [circuits, devices, selection, navigation]
aliases: []
last_verified: "2026-08-24"
---

# Find Circuited Elements

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need to locate the devices associated with one or more circuits. | Selection changes only. | Choose circuits through the command or start from a relevant selection. |

## What it does

Lets you select electrical circuits, their connected elements, or both, then updates the Revit selection for review and navigation.

## When to use it

Use it to visually inspect the loads served by a circuit or to select circuit objects and devices together for QA.

## Before you start

- Open a view appropriate for reviewing the expected devices.

## Steps

1. On **AE pyTools > Electrical**, click **Find Circuited Elements**.
2. Choose the circuits to inspect.
3. Choose whether to select circuits, connected elements, or both.
4. Review the resulting Revit selection.

## Results and verification

The requested circuits and/or connected elements become selected. The model is unchanged.

## Notes and limitations

- Visibility in the active view can affect how easily selected elements can be reviewed.
- Use [Wire Tools](wire-tools.md) when the next task is wiring rather than selection.

## Related pages

- [Wire Tools](wire-tools.md)
- [Electrical System Check](electrical-system-check.md)
- [CED ElecTools](index.md)
