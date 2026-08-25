---
id: ced-electools-color-circuits-check
doc_type: tool
title: Color Circuits Check
summary: Toggles active-view filters that color circuited fixtures green and uncircuitied fixtures red.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > QC Check > Color Circuits Check
navigation_path: [QC Check]
status: production
audience: [electrical]
model_impact: Creates or updates view filters and toggles their visibility in the active view.
keywords: [electrical, circuits, fixtures, qc, filters, color]
aliases: [Fixture Check]
last_verified: "2026-08-24"
---

# Color Circuits Check

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need a quick visual check of circuited versus uncircuitied fixtures. | Creates/updates and toggles active-view filters. | An active view with electrical fixture categories and circuit-related parameters. |

## What it does

Creates or updates `Uncircuited Fixtures` and `Circuited Fixtures` filters, then toggles them on or off. When on, uncircuitied items are red and circuited items are green.

## When to use it

Use it for a fast active-view QA pass before running reports or issuing electrical drawings.

## Before you start

- Open the intended coordination view.
- Confirm the model contains electrical fixtures and applicable circuit-number, panel, or system-name parameters.

## Steps

1. Run **Color Circuits Check**.
2. Read the confirmation: the first applicable run turns the check on; the next toggles it off.
3. Review the red and green results in the active view.

## Results and verification

When enabled, verify that red fixtures are truly uncircuitied and green fixtures have a circuit assignment. Run again to disable the filters when finished.

## Notes and limitations

- The command can enable temporary view-properties mode when the active view supports it.
- It works from the active view and does not change circuit membership.

## Related pages

- [Color Circuits by Panel](color-circuits-by-panel.md)
- [Electrical System Check](electrical-system-check.md)
