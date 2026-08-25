---
id: ae-pytools-edit-schedule-headers
doc_type: tool
title: Edit Schedule Headers
summary: Standardizes schedule column headings for cleaner sheet presentation.
extension: AE pyTools
ribbon_path: AE pyTools > Miscellaneous > MiscTools > Edit Schedule Headers
navigation_path: [Miscellaneous]
status: production
audience: [all]
model_impact: modifies-schedule-headers
keywords: [schedules, headers, sheet presentation, CED parameters]
aliases: []
last_verified: "2026-08-24"
---

# Edit Schedule Headers

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Schedule column labels need consistent publishing names. | Updates column headings in selected schedules. | Select schedule instances on sheets or choose project schedules when prompted. |

## What it does

Converts schedule headings to uppercase and removes CED-style parameter prefixes and suffixes, including `CKT_`, `_CED`, `_CEDT`, and `_CEDR`.

## When to use it

Use it before placing schedules on sheets or publishing a schedule package.

## Before you start

- Review the schedule headings that should be changed.
- Select placed schedule instances for a focused run, or be ready to select schedules from the prompt.

## Steps

1. Select schedule instances on sheets, if applicable.
2. On **AE pyTools > Miscellaneous > MiscTools**, click **Edit Schedule Headers**.
3. Select schedules when prompted and review the output.

## Results and verification

The command updates eligible schedule column headings. Open each affected schedule and confirm the published labels are correct.

## Notes and limitations

- This is a model change; use it only on schedules whose headings should follow the CED naming convention.
- Custom headings that do not use those prefixes or suffixes can still be converted to uppercase.

## Related pages

- [Toggle Grid Bubbles](toggle-grid-bubbles.md)
- [Miscellaneous Tools](miscellaneous.md)
