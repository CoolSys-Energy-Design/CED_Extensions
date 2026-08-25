---
id: ae-pytools-revision-cloud-issues-report
doc_type: tool
title: Revision Cloud Issues Report
summary: Reports revision clouds with missing comments or problematic sheet/view placement.
extension: AE pyTools
ribbon_path: AE pyTools > CED Tools > Revisions > Revision Cloud Issues Report
navigation_path: [Revisions]
status: production
audience: [all]
model_impact: report-only
keywords: [revisions, revision clouds, quality control, sheets, views]
aliases: [List Revision Cloud Issues]
last_verified: "2026-08-24"
---

# Revision Cloud Issues Report

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Revision clouds need a pre-issue quality-control review. | None - creates a pyRevit output report only. | A project with revision clouds. |

## What it does

Checks all revision clouds for blank comments, clouds placed in views instead of directly on sheets, and view-based clouds whose views are not placed on a sheet.

## When to use it

Run it before publishing sheets to identify cloud documentation and placement issues.

## Before you start

- Save or synchronize current sheet and revision-cloud work.

## Steps

1. On **AE pyTools > CED Tools > Revisions**, click **Revision Cloud Issues Report**.
2. Review the output table.
3. Use its linked revision-cloud IDs to locate and correct reported items.

## Results and verification

The report identifies each applicable issue with an `X` in its issue column. No model data is changed.

## Notes and limitations

- A cloud can appear in more than one issue category.
- A view-based cloud can be reported even if it is associated with a sheet; review the intended publishing workflow before changing it.

## Related pages

- [Revision Report](revision-report.md)
- [Revision Tools](revisions.md)
