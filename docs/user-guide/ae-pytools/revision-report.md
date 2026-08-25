---
id: ae-pytools-revision-report
doc_type: tool
title: Revision Report
summary: Creates a sheet-sorted revision narrative from selected revisions and their cloud comments.
extension: AE pyTools
ribbon_path: AE pyTools > CED Tools > Revisions > Revision Report
navigation_path: [Revisions]
status: production
audience: [all]
model_impact: report-only
keywords: [revisions, revision clouds, sheets, comments, report]
aliases: []
last_verified: "2026-08-24"
---

# Revision Report

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| A formal, sheet-sorted revision narrative is needed. | None - creates a pyRevit output report only. | A project with revisions and revision-cloud comments. |

## What it does

Prompts for one or more revisions, then lists their revision-cloud comments by sheet number in a formatted output report.

## When to use it

Use it to prepare a revision summary before issuing sheets or reviewing a revision package.

## Before you start

- Confirm revision clouds have accurate **Comments** values.
- Use Shift-click only when additional revision-cloud parameters should be included in future reports.

## Steps

1. On **AE pyTools > CED Tools > Revisions**, click **Revision Report**.
2. Select the revisions to include.
3. Review the generated report in the pyRevit output window.

## Results and verification

The report lists revision information and deduplicated cloud comments by sheet. Verify missing or unexpected comments against the sheets before issuing documents.

## Notes and limitations

- The default report columns are Sheet Number, Sheet Name, and Comments.
- Shift-click can save additional cloud-parameter columns in the user configuration; unavailable saved parameters are reset to the defaults.

## Related pages

- [Revision Cloud Issues Report](revision-cloud-issues-report.md)
- [Revision Tools](revisions.md)
