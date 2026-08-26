---
id: ced-mechtools-print-pipe-data
doc_type: tool
title: Print Pipe Data
summary: Summarizes PipeSegment totals by System ID for selected worksets and exports the results to Excel.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Ref Ops > Print Pipe Data
navigation_path: [Refrigeration]
status: beta
audience: [refrigeration]
model_impact: exports-report-only
keywords: [refrigeration, piping, worksets, PipeSegment, System ID, Excel]
aliases: []
last_verified: "2026-08-24"
---

# Print Pipe Data

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Refrigeration pipe quantities need an Excel summary by System ID. | Does not edit the model; creates an Excel export. | Select the worksets and pipe-system types to include. |

## What it does

Collects pipe elements from selected worksets, summarizes PipeSegment totals by System ID, and exports the result to Excel.

## When to use it

Use it for refrigeration piping QA, quantity review, or controlled external reporting.

## Before you start

- Confirm pipe Identity Mark/System ID data is current.
- Confirm the intended worksets and pipe-system types.

## Steps

1. On **AE pyTools > Mechanical > Ref Ops**, click **Print Pipe Data**.
2. Select the worksets and pipe-system types to collect.
3. Run the collection and choose the Excel export location when prompted.
4. Review the exported totals.

## Results and verification

An Excel file is produced with PipeSegment totals grouped by System ID. Confirm included worksets, system IDs, pipe types, and totals against the model.

## Notes and limitations

- The report is only as reliable as the pipe-system and Identity Mark data in the model.
- This beta command exports data but does not correct missing or inconsistent System IDs.

## Related pages

- [Name Piping Systems](name-piping-systems.md)
- [System Tagger](system-tagger.md)
- [CED MechTools](index.md)
