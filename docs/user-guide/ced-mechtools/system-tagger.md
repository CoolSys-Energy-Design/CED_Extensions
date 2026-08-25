---
id: ced-mechtools-system-tagger
doc_type: tool
title: System Tagger
summary: Assigns refrigeration System IDs to selected cases and optionally creates text tags and Identity Mark values.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Ref Ops > System Tagger
navigation_path: [Refrigeration]
status: beta
audience: [refrigeration]
model_impact: updates-equipment-data-and-creates-tags
keywords: [refrigeration, System ID, cases, tags, Identity Mark, Excel]
aliases: []
last_verified: "2026-08-24"
---

# System Tagger

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Refrigerated cases need System IDs assigned from a controlled list. | Can update Identity Mark values and create text-note tags on selected mechanical equipment. | A System ID list and refrigerated cases in an appropriate active view. |

## What it does

Accepts System IDs pasted from Excel, lets you assign each ID to one or more refrigerated cases, and can tag the cases, update Identity Mark, or do either action alone.

## When to use it

Use it to apply consistent case System IDs after equipment placement and before refrigeration deliverables are annotated.

## Before you start

- Prepare the System ID list from the controlled Excel source.
- Open the target view and confirm case selection/visibility.
- Decide whether the workflow should tag, update Identity Mark, or both.

## Steps

1. On **AE pyTools > Mechanical > Ref Ops**, click **System Tagger**.
2. Paste the System ID list from Excel.
3. For each System ID, click the refrigerated cases to assign; click again to unselect and press Esc when finished picking.
4. Continue through the IDs and complete the chosen action mode.
5. Review tags and equipment data in the model.

## Results and verification

Selected cases receive the assigned System ID action. When several cases share an ID, generated text labels can append `A`, `B`, `C`, and so on. Verify every case, tag, and Identity Mark value.

## Notes and limitations

> [!WARNING]
> This beta workflow writes project data and can create tags. Confirm the ID list before applying it to a large set of cases.

## Related pages

- [Name Piping Systems](name-piping-systems.md)
- [Place All Coils](place-all-coils.md)
- [CED MechTools](index.md)
