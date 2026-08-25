---
id: ced-mechtools-place-all-coils
doc_type: tool
title: Place All Coils
summary: Places coil family instances in matching Spaces or Rooms from an Excel source file.
extension: CED MechTools
ribbon_path: AE pyTools > Mechanical > Ref Ops > Place All Coils
navigation_path: [Refrigeration]
status: beta
audience: [refrigeration]
model_impact: creates-coil-family-instances
keywords: [refrigeration, coils, Excel, rooms, spaces, mechanical equipment]
aliases: []
last_verified: "2026-08-24"
---

# Place All Coils

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Refrigeration coils need to be placed from a coordinated Excel schedule. | Creates coil family instances in matched Spaces or Rooms. | An Excel file with Description, Coil Count, and Model # data, plus loaded coil families. |

## What it does

Matches Excel descriptions to Space or Room names, selects the closest mechanical-equipment family type to each model number, and places the requested coils with a vertical offset.

## When to use it

Use it to accelerate initial coil placement from a controlled equipment schedule.

## Before you start

- Validate the Excel descriptions, coil counts, and model numbers.
- Confirm the destination Spaces or Rooms exist and coil family types are loaded.
- Test a small representative import when possible.

## Steps

1. On **AE pyTools > Mechanical > Ref Ops**, click **Place All Coils**.
2. Select the Excel source file.
3. Review proposed type matches and placement results.
4. Inspect placed coils in each affected Space or Room.

## Results and verification

Coil family instances are created from the Excel data and offset 2 feet vertically. Verify quantities, family types, locations, levels, and offsets.

## Notes and limitations

> [!WARNING]
> This beta command creates model elements. Description or model-number mismatches can create incorrect placements or types.

## Related pages

- [Space Coils](space-coils.md)
- [System Tagger](system-tagger.md)
- [CED MechTools](index.md)
