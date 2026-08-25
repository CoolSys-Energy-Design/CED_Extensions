---
id: ced-electools-let-there-be-light
doc_type: tool
title: Let There Be Light
summary: Imports lighting-fixture data from a CSV file and places mapped Revit family types at supplied coordinates.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Let There Be Light
navigation_path: []
status: production
audience: [electrical]
model_impact: creates-lighting-fixtures
keywords: [lighting, CSV, import, fixtures, coordinates]
aliases: []
last_verified: "2026-08-24"
---

# Let There Be Light

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| External lighting-layout data must be translated into placed Revit fixtures. | Creates lighting-fixture instances and applies mapped placement data. | A supported CSV file and loaded destination family types. |

## What it does

Reads a user-selected CSV, maps its fixture names to Revit family types, applies the selected coordinate divisor, and places fixtures at the supplied locations and rotations.

## When to use it

Use it when a lighting layout is prepared in another system and delivered as structured CSV data.

## Before you start

- Review the CSV source and confirm coordinate units, fixture names, levels, and rotation values.
- Load the destination Revit family types before starting.
- Test the mapping with a small representative source file when possible.

## Steps

1. On **AE pyTools > Electrical**, click **Let There Be Light**.
2. Select the source CSV file.
3. Map each unique fixture name to a Revit family type.
4. Enter the coordinate divisor required by the source units.
5. Run the import and review the placed fixtures.

## Results and verification

Mapped fixture instances are placed from the CSV data. Verify fixture type, level, location, orientation, quantity, and coordinate scaling before continuing with circuiting or documentation.

## Notes and limitations

> [!WARNING]
> This creates model elements. An incorrect family mapping or coordinate divisor can create a large number of incorrectly placed fixtures.

## Related pages

- [Create Dedicated Circuits](create-dedicated-circuits.md)
- [CED ElecTools](index.md)
