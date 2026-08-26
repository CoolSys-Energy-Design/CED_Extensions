---
id: ae-pytools-refrigeration-plan-tag-text
doc_type: tool
title: Refrigeration Plan Tag Text
summary: Writes formatted panel and circuit text to qualifying refrigeration-plan electrical fixtures.
extension: AE pyTools
ribbon_path: AE pyTools > Miscellaneous > Misc Tools > Refr Tag TJ-SATX
navigation_path: [Miscellaneous]
status: production
audience: [electrical, refrigeration]
model_impact: Writes the `Tag_Text` parameter on qualifying electrical fixture instances.
keywords: [refrigeration, tag text, junction box, panel, circuits, defrost]
aliases: [Refr Tag TJ-SATX]
last_verified: "2026-08-24"
---

# Refrigeration Plan Tag Text

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Refrigeration-plan fixture tags need formatted panel and circuit data. | Writes `Tag_Text` on qualifying instances. | Electrical fixtures whose type name is `REFRIGERATION PLAN`. |

## What it does

Finds electrical fixtures with type name `REFRIGERATION PLAN`, reads their electrical systems, and writes formatted multi-line panel/circuit text to `Tag_Text`.

## When to use it

Use it after refrigeration-plan fixtures have assigned electrical systems and before issuing plans that display their tag text.

## Before you start

- Confirm the intended fixtures use the exact `REFRIGERATION PLAN` type name.
- Verify panel names and circuit numbers are assigned.
- Save the project because this updates parameters in batch.

## Steps

1. Run **Refr Tag TJ-SATX**.
2. Review the pyRevit output count and any skipped fixtures.
3. Check a representative refrigeration-plan tag.

## Results and verification

Verify each tag's panel/circuit lines and classifications. The primary circuit is classified as `FAN+LTS+WMR` when a defrost circuit also exists, otherwise `FAN+LTS`; remaining circuits are `DEFROST`.

## Notes and limitations

- Fixtures without electrical systems are skipped.
- The current implementation writes `Tag_Text`; its documented fixture-type write is disabled in code.

## Related pages

- [System Tagger](../ced-mechtools/system-tagger.md)
