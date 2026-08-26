---
id: ae-pytools-update-custom-dictionary
doc_type: tool
title: Update Custom Dictionary
summary: Copies the bundled custom spelling dictionary to installed Revit custom-dictionary locations.
extension: AE pyTools
ribbon_path: AE pyTools > CED Tools > Update Custom Dictionary
navigation_path: [CED Tools]
status: beta
audience: [all]
model_impact: none
keywords: [dictionary, spelling, revit, custom.dic, update]
aliases: [Custom Dic Update]
last_verified: "2026-08-24"
---

# Update Custom Dictionary

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Revit spelling should recognize the distributed CED custom terms. | None. Copies a local application-support file. | Permission to write installed Revit custom-dictionary folders. |

## What it does

Copies the command's `Custom.dic` file into existing Revit custom-dictionary locations for supported Revit years. The dictionary improves **Ideate Spell Check** results by adding commonly used AEC-industry abbreviations.

## When to use it

Use it after installing or updating the toolbar when custom spelling entries are missing.

## Steps

1. Run **Update Custom Dictionary**.
2. Read the pyRevit output for successful, skipped, and failed target locations.

## Results and verification

Restart or reopen the relevant Revit spelling workflow if necessary, then confirm a known custom term is recognized.

## Notes and limitations

- Missing Revit-version folders are skipped.
- Permission or file-access errors are listed in the output.
- The command does not change the open Revit model.

## Related pages

- [About](about.md)
