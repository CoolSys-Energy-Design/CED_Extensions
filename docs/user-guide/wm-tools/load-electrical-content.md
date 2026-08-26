---
id: wm-tools-load-electrical-content
doc_type: tool
title: Load Electrical Content
summary: Loads CED electrical content, shared parameters, schedules, filters, templates, and starter-project assets.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Load Electrical Content
navigation_path: []
status: production
audience: [electrical, refrigeration]
model_impact: Adds or updates project parameter bindings and loads or copies CED content into the active project.
keywords: [wm, electrical content, shared parameters, families, schedules, templates, filters]
aliases: [Load Content]
last_verified: "2026-08-24"
---

# Load Electrical Content

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| A WM electrical project needs its standard CED content and setup. | Loads content and changes project parameter bindings. | An open project and the synced CED content collection. |

## What it does

Loads the configured shared parameters and available CED families, schedules, panel templates, filters, groups, and related starter-project assets. Existing items with matching names are generally left in place.

## When to use it

Use it at project setup before placing WM electrical content or running the other WM workflows.

## Before you start

- Sync the required CED content collection through Desktop Connector.
- Save or coordinate the project before changing shared-parameter bindings.
- Close views that could prevent copied schedules or content from being changed.

## Steps

1. Open the target project and run **Load Electrical Content**.
2. Confirm the sync reminder, then let the tool complete its output report.
3. Review skipped or failed items in the pyRevit output.

## Results and verification

Verify that required families, shared parameters, panel templates, schedules, and filters are available before using WM placement or circuit tools.

## Notes and limitations

- The command depends on a locally synced content collection and starter project.
- It can update existing shared-parameter bindings, so coordinate project standards first.
- Missing source content or inaccessible Desktop Connector paths stop the workflow.

## Related pages

- [Create Panels From Excel](create-panels-from-excel.md)
- [Create Circuits From Excel](create-circuits-from-excel.md)
