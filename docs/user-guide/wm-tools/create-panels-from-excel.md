---
id: wm-tools-create-panels-from-excel
doc_type: tool
title: Create Panels From Excel
summary: Places panel and switchboard instances from rows in a Panel Creation workbook.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Create Panels From Excel
navigation_path: []
status: production
audience: [electrical]
model_impact: Creates panel or switchboard instances in the active project.
keywords: [wm, panels, switchboards, excel, placement]
aliases: [Panel Creation]
last_verified: "2026-08-24"
---

# Create Panels From Excel

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Panel locations and types have been prepared in the WM workbook. | Creates model instances. | An active plan view and a valid `Panel Creation` worksheet. |

## What it does

Reads panel and switchboard rows from the project workbook and places the matching family types at their specified locations. It asks you to resolve missing distribution systems or family types.

## When to use it

Use it after the standard electrical content is loaded and the panel-creation workbook has been checked.

## Before you start

- Use the project copy of the Panel Creation workbook.
- Open a plan view with an associated level.
- Confirm family/type names, distribution systems, and placement data.

## Steps

1. Run **Create Panels From Excel** and choose the workbook when prompted.
2. Resolve any missing distribution-system or family/type prompts.
3. Review the placement report after the transaction completes.

## Results and verification

Verify the panel count, locations, orientation, types, and assigned distribution systems against the workbook.

## Notes and limitations

- The command creates instances; it does not replace existing matching panels.
- It exits if an active view has no associated level or a required replacement is cancelled.
- The bundle tooltip supports Alt+click to locate a copy of the Panel Creation workbook.

## Related pages

- [Load Electrical Content](load-electrical-content.md)
- [Create Circuits From Excel](create-circuits-from-excel.md)
