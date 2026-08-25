---
id: ae-pytools-gray-x-bg-cad-layers
doc_type: tool
title: Gray X_BG CAD Layers
summary: Sets projection-line overrides to gray for every layer of the linked X_BG CAD file in views and templates.
extension: AE pyTools
ribbon_path: AE pyTools > Miscellaneous > Views > XBG Grey All Layers
navigation_path: [Miscellaneous]
status: beta
audience: [all]
model_impact: Changes graphic overrides for X_BG CAD subcategories across applicable views and view templates.
keywords: [cad, x_bg, layers, gray, graphics, view templates]
aliases: [XBG Grey All Layers]
last_verified: "2026-08-24"
---

# Gray X_BG CAD Layers

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| The linked X_BG CAD background should use gray linework everywhere. | Edits graphic overrides across views and templates. | A linked CAD instance with `X_BG` in its category name. |

## What it does

Finds the linked X_BG CAD category and changes every subcategory's projection line color to gray in applicable project views and view templates.

## When to use it

Use it when a project background CAD needs a consistent subdued display.

## Before you start

- Confirm the correct linked CAD category includes `X_BG` in its name.
- Save the project; the command affects many views and templates.

## Steps

1. Run **XBG Grey All Layers**.
2. Read the completion message for the number of views/templates and layers updated.
3. Check representative views and templates.

## Results and verification

Verify X_BG layer linework is gray in the intended views and view templates without disturbing other CAD links.

## Notes and limitations

- It skips schedule and panel-schedule views.
- The command does not provide a category picker or undo preview.

## Related pages

- [Copy V/G Settings to View Templates](copy-vg-settings-to-view-templates.md)
