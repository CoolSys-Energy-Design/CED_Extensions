---
id: ae-pytools-copy-vg-settings-to-view-templates
doc_type: tool
title: Copy V/G Settings to View Templates
summary: Copies selected category visibility and graphic overrides from the active view to selected view templates.
extension: AE pyTools
ribbon_path: AE pyTools > Miscellaneous > Views > Copy V/G Settings to View
navigation_path: [Miscellaneous]
status: production
audience: [all]
model_impact: Changes selected view templates' category visibility and graphic overrides.
keywords: [view graphics, visibility, overrides, view templates, categories]
aliases: [Copy VG Settings]
last_verified: "2026-08-24"
---

# Copy V/G Settings to View Templates

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Selected category visibility and overrides must be replicated to templates. | Edits selected view templates. | A supported active view with the source V/G configuration. |

## What it does

Collects the selected categories' visibility and graphic overrides, including subcategories, from the active view and applies them to selected compatible view templates.

## When to use it

Use it after refining V/G settings in a representative view and before applying the same settings to templates.

## Before you start

- Open a floor plan, RCP, drafting, 3D, section, or elevation view with the desired source V/G settings.

## Steps

1. Run **Copy V/G Settings to View Templates**.
2. Select source categories to copy.
3. Select one or more target view templates to apply settings.
4. Confirm completion and review the templates.

## Results and verification

Open a target template or controlled view and verify category overrides and hidden state, including relevant subcategories.

## Notes and limitations

- Schedules are not offered as target templates.
- Unsupported active view types are blocked before any change.
- The command copies selected category settings, not all V/G settings.

## Related pages

- [Gray X_BG CAD Layers](gray-x-bg-cad-layers.md)
