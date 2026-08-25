---
id: wm-tools-toggle-connector-symbols
doc_type: tool
title: Toggle Connector Symbols
summary: Toggles the visibility setting for all configured refrigeration power-connector family types.
extension: WM Tools
ribbon_path: AE pyTools > WM Tools > Toggle Connector Symbols
navigation_path: []
status: production
audience: [electrical, refrigeration]
model_impact: Changes the `Symbol Visible_CED` type parameter on matching connector family types.
keywords: [wm, connectors, symbols, visibility, family types]
aliases: [Toggle Case Power Symbols]
last_verified: "2026-08-24"
---

# Toggle Connector Symbols

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Connector symbols need to be shown or hidden project-wide by type. | Edits a type parameter across matching family types. | The configured refrigeration power-connector family is loaded. |

## What it does

Reads the first matching connector type's `Symbol Visible_CED` setting and applies the opposite value to every type in the configured connector family.

## When to use it

Use it to switch the display state of case power symbols consistently across all connector types.

## Before you start

- Confirm the project uses `EF-U_Refrig Power Connector-Balanced_CED-WM`.
- Coordinate the visual change with the team because it updates type-level data.

## Steps

1. Run **Toggle Connector Symbols**.
2. Read the balloon notification confirming the new ON or OFF state and affected type count.

## Results and verification

Check a representative connector in an applicable view. The command changes the shared type parameter, not individual instances.

## Notes and limitations

- The command has no picker or confirmation dialog.
- It does nothing when the required family or editable parameter is unavailable.

## Related pages

- [Load Electrical Content](load-electrical-content.md)
- [Select Connectors By Case](select-connectors-by-case.md)
