---
id: ced-electools-edit-circuit-properties
doc_type: tool
title: Edit Circuit Properties
summary: Stages and applies property edits and recalculation for selected electrical circuits.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Edit Circuit Properties
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: Edits circuit properties and can recalculate affected circuits.
keywords: [electrical, circuits, properties, editor, recalculate]
aliases: [Circuit Property Editor]
last_verified: "2026-08-24"
---

# Edit Circuit Properties

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Several circuit values need review and controlled batch editing. | Updates selected circuit properties and recalculates them. | Selected circuits, or circuits selected in the tool. |

## What it does

Opens a property editor that lets you stage edits per circuit, preview recalculated values, and apply the staged changes together.

## When to use it

Use it for coordinated circuit-property corrections where reviewing staged changes is preferable to direct ad hoc edits.

## Before you start

- Select the intended circuits, or plan to choose them in the command.
- Confirm the desired property values and save the model.

## Steps

1. Run **Edit Circuit Properties**.
2. Select circuits if none were preselected.
3. Stage and review edits in the editor.
4. Apply changes and review the edited and recalculated counts.

> [!TIP]
> Edit Circuit Properties is also executable through the **Circuit Manager** interface. You can either select circuits in the manager and right click, or check circuits and execute from the **Actions** menu.

## Results and verification

Check each edited circuit's values and calculated results in Revit. The confirmation reports how many circuits were edited and recalculated.

## Notes and limitations

- The command exits when no valid circuit updates are staged.
- Missing editor resources or a closed project document prevent it from opening.

## Related pages

- [Circuit Manager](circuit-manager.md)
- [Calculate Circuits](calculate-circuits.md)
