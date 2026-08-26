---
id: ced-electools-add-spares-and-spaces
doc_type: tool
title: Add Spares/Spaces
summary: Batch adds or removes default SPARE and SPACE entries across selected panel schedules.
extension: CED ElecTools
ribbon_path: AE pyTools > Electrical > Circuit Tools > Add Spares/Spaces
navigation_path: [Circuit Tools]
status: production
audience: [electrical]
model_impact: Adds or removes spare and space circuit entries in selected panel schedules.
keywords: [electrical, panels, spares, spaces, panel schedule]
aliases: [Add Remove Spares and Spaces]
last_verified: "2026-08-24"
---

# Add Spares/Spaces

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| Selected panel schedules need standard SPARE or SPACE entries added or removed. | Changes circuit entries in panel schedules. | One or more panel schedule views in the model. |

## What it does

Opens a staged panel-schedule workflow for adding or removing default SPARE and SPACE entries. Quick actions support fill or remove-all behavior.

## When to use it

Use it when empty panel schedule slots need to be filled with spare/spaces before issuance.

>[!NOTE]
> The tool has a **Quick Apply mode**, where you can add spares or spaces directly while reviewing panel schedules, without selecting panels by name in the main editor.
>1. If the tool is launched from an active Panel Schedule view, Quick Apply targets that schedule.
>2. If Panel Schedule Sheet Instances are selected, Quick Apply targets the selected schedules.
>3. If neither applies, the tool opens the main editor, where you can view and select panels from the full panel list.


## Steps

1. Run **Add Spares/Spaces**.
2. Choose the panel schedules and the required mode, or use a quick action.
3. Review staged actions in the table.
4. Apply and read the completion result.

## Results and verification

Open each affected panel schedule and confirm the expected SPARE/SPACE rows, circuit numbers, and remaining capacity.

## Notes and limitations

- The command requires panel schedule views to be created for Electrical Equipment; it reports when none are present.
- The tool adds 1-pole spares/spaces to panelboards and 3-pole spares/spaces to switchboards.
- Added Spares/Spaces follow the project's default **Name** and **Rating** settings. If different names or ratings are required, they must be updated manually after creation.

## Related pages

- [Panel Report](panel-report.md)
- [Move Selected Circuits](move-selected-circuits.md)
