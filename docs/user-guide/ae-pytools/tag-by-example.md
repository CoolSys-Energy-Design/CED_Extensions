---
id: ae-pytools-tag-by-example
doc_type: tool
title: Tag by Example
summary: Creates matching tags for compatible family instances using selected example tags and configurable placement behavior.
extension: AE pyTools
ribbon_path: AE pyTools > Tags > Tag by Example
navigation_path: [Tags]
status: production
audience: [all]
model_impact: Creates tag elements in the active view.
keywords: [tags, example, placement, leaders, rotation, family instances]
aliases: []
last_verified: "2026-08-24"
---

# Tag by Example

Tag by Example reproduces one or more reference tags on compatible components in the active view. It retains the reference tag type and can transfer host-relative placement, rotation behavior, and leaders.

> [!NOTE]
> This workflow is for single-reference `IndependentTag` tags hosted by loadable `FamilyInstance` components. It works in floor plans, reflected ceiling plans, and drafting views—not rooms, areas, spaces, walls, materials, linked-model objects, or other non-component hosts.

## Choose references and targets

Start with valid tags already selected in Revit, or use **Pick New Reference(s)** in the window. Every reference must be hosted by the same element. The tool creates one copy of each distinct reference tag type per target, so selecting the same tag type twice does not make duplicate tags.

The window identifies each reference's tag type, host category/family/type, owner view, count, leader state, and orientation. References can come from the active view or its associated primary/dependent view.

Choose how the target components are collected:

| Target mode | Target set | Appropriate use |
|---|---|---|
| **All visible elements of the same type** | Visible components with the reference host's exact Revit type. | Standardizing a repeated equipment or fixture type. |
| **All visible elements of the same family** | Visible components in the same family, including other types. | Applying a consistent convention across a family. |
| **All visible elements of the same category** | Visible components in the same category. | Broad category-level tagging. |
| **User selection** | A set you create with **New**, adjust with **Edit**, and inspect with **Preview Selection**. | Exception cases or an intentionally limited group. |

Automatic modes search only the active view. The original reference host, hidden components, unsupported hosts, and incompatible targets are excluded. **Include nested family instances** adds visible, independently referenceable nested components, but some nested families cannot be tagged reliably.

## Control placement and leaders

Placement follows the target host's local orientation and mirror state; the tool does not apply a single world-coordinate offset to every target.

- **Preserve tag rotation relative to the host** adjusts the tag angle for the difference between source and target host orientation. Turn it off when every tag should keep the same displayed angle.
- **Use model orientation** uses Revit's model-direction tag orientation. When disabled, the command uses compatibility orientation rules and common-angle snapping.
- **Copy leader** transfers a leader only when the reference has one, including supported end condition, elbow, and end position. Disable it to create no-leader tags.

> [!TIP]
> Use a carefully placed, representative tag as the reference. Its tag type, rotation, leader configuration, and position drive the batch result.

## Decide how to handle existing tags

Existing-tag matching is based on the target host and tag type ID, not on tag location or rotation.

| Existing Tags option | Behavior |
|---|---|
| **Replace all tags** | Deletes safely indexed tags on each target before creating reference tags. A multi-reference tag is unsafe and causes that target to be reported rather than retagged. |
| **Replace matching reference tag types only** | Replaces only tags whose type matches a chosen reference; other tag types remain. |
| **Skip matching reference tag types** | Leaves matching tags and adds only missing reference types. This is the default and safest repeat-run option. |

> [!WARNING]
> Review the Existing Tags setting before running. Replacement affects tags in the active view, and multi-reference tags are intentionally not deleted when safety cannot be confirmed.

## Run and check the result

1. Select valid example tags, or pick them in the window.
2. Choose the target mode and refine a manual target set if needed.
3. Set placement, leader, nested-instance, and existing-tag behavior.
4. Run the action and review tags in the active view.

The status line reports created tags, deleted existing tags, skipped matches, and failures. Creation is isolated by target/reference-type pair where possible, so a problem on one item does not necessarily stop the rest of the batch.

## Troubleshooting and limitations

| Situation | What to check |
|---|---|
| The command is unavailable | Use a floor plan, reflected ceiling plan, or drafting view. |
| A reference is rejected | It must be a single-reference component `IndependentTag` hosted by a loadable family instance. |
| References become invalid | All references must share one host. If a reference is deleted, pick a new set. |
| No targets are found | Check view visibility, selected target mode, category/type match, and the manual set. Enable nested instances only when needed. |
| Existing tags were not replaced | The tag may be multi-reference or not safely deletable; review the reported item manually. |
| Leader geometry differs | A target tag or Revit version may not support every copied leader property; review warnings. |

- Moving to another supported view retains references, clears targets, and recollects automatic targets.
- Changing documents requires closing and reopening the tool.
- Tool options are saved in its pyRevit configuration.

## Related pages

- [Clean Model Tags](clean-model-tags.md)
- [Navigate Tags and Hosts](navigate-tags-and-hosts.md)
- [Tag Behavior Settings](tag-behavior-settings.md)
