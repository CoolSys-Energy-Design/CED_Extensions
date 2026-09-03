# Tag by Example

Tag by Example copies one or more reference tags to compatible elements in the active view. It copies the reference tag type, host-relative placement, rotation behavior, and—when enabled—leader geometry.

It works in floor plans, reflected ceiling plans, and drafting views. The current version supports single-reference `IndependentTag` elements hosted by loadable `FamilyInstance` elements. Room, area, space, wall, material, linked-model, and other non-component host workflows are not supported.

## References and targets

Select one or more reference tags with **Pick New Reference(s)**, or start with valid tags already selected in Revit. All references must be hosted by the same element. Each distinct reference tag type is copied once per target; selecting the same type more than once does not create duplicates.

The tool shows the reference tag type, host category/family/type, owner view, reference count, leader state, and orientation. Reference tags may belong to the active view or its related primary/dependent view.

Choose a target mode:

| Target mode | Finds | Best use |
|---|---|---|
| **All visible elements of the same type** | Visible elements with the same Revit type as the reference host. | Exact equipment/fixture standardization. |
| **All visible elements of the same family** | Visible elements in the same family, across types. | Applying one family-wide tagging convention. |
| **All visible elements of the same category** | Visible elements in the same category. | Broad category tagging. |
| **User selection** | **New** creates a set, while **Use Current Selection** captures the active Revit selection. **Clear** resets the saved set. | Hand-picked or exception-based tagging. |

Automatic targets are collected from the active view. The reference host, hidden elements, unsupported hosts, and incompatible elements are excluded. **Include nested family instances** adds nested family instances when they are visible and independently referenceable; some nested families cannot be tagged reliably.

For **Use Current Selection**, only supported elements matching the reference
host category are valid targets. Other selected elements are counted as
invalid and excluded from tag creation. **Clear** resets the tool's saved
manual targets without changing the active Revit selection.

## Placement options

- **Preserve tag rotation relative to the host** — Adjusts the tag angle for the difference between source and target host orientation while transferring the placement through each host's local frame. Turn it off when a consistent tag angle is preferred.
- **Use model orientation** — Uses Revit's model-direction tag orientation. When off, the tool uses compatibility orientation rules and common-angle snapping.
- **Copy leader** — Copies a leader only when the reference has one, including supported end condition, elbow, and end position. Turn it off to create no-leader tags.

Placement follows each target host's rotation and mirror state rather than applying one global XYZ offset.

## Existing Tags options

Matching is based on the target host and tag type ID—not tag position or rotation.

- **Replace all tags** — Removes all safely indexed existing tags on each target, then creates the reference tags. A multi-reference tag is considered unsafe; that target is reported and not retagged.
- **Replace matching reference tag types only** — Replaces only existing tags whose type matches a reference tag. Other tag types remain.
- **Skip matching reference tag types** — Leaves matching tags in place and creates only missing reference tag types. This is the default and is safest for repeated runs.

## Results and important behavior

The status line reports tags created, existing tags deleted, matching tags skipped, and failures. Creation is isolated per target/reference-type pair where possible, so one failed target does not necessarily stop the batch.

- Changing to another supported view retains the references, clears targets, and recollects automatic targets.
- Deleting a reference tag invalidates the reference set; pick a new reference.
- Changing the active document requires closing and reopening the tool.
- Existing multi-reference tags are not automatically deleted when deletion would be unsafe.

| Problem | Check |
|---|---|
| Tool is disabled | Use a floor plan, reflected ceiling plan, or drafting view. |
| No reference tags accepted | Select a single-reference component `IndependentTag` hosted by a loadable family instance. |
| References are invalid | All reference tags must be hosted by the same element. |
| Target count is zero | Check the target mode, view visibility, category/type match, and manual target set. Enable nested instances only when needed. |
| Existing tags are not replaced | The tag may reference multiple hosts or may not be safely deletable; review the reported item manually. |
| Leader placement differs | The target tag or Revit version may not support all copied leader properties; review warnings. |

