# Tag by Example

## Architecture summary

The command is a modeless `forms.WPFWindow`. UI event handlers only collect
settings and raise `ExternalEvent`; all Revit document reads, selection picks,
and writes execute in `lib/tag_by_example_events.py`.

The implementation is deliberately layered:

- `CEDLib.lib/Snippets/tag_host_transform.py` provides the reusable
  `HostPlacementFrame` and local/world point conversions.
- `CEDLib.lib/Snippets/tag_geometry.py` stores tag-head and leader geometry in
  host-local coordinates and contains pure conversion helpers.
- `lib/tag_api_compat.py` contains version-tolerant tagged-reference, leader,
  tag-creation, and type helpers.
- `lib/tag_host_adapters.py` implements Phase 1 support for
  `IndependentTag` on loadable `FamilyInstance` hosts and selection filters.
- `lib/tag_by_example_events.py` collects view-scoped candidates, keeps a
  same-host reference set, builds one existing-tag index, and performs a
  grouped, per-target-isolated batch operation. Its modeless lifecycle is
  event-driven: `ViewActivated` handles active-view changes, `DocumentChanged`
  detects deleted references, and `DocumentClosing` disables the window.
- `lib/tag_by_example_ui.py` owns the compact modeless WPF interface and
  pyRevit configuration persistence.

## Files created or modified

Created:

- `Tag by Example.pushbutton/script.py`
- `Tag by Example.pushbutton/TagByExample.xaml`
- `Tag by Example.pushbutton/bundle.yaml`
- `Tag by Example.pushbutton/lib/tag_api_compat.py`
- `Tag by Example.pushbutton/lib/tag_host_adapters.py`
- `Tag by Example.pushbutton/lib/tag_by_example_events.py`
- `Tag by Example.pushbutton/lib/tag_by_example_ui.py`
- `Tag by Example.pushbutton/tests/geometry_round_trip.py`
- `Tag by Example.pushbutton/README.md`
- `CEDLib.lib/Snippets/tag_host_transform.py`
- `CEDLib.lib/Snippets/tag_geometry.py`

Modified:

- `Tags.Panel/bundle.yaml` to add the command to the existing Tags panel.

The existing Orientation Tool files, including `_rotateutils.py`, are not
modified.

## Orientation logic reused or extracted

The Orientation Tools’ established intent—using insertion location, facing
orientation, and view-relative tag movement—was reviewed. Their
`_rotateutils.py` implementation uses a global offset rotated around Z, which
is not sufficient for a host-relative tag-copy operation, so it was left
unchanged. The new shared frame module prefers a point-based host's
`LocationPoint` for its origin and rotation. It uses the Revit instance
transform (`GetTotalTransform`, with `GetTransform` fallback) for the frame
axes, mirror state, and as an origin fallback when no `LocationPoint` is
available. Facing, hand, and mirrored flags are retained for reporting while
all three flip conditions are handled by the host frame automatically.

## Revit compatibility decisions

- Phase 1 supports `IndependentTag` with one local reference per selected
  reference tag and loadable `FamilyInstance` hosts.
- Multiple reference tags are allowed, but every reference must resolve to the
  same host. Each distinct reference tag type is copied to each target.
- **Use Current Selection** requires targets to match the reference host
  category, including when the selected reference is a multi-category tag.
  Other selected elements are retained as invalid diagnostics but are
  excluded from tag creation. The existing **New** picker behavior is retained.
- Reference tags are retained when the active view changes. Targets are
  cleared and recollected in the new floor plan, reflected ceiling plan, or
  drafting view.
- Tagged local IDs use `GetTaggedLocalElementIds()` where available and fall
  back to `TaggedLocalElementId`.
- Tag creation first attempts the type-aware `IndependentTag.Create` overload,
  then falls back to the older overload and `ChangeTypeId`.
- Leader elbow/end access uses newer reference-based methods when available,
  with older properties as fallback.
- No Revit 2027-only API is required.

## Known limitations

- Room, area, space, wall, material, linked-model, and multi-reference tags
  are intentionally deferred to Phase 2.
- Multi-reference existing tags are treated as unsafe to delete and are
  reported rather than modified.
- Manual targets are restricted to compatible component-family categories in
  Phase 1. Revit remains the final authority on tag-type compatibility; a
  failed target is isolated and reported.
- MEP curves, rooms, spaces, walls, linked-model hosts, and other non-component
  host adapters are deferred.
- Nested instances are included only when independently referenceable and
  visible. Some nested families cannot accept an independent tag and will be
  reported as failures.
- Families with unusual non-planar or custom reference behavior may require a
  future adapter-specific frame.

## Manual testing steps

1. Open a floor plan, reflected ceiling plan, or drafting view containing a
   component tag.
2. Select one or more supported tags on the same host and run **Tag by
   Example**, or use **Pick New Reference(s)**.
3. Confirm the displayed type, host, owner view, leader, and orientation.
4. Test same-type, same-family, same-category, and manual target modes. For
   manual targets, test **New**, **Use Current Selection**, and **Clear**.
   Include a mixed-category current selection and confirm that only elements
   matching the reference host category are valid.
5. Run the geometry debug routine `run_geometry_round_trip_tests()` from a
   pyRevit console.
6. Create tags and verify the completion report.
7. Test targets rotated 90 degrees and by a non-orthogonal angle.
8. Test mirrored, facing-flipped, hand-flipped, and mirrored-plus-rotated
   family instances.
9. Repeat with no leader, attached leader, and free leader examples.
10. Test **Replace all tags**, **Replace matching reference tag types only**,
    and **Skip matching reference tag types**. Matching is by tag type ID,
    regardless of tag head position or rotation.
11. Include a target that cannot be tagged and confirm other targets succeed.
12. Switch between supported views and a schedule, close the document, cancel
    Revit target selection, and delete a reference tag while the window is
    open. References should remain across supported view changes, targets
    should clear, and unsupported views should disable the tool with a clear
    status message.
13. Run the existing Orientation Tools before and after the tests to confirm
    their behavior is unchanged.

## Recommended Phase 2 work

Add separate adapters for walls and location curves, `RoomTag`, `AreaTag`,
`SpaceTag`, multi-category/material tags, linked-model references, and safely
shared nested families. Each adapter should define host-reference resolution,
view-plane frame construction, tag compatibility, and leader capabilities
without changing the Phase 1 FamilyInstance adapter.
