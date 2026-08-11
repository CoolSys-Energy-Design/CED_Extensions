# Parameter Monitor

Parameter Monitor is a read-only Phase 1 pyRevit tool for comparing selected
host- or linked-model data against an explicitly accepted project baseline.
Checks never update accepted values automatically. Resolve, Untrack, and Remove
actions update only Parameter Monitor-owned data.

## Project persistence

- Primary host: one dedicated
  `Autodesk.Revit.DB.ExtensibleStorage.DataStorage` element named
  `CED Parameter Monitor`.
- Backup host: a Revit global multiline-text parameter named
  `CED_Parameter_Monitor_Data`. It is read and written only when the preferred
  Extensible Storage route is unavailable or fails. A later successful save to
  DataStorage automatically restores the preferred route.
- Extensible Storage schema GUID: `3C41541B-82C2-4CD9-8E52-6B89E9B8E6F2`.
- Stable envelope schema: `CED_ParameterMonitor_v1`.
- Envelope fields: `SchemaVersion`, `ToolVersion`, `ProjectIdentity`, and
  `PayloadJson`.
- Payload schema version: `1`; load-time migrations live in `models.py`.
- The payload has a 12 MiB warning threshold and a 15 MiB hard safety limit.
  A future storage migration can split sets across DataStorage records without
  changing Tracking Set IDs.

This schema is intentionally separate from the generalized MEPRFP YAML/history
storage. Parameter Monitor does not attach entities to monitored model elements.

## Reused repository helpers

- `CEDLib.lib/Snippets/revit_helpers.py`: Revit-version-safe `ElementId`
  conversion, type resolution, and parameter lookup.
- `CEDLib.lib/Snippets/_elecutils.py`: filtering current electrical systems.
- `CEDLib.lib/UIClasses/ui_bases.py` and the shared resource dictionaries: CED
  WPF theme behavior.
- Existing Alerts Manager / Tag by Example patterns: one guarded
  `ExternalEvent` gateway for a modeless window.

## Phase 1 boundaries

- Manual baseline and checks only; no document-open or synchronize hooks.
- Source collection is category-scoped and excludes element types.
- Linked element identity includes the specific link instance `UniqueId`.
- Unloaded links produce `Source Unavailable` and never infer removals.
- Instance/type parameters are tracked separately with shared GUID or built-in
  identity preferred over names.
- Location support is intentionally limited to `LocationPoint`; coordinates are
  stored in the source document's internal coordinate system. A linked model's
  transform is compared once at source level.
- Host device relationships persist the host device `UniqueId`; circuit context
  is resolved live and never written back to Revit.
- Definition export JSON excludes accepted/current values, element identities,
  removed/untracked records, device relationships, and last-check metadata.

## UI and scale behavior

- External-event results are applied after Revit yields back to the WPF
  dispatcher. Grids are rebound with one .NET list assignment rather than one
  collection notification per row.
- The element grid uses row/column virtualization and extended selection for
  category sets containing thousands of elements.
- One selected row shows its property and device details. Multiple selected
  rows hide element-specific details and enable bulk model selection, show,
  untrack, and location-tracking actions.

## Tests

Run pure-Python tests outside Revit from this bundle directory:

```text
python tests/run_tests.py
```
