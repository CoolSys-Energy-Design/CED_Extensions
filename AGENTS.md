# Repository Agent Instructions

## Revit ElementId safety — mandatory

Before modifying Revit code, inspect and reuse the shared ElementId utilities in
`CEDLib.lib/Snippets/revit_helpers.py`. Do not create tool-local conversion
helpers without first verifying the shared helper contracts.

- Treat `Autodesk.Revit.DB.ElementId` and numeric ID values as different types.
- Never call `int()` on a value that might be a `DB.ElementId`.
- Never pass an existing `DB.ElementId` to a helper that expects an integer.
- Keep IDs returned by Revit APIs as `DB.ElementId` objects while interacting
  with the Revit API.
- Convert an ElementId to a numeric value only at serialization, DTO, logging,
  persistence, or external-event payload boundaries.
- Rehydrate a numeric ID exactly once, inside the Revit API execution context.
- ElementIds are document-specific. Deferred or modeless operations must verify
  the target document and confirm `doc.GetElement(element_id)` succeeds before
  acting.
- Revit selection collections must use `List[DB.ElementId]`.
- ID-handling tests must cover both numeric IDs and native or representative
  `DB.ElementId` inputs; numeric-only tests are insufficient.
- Search all callers before changing an ElementId helper's accepted input or
  output type.

## Revit API import conventions — new code only

For new Revit code and newly added imports, favor:

- `from pyrevit import DB, revit` for the core Revit database namespace and
  pyRevit document/application context.
- `import Autodesk.Revit.DB.Electrical as DBE` for the electrical namespace.
- `import Autodesk.Revit.DB.Mechanical as DBM` for the mechanical namespace.

These conventions are prospective. Do not revise, normalize, or reorganize
existing codebase imports or other existing code solely to conform to them
unless the user specifically requests a refactor.
