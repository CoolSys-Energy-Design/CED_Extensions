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

## Revit document and UI context lifetime — mandatory

pyRevit Rocket Mode can reuse one IronPython engine for an extension. Shared
module state can therefore outlive the command invocation and the document that
was active when the module was first imported.

- Do not capture or cache `revit.doc`, `revit.uidoc`,
  `__revit__.ActiveUIDocument`, `DB.Document`, `UIDocument`, Revit elements, or
  native `ElementId` objects in module-level variables inside model, domain,
  application-service, repository, or other reusable library modules.
- Do not use import-time defaults such as `def method(doc=revit.doc)` in new
  shared or model-layer code.
- Entry-point and UI composition code may read the active `doc`/`uidoc` at the
  command boundary, but it must pass that context explicitly to lower layers.
- A model object that already owns a Revit element should resolve its document
  from that element (for example, `self.circuit.Document`) rather than from the
  currently active UI document.
- Deferred/modeless work must acquire the active UI context inside
  `ExternalEvent.Execute`, validate it against the payload's captured document
  identity, and only then resolve IDs or elements.
- Tests for document-sensitive model logic must cover a circuit or element whose
  owning document differs from the active UI document.

Reference: pyRevit's Rocket Mode guidance warns that custom modules must not
retain active-document or element information in module-level variables.

## FamilySymbol/type-name access — mandatory for new Revit code

- Do not rely on `element.Name` or `family_symbol.Name` in IronPython; the
  CLR property binding is not reliable in Revit 2024.
- Use the shared `revit_helpers.get_family_symbol_name()` helper for display
  type names. Its first attempt must be `DB.Element.Name.__get__(family_symbol)`.
- For a `FamilySymbol` fallback, read
  `DB.BuiltInParameter.SYMBOL_NAME_PARAM` as `AsString()` or
  `AsValueString()`. `SYMBOL_NAME_PARAM` is a type-name fallback only; it is
  not the instance type-reference parameter.
- For a `FamilyInstance`, resolve its FamilySymbol first. If that path fails,
  use `DB.BuiltInParameter.ELEM_TYPE_PARAM` on the instance, preferring
  `AsValueString()` and otherwise resolving its native `ElementId` through the
  active document. `ELEM_TYPE_PARAM` is not available on a FamilySymbol.
- Keep the `ElementId` native while resolving it through Revit. Convert it to
  a numeric value only at serialization, logging, or other DTO boundaries.
