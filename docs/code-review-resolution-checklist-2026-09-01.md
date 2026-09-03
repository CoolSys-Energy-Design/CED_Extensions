# CED_Extensions Code Review Resolution Checklist — 2026-09-01

Source review: `docs/code-review-2026-09-01.md`

This is the working resolution tracker for the 32 numbered findings. A checked box means the issue has been closed by a verified fix, an explicit rejection, or an accepted disposition. Static verification and live Revit verification must be recorded separately where relevant.

Status values:

- `Open` — accepted underlying issue; no verified resolution yet.
- `Review` — partially valid, overstated, or dependent on an unresolved design decision.
- `Accepted` — behavior or risk was reviewed and consciously retained; no code change is planned.
- `Rejected` — the reported issue is not valid against the current repository/runtime.
- `Fixed` — implementation completed but required verification is not yet complete.
- `Verified` — resolution verified at the level stated in the comments.

| Done | ID | Priority | Status | Assessment | Resolution / comments |
|---|---|---:|---|---|---|
| [ ] | C1 | P0 | Fixed | Cancellation could assimilate preparatory changes made by Circuit Manager’s Mark as New/Existing, Neutral, Isolated Ground, and Edit Circuit Properties actions, plus the standalone Edit Circuit Properties command. | The three compound operations now assimilate only when calculation returns `status: ok`; every other result rolls back the outer group while preserving result metadata. Focused static contracts and the existing calculation-caller contracts pass. Live Revit verification remains: exercise a successful run and cancellation/no-eligible-circuit paths and confirm the model is unchanged after cancellation. |
| [ ] | C2 | P0 | Fixed | Creating or saving `CED_Circuit_Settings` could leave a started Revit transaction open when creation, value assignment, or commit raised. | Both settings transactions now perform best-effort `HasStarted()`/`RollBack()` cleanup and re-raise the original exception. Focused static contracts pass. Live Revit verification remains: force a settings initialization/save failure inside a caller-owned `TransactionGroup` and confirm the group can roll back cleanly and later commands still run. |
| [ ] | C3 | P0 | Fixed | Alerts Manager deferred actions carried numeric element IDs without the model identity that gave those IDs meaning. | Snapshots and action payloads now carry an open-document key using path, title, project-information identity, and runtime document identity. The external-event handler rejects mismatches before resolving an ID, refreshes the active model snapshot for review/retry, and verifies the target is an `ElectricalSystem`. Focused static contracts pass. Live Revit verification remains: open two models, switch between them with Alerts Manager open, and confirm Select/Recalculate/Hide never acts on the second model using the first model’s IDs. |
| [ ] | C4 | P2 | Open | Startup route discovery runs despite disabled telemetry and has a per-anchor budget. | Remove or defer startup discovery. Benchmark before and after; the original “unbounded” wording is inaccurate. |
| [ ] | C5 | P2 | Open | Base XAML dictionaries are recreated for each themed window. | Benchmark first. Any cache or application-level merge must preserve theme isolation and resource precedence. |
| [ ] | C6 | P2 | Open | ComboBox templates use non-virtualizing item hosts. | Restore virtualization and verify filtering, highlighting, keyboard navigation, and large-list behavior live. |
| [ ] | H1 | P1 | Open | Parameter-write failures can be indistinguishable from no-ops while the run reports success. | The shared setter is already storage-aware and `cg_apply` checks read-only state. Add an observable failure result and surface failed parameters. |
| [ ] | H2 | P1 | Open | A model-entered Include Neutral value is not treated as explicit on first calculation. | Preserve `Parameter.HasValue` state separately from preview-sourced state; test explicit Yes and No. |
| [ ] | H3 | P1 | Fixed | Feeder lookup used global `revit.doc` rather than the circuit’s document. | `_has_feeder_ln_voltage()` now resolves through `self.circuit.Document` and fails safe when the owning document is unavailable. A representative two-document contract proves lookup uses the circuit's owning document rather than active UI context. Live Revit verification remains: exercise feeder calculation while another model is active and confirm distribution-system lookup stays bound to the circuit's model. |
| [ ] | H4 | P1 | Open | Circuit Grouper lacks close/dispose lifecycle handling and callbacks can target a closed window. | Add symmetric teardown, closed-state guards, and a close policy while an operation is pending. |
| [ ] | H5 | P1 | Review | Automatic document retargeting can discard staged edits, but document-object identity is intentional safety. | Preserve invalidation for a reopened `Document`; replace destructive auto-retargeting with an invalid-context/reload decision. Do not blindly replace identity with only `PathName|Title`. |
| [ ] | H6 | P2 | Open | Installed pyRevit instantiates Circuit Manager during dockable-pane registration. | Evaluate a lightweight registered shell with deferred content construction; measure actual startup cost. |
| [ ] | H7 | P2 | Open | Startup reads and hashes every Python file in `CEDLib.lib`. | Replace or gate the content hash only if cache invalidation guarantees remain acceptable; benchmark deployed paths. |
| [ ] | H8 | P1 | Open | Circuit Manager fully rebuilds rows after operations. | Prefer targeted refresh by affected circuit IDs. Do not broadly enable the existing fast path because it reuses stale row data. |
| [x] | H9 | — | Accepted | The pane discards a valid standalone `staged_calculation` after preview and reruns the engineering calculation. Compound-operation reruns are intentional. | Closed by design decision: Circuit Manager selections are not expected to be large enough for the repeated calculation to matter, and the model is not expected to change between preview and acceptance in a way that changes the calculated result. No code change planned. |
| [ ] | H10 | P2 | Fixed | Tag Homeruns regenerated once per tag. | Wire Tools now follows Tag by Example's two-pass placement strategy: each tag is created in its own subtransaction, the complete batch regenerates once, and no-leader bounding-box clearance is refined afterward. Static contracts confirm no helper regenerates per tag and per-wire rollback remains. Live Revit verification remains: compare leader/no-leader placement and timing on a representative multi-homerun view. |
| [ ] | H11 | P2 | Open | FilterableComboBox performs uncoalesced layout and visual-tree work per keystroke. | Restore virtualization first, then coalesce highlight refresh and reduce container-tree traversal. |
| [ ] | M1 | P1 | Open | Several callers use the native-ElementId extractor where boundary coercion is required. | Preserve the intentional helper split; move numeric/native boundary callers to `coerce_elementid_value()` and add both input forms to tests. |
| [ ] | M2 | P1 | Open | Structured-search input moves the caret to the end after every edit. | Preserve the user’s caret when state synchronization does not rewrite the text. |
| [x] | M3 | P1 | Verified | `select_filter()` returned the last sorted token instead of the newly created token. | It now returns `new_token`; the grouping regression test verifies object identity after definition-order sorting. |
| [ ] | M4 | P1 | Open | ComboBoxes sharing one source also share the filtered view. | Give each behavior an independent view or filtered collection; verify multiple simultaneous row editors. |
| [ ] | M5 | P1 | Fixed | `IsViewValidForElementIteration` was called without the required `Document`. | Sheet Visibility QC now passes `(doc, view.Id)` and only skips the check when the API member is absent. A representative contract verifies the exact document and native view ID are supplied. Live Revit verification remains: scan a sheet containing schedule, legend, and unsupported view types and confirm only unsupported iterations are skipped. |
| [ ] | M6 | P2 | Open | SheetVisibilityQC finalization processes all candidates in one uncancellable callback. | Batch read-only finalization and account for candidate work in progress reporting. |
| [ ] | M7 | P2 | Open | Let There Be Light placement is a long blocking write with no practical cancellation. | Do not carry an open Revit transaction across ExternalEvents. Decide explicitly between atomic placement and committed cancellable batches. |
| [ ] | M8 | P2 | Open | Let There Be Light eagerly collects family types for every supported category. | Load the selected/default category first and collect other categories through the modeless API gateway on demand. |
| [ ] | M9 | P2 | Open | Several modeless tools omit `ExternalEvent.Dispose()`. | Add symmetric disposal and clear handler/window references without disposing a pending event prematurely. |
| [ ] | M10 | P3 | Review | Exact delegate retention is safer, but failed bound-method unsubscription was not demonstrated. | Store and remove the same delegate object as lifecycle hardening; verify under the supported IronPython host. |
| [x] | M11 | P3 | Verified | `theme_audit.py` contained a developer-specific path, but it is not imported by runtime code. The claimed missing DarkAlt file exists. | The audit now resolves its theme directory from `__file__`; it ran successfully from the repository without a developer-specific path. |
| [x] | M12 | — | Rejected | Docs has `bundle.yaml`, a generated catalog, deployed source pages, and an explicit release-manifest entry. | Closed as invalid against the current tree. Reopen only if packaging validation proves the documentation assets are omitted from the deployed build. |
| [ ] | M13 | P3 | Open | Documentation search rebuilds its search blob and result visuals on each keystroke. | Cache the joined search blob; measure before adding result debounce. |
| [ ] | M14 | P3 | Review | Prompts inside an open transaction group are undesirable, but `forms.alert` is a native Revit `TaskDialog`, not an unowned WPF dialog. | Move decisions to preflight where practical. Do not use the original hidden/unowned-window scenario as proof; C1 tracks the real cancel/rollback defect. |
| [ ] | M15 | P3 | Review | Revision Manager prepends persistent search paths; Tag by Example appends rather than prepends. | Prefer scoped import/loading or restore `sys.path` after imports. Verify actual persistent-engine sharing before claiming cross-command shadowing. |

## H9 clarification

The calculation backend intentionally supports two different preview contracts:

1. A standalone `calculate_circuits` request performs the engineering calculation before showing the preview and returns a plain-data `staged_calculation`. After the user accepts the preview, the caller should raise `apply_calculated_circuits` with that stage. The apply path revalidates the document and source snapshot before writing, so it avoids both duplicate calculation and stale application.
2. A compound request such as Mark Existing + Recalculate performs preparatory writes inside an outer `TransactionGroup`. Its preview response deliberately returns `preview_contract: rerun_original_operation` without a stage. The group is rolled back while the preview is shown, so accepting the preview must rerun the original compound operation to recreate those preparatory writes. A calculation-only stage cannot represent that mutation plan.

The ribbon Calculate Circuits command distinguishes these contracts. Circuit Manager currently does not: after every preview decision it raises the original `operation_key` again. That is correct for compound operations but unnecessarily recalculates standalone requests and bypasses the staged-result validation path.

The intended correction is therefore conditional, not “never rerun”:

- If `result["staged_calculation"]` is a dictionary, raise `apply_calculated_circuits` with that stage and the preview decision.
- Otherwise, preserve the original-operation rerun contract.

Disposition: accepted on 2026-09-01. The expected Circuit Manager selection sizes make the duplicate calculation cost immaterial, and the supported workflow assumes the model does not change between preview and acceptance in a way that would alter the recalculated result. The existing behavior will be retained unless field evidence contradicts those assumptions.

## Additional review dispositions

| Item | Status | Comments |
|---|---|---|
| `extension.json` wrapper / string booleans | Rejected | The installed pyRevit parser accepts both a flat object and `{"extensions": [...]}` and normalizes string boolean values. |
| QC repository ID type concern | Rejected for cited path | `distribution_equipment_repository` serializes the cited circuit collections to numeric IDs before the QC service consumes them. |
| Wire Tools `system_type=None` concern | Rejected | `_elecutils.is_circuit_eligible()` explicitly treats `None` as no system-type filter. |
| Sloped-host tag frame | Review | Independent projection can produce non-orthogonal axes, but the proposed cross-product fix would erase mirrored handedness. Requires a mirror-preserving design and targeted tests. |
| Plan-to-RCP tag-head Z | Live verification required | Static inspection does not establish the intended behavior across opposite view directions. |

## Verification record

Add dated entries here as findings move to `Fixed` or `Verified`.

- 2026-09-01 — Initial checklist created from the Codex audit. No production code changes were made.
- 2026-09-01 — H9 closed as an accepted design/performance tradeoff. No code change required.
- 2026-09-02 — C1–C3 implementation completed. Focused C1–C3 static contracts (4, including numeric and representative native `ElementId` inputs) and existing Calculate Circuits caller contracts (4) passed; all modified Python files passed syntax compilation and `git diff --check` reported no whitespace errors. Live Revit verification is still outstanding, so the rows remain unchecked with status `Fixed` rather than `Verified`.
- 2026-09-03 — H3, H10, M3, M5, and M11 implementation completed. Focused contracts passed for circuit-owned document lookup (3), batched Wire Tools tag regeneration with per-wire failure isolation (3), structured-search token identity (1), and the Sheet Visibility QC API signature (1); the portable theme audit ran successfully and `git diff --check` reported no whitespace errors. M3 and M11 are fully verified offline. H3, H10, and M5 remain unchecked as `Fixed` pending the live Revit checks recorded in their rows. M2 was intentionally left unchanged.
