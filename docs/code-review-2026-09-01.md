# CED_Extensions Code Review — 2026-09-01

Scope: work landed on `develop` between 2026-08-18 and 2026-09-01 (20 commits, head `b7380a4`),
plus a startup and per-tool load budget for the whole extension set.

**How the window was determined:** history came from `.git/logs/HEAD` (reflog) and changed files
were identified by working-tree mtimes over the same window. Per-commit diffs were not available,
so every finding below is stated against **current file contents**. Line numbers were verified
against the working tree at the time of review — re-confirm before editing if the file has moved on.

**Counts:** 6 Critical · 11 High · 14 Medium.
**Headline numbers:** ~0.8–2.4 s added to Revit startup (8–22 s worst case); 19 XAML resource
dictionaries parsed per window opened.

---

## How to use this document

Each finding gives: severity, file path + line numbers, what is wrong, a concrete failure
scenario, and a specific fix. Findings with a failure scenario were confirmed by reading the code.
Where a conclusion depends on a file outside the reviewed set, the finding says so explicitly.

Repo rules that several findings turn on are in `AGENTS.md` — in particular the ElementId
handling contract and the `revit_helpers.get_family_symbol_name()` requirement. Read that file
before acting on anything here.

---

## CRITICAL

### C1. Cancelling a calculation commits the destructive prep work anyway

- `CEDLib.lib/CEDElectrical/Application/operations/mark_existing_and_recalculate_operation.py:169-198`
- `CEDLib.lib/CEDElectrical/Application/operations/set_include_and_recalculate_operation.py:137-168`
- `CEDLib.lib/CEDElectrical/Application/operations/edit_circuit_properties_and_recalculate_operation.py:108-142`

All three compound operations roll their `TransactionGroup` back on exactly three statuses —
`calc_preview_skipped`, `preview_required`, `stale` — then fall through to `tg.Assimilate()`.

```python
if ... == 'stale':
    try:
        tg.RollBack()
    ...
tg.Assimilate()   # reached on every other 'cancelled' reason
```

But `calculate_circuits_operation` also returns `cancelled` with reasons
`large_selection_cancel` (`calculate_circuits_operation.py:181`), `no_circuits` (`:171`),
`no_branches` (`:242`), `invalid_staged_result` (`:370`) and `missing_staged_result` (`:1889`).
None of those roll back.

**Failure:** user checks 1,200 circuits, picks Mark as Existing + Clear Wire + Clear Conduit,
hits Apply. The prep pass writes `CKT_User Override_CED = 1` and blanks the wire and conduit size
strings on all 1,200. Calculate then shows the >1000-circuit warning; the user clicks **Cancel**.
The group assimilates the wipe. The status bar reads "Operation cancelled" and nothing else, so
the user believes nothing happened.

**Fix:** invert the condition — roll back on anything where `status != 'ok'`, and only assimilate
on `ok`.

---

### C2. Settings transactions have no rollback, and run inside a caller's transaction group

- `CEDLib.lib/CEDElectrical/Domain/settings_manager.py:77-84` (`_create_global_param`)
- `CEDLib.lib/CEDElectrical/Domain/settings_manager.py:134-143` (`save_circuit_settings`)

```python
t = DB.Transaction(doc, "Create {}".format(GP_NAME))
t.Start()
gp = DB.GlobalParameter.Create(doc, GP_NAME, spec)   # no try/except
t.Commit()
```

Reached from `ensure_circuit_settings` -> `calculate_circuits_operation.py:106`, which runs inside
the still-open `TransactionGroup` the compound operations started.
`sync_electrical_parameter_bindings` at `settings_manager.py:441-446` already has the correct
guarded pattern; these two never got it.

**Failure:** workshared model where `CED_Circuit_Settings` does not exist yet and another user owns
the global-parameter element. `Create` throws, `t` stays open, the outer `tg.RollBack()` then also
throws (a group cannot roll back over a live child transaction) and is swallowed. Every subsequent
edit in the session fails with "another transaction is already active" until Revit restarts.

**Fix:** wrap both bodies with `if t.HasStarted(): t.RollBack()` and re-raise, matching `:441-446`.

---

### C3. Alerts Manager resolves stored ElementIds against whatever document is active

- `CED ElecTools.extension/AE pyTools.tab/Electrical.panel/Alerts Manager.pushbutton/script.py:163-213`

```python
uidoc = application.ActiveUIDocument
doc = uidoc.Document if uidoc else None
circuit_id = int(payload.get("circuit_id") or 0)
circuit = doc.GetElement(_idfrom(circuit_id))
```

No document key is stored anywhere and there is no `DocumentClosing` subscription. Direct
violation of the AGENTS.md rule that deferred/modeless operations must verify the target document.

**Failure:** Alerts Manager open on Model A. User switches to Model B and clicks Recalculate. The
id resolves against Model B and frequently lands on an unrelated `ElectricalSystem` (circuit ids
cluster). CED result parameters are written to the wrong circuit in the wrong model, silently, and
the tool reports success. `set_hidden` has the same exposure.

**Fix:** copy the Wire Tools pattern — store `"{}|{}".format(doc.PathName, doc.Title)` at snapshot
time (see `wire_tools_events.py:63`), compare at the top of `Execute`, bail with `invalid_context`
on mismatch, and `isinstance`-check the resolved element against `DBE.ElectricalSystem`.

---

### C4. Every Revit launch runs an unbounded cloud-folder scan for a dialog that can never appear

- `AE pyTools.extension/startup.py:1418`, `:1326-1338`
- `AE pyTools.extension/telemetry_route.py:25-27`, `:268-286`, `:493-516`

Two defects that compound.

**(a) The discovery budget is per-anchor, not global.** `DISCOVERY_MAX_SECONDS = 4.0` is started
inside `_discover_project_roots_under` (`telemetry_route.py:204`). With `C:\ACC`, `C:\DC`, `~\ACC`
and `~\DC` present the worst case is 4 s x 4 = 16 s of blocked startup. These are Autodesk Desktop
Connector virtual paths, so every `os.listdir` / `os.path.isdir` can trigger a synchronous cloud
metadata fetch — the walk hits the time cap far more often than a local-disk walk would.

**(b) The whole thing is unreachable.** `_known_candidate_roots()` (`telemetry_route.py:268`)
appends five paths with **no `isdir` check**, so `candidate_count` is always >= 5 and this always
returns early:

```python
if int(route_result.get("candidate_count", 0) or 0) > 0:
    # Candidates exist; user can resolve manually from ACC Path Resolver.
    return
```

The ~75 lines of WPF construction below it, ending in a blocking `win.ShowDialog()` at
`startup.py:1415`, are dead except on the exception fall-through — where they would block Revit
startup indefinitely waiting for a click on a possibly-unattended machine. Users who genuinely have
not synced ACC never see the instructions. Meanwhile `persist=True` rewrites a JSON state file on
every launch for a feature both release switches (`startup.py:60-61`) have disabled.

**Fix:** take route resolution off startup entirely — resolve lazily on first use from the ACC Path
Resolver tool, cache the resolved root with a TTL, pass `persist=False` from anything
startup-adjacent, and never call `ShowDialog()` from a startup script. If a startup check is kept,
gate the dialog on `viable_count == 0` (not `candidate_count`) and make the discovery budget global
and ~250 ms.

---

### C5. Every window parses the entire design system from disk; nothing is cached or shared

- `CEDLib.lib/UIClasses/resource_loader.py:22-41` (18 default paths)
- `CEDLib.lib/UIClasses/resource_loader.py:149-157` (`_load_dictionary`)
- `CEDLib.lib/UIClasses/resource_loader.py:193-218` (`ensure_base_resources`)
- `CEDLib.lib/UIClasses/ui_bases.py:689`, `:811`

Every `CEDWindowBase.__init__` / `CEDPanelBase.__init__` ends in `apply_ced_theme()` ->
`apply_theme()` -> `ensure_base_resources()`, which constructs a brand-new `ResourceDictionary`
per file per window. There is no `Application.Current.Resources` merge, no module-level cache of
parsed dictionaries, and no reuse between windows.

```python
dictionary = ResourceDictionary()
dictionary.Source = Uri(path)   # WPF parses synchronously, per window
```

**Cost:** 18 base dictionaries + 1 theme = **19 parses per window**. The files present on disk total
~200 KB of XAML — `InputStyles.xaml` 43 KB, `ListStyles.xaml` 38 KB, `DataGrids.xaml` 32 KB —
containing ~96 styles and ~55 control templates. A tool that needs one button style pays for all of
it. Only `Documentation/viewer.py:55-63` overrides the default list. On a network-hosted
`CEDLib.lib` this is the dominant part of "click button, wait".

**Fix:** merge the base dictionaries **once** into `Application.Current.Resources` at extension
startup and have `ensure_base_resources` no-op when that merge is present — the existing
"already merged" check at `:199-204` makes this safe. Failing that, add a module-level cache in
`resource_loader` keyed by normalized absolute path; one `ResourceDictionary` instance may legally
appear in many `MergedDictionaries`. Then trim `DEFAULT_BASE_RESOURCE_RELATIVE_PATHS` to a minimal
set and make `DataGrids.xaml` / `ListStyles.xaml` opt-in.

*Highest-leverage change in this review. Touches one file and no tool code.*

---

### C6. All three CED ComboBox templates disable WPF virtualization

- `CEDLib.lib/UIClasses/Resources/Styles/InputStyles.xaml:376`, `:580`, `:671`
- `CEDLib.lib/UIClasses/Resources/Styles/ListStyles.xaml:311`, `:398`

```xml
<StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
```

`IsItemsHost="True"` on a bare `StackPanel` hard-codes a non-virtualizing items host and makes the
ComboBox's default `VirtualizingStackPanel` irrelevant. Affects
`CED.Input.ComboBox.Template.Flat` (`:321`), `.Flat.Compact` (`:522`) and `.Flat.Compact.Editable`
(`:610`) — the last is what `FilterableComboBox` runs on.

**Failure:** a panel or type picker with 2,000 entries realizes 2,000 `ComboBoxItem` containers on
first open, and again on every filter refresh. Multi-second freeze.

**Fix:** replace each with `<VirtualizingStackPanel IsItemsHost="True" VirtualizationMode="Recycling"/>`,
or use `<ItemsPresenter/>` and set `VirtualizingPanel.IsVirtualizing="True"` /
`VirtualizationMode="Recycling"` on the ComboBox styles at `:254`, `:406`, `:701`, `:708`.

Note: `ListStyles.xaml:735` (ListBox) and `CED.ListView.Virtualized` at `:449-456` are correct and
should be left alone. The two `CanContentScroll="False"` hits elsewhere in the repo
(`Calculate Circuits/settings.xaml:233`, `WireTools.xaml:208`) wrap static `StackPanel`s with no
items host — pixel-scrolling choices, not virtualization bugs.

---

## HIGH

### H1. Parameter writes have no read-only or storage-type check, and failures are swallowed

- `CEDLib.lib/CEDElectrical/Infrastructure/Revit/writers/revit_circuit_writer.py:18-25`, `:87-96`
- `.../CreateCircuitsByDeviceParameter.pushbutton/cg_apply.py:72-91` (`_set_text`, `_set_double`)

```python
try:
    revit_helpers.set_parameter_if_changed(param, value)
except Exception:
    continue
```

No `IsReadOnly` check, no storage-type check, the boolean return is discarded, and
`apply_staged_result` reports `'status': 'ok'` with `updated_circuits = len(approved_items)`
(`calculate_circuits_operation.py:600-605`) regardless. `cg_apply` already imports
`revit_helpers` (`:15`) without using its guarded helper for these writes.

**Failure:** a template where `Conduit and Wire Size_CEDT` is bound as a **type** parameter instead
of instance — a state `_sync_parameter_bindings` can itself produce (`settings_manager.py:299`).
Every `Set` throws, every exception is swallowed, 3,000 circuits keep stale wire callouts, panel
schedules print last month's sizes, and the pane reports "Calculated 3000 circuits."

**Fix:** check `param.IsReadOnly` and `StorageType` before writing, accumulate a per-circuit
`failed_params` list, return it, and surface it in `_show_run_summary_if_needed`. Same treatment for
the ignored `False` return of `ParameterAlertStore.write_alert_payload`
(`calculate_circuits_operation.py:551-553`).

---

### H2. First calculation overwrites an explicit user "Include Neutral = Yes"

- `CEDLib.lib/CEDElectrical/Model/CircuitBranch.py:2526-2537`, `:734-751`, `:1211-1226`

```python
param = self.circuit.get_Parameter(Guid(guid))
if not param:
    return False, False
return False, bool(param.AsInteger())   # 'explicit' is ALWAYS False here
```

The `explicit` flag is only ever `True` for a staged **preview** value; a value the user typed into
the model can never be explicit. `_should_apply_default_multipole_branch_neutral` (`:1222`) relies
on that flag, and its only other guard is `_has_existing_branch_results()` — false on a first
calculation.

**Failure:** engineer creates a new 3-pole 208 V branch circuit feeding a receptacle bank, ticks
`CKT_Include Neutral_CED = Yes` in Properties, runs Calculate for the first time. `_load_core_inputs`
reads `(False, True)`; the default resets `_include_neutral = False` at `:749`;
`_collect_shared_param_values` writes `0` back to the model. The wire callout drops the neutral and
the user's edit is gone.

**Fix:** return `bool(param.HasValue)` as the explicit flag for the model-read branch
(`revit_helpers.parameter_matches_value:157-158` already models this distinction), and keep a
separate flag for preview-sourced values.

---

### H3. Feeder neutral detection resolves an ElementId against `revit.doc`

- `CEDLib.lib/CEDElectrical/Model/CircuitBranch.py:1192-1209`

```python
def _has_feeder_ln_voltage(self):
    doc = revit.doc
    ...
        ds_elem = doc.GetElement(ds_param.AsElementId())
```

AGENTS.md violation. Every other site in this stack threads the operation's `doc` through correctly
— compare `calculate_circuits_operation.py:776`.

**Failure:** two models open; the user switches the active view during the queued external event.
For a FEEDER circuit in Model A the lookup runs against Model B and returns `None` or an unrelated
element. `_expected_neutral_qty` (`:1186-1187`) returns 0, the neutral is silently dropped, and
`CKT_Include Neutral_CED` is written as `0` (`calculate_circuits_operation.py:1676`). An undersized
feeder is issued for construction with no warning.

**Fix:** `doc = getattr(self.circuit, "Document", None)` (or thread the operation's `doc` into the
constructor), guarded for `None`.

---

### H4. Circuit Grouper has no close handler: nothing unsubscribes, the ExternalEvent leaks

- `.../CreateCircuitsByDeviceParameter.pushbutton/cg_window.py:2592-2593`
- `.../CircuitGrouperWindow.xaml` — no `Closing=` on the `Window` element
- `.../script.py:146-151`

```python
def cancel_clicked(self, sender, args):
    self.Close()
```

The window was made modeless without the lifecycle contract. `ExternalEvent.Create` is called and
`Dispose()` never is, so every open/close cycle leaks a live `IExternalEventHandler` registered with
Revit for the rest of the session. Wire Tools does this correctly at `wire_tools_ui.py:962-965`
(`gateway.detach_lifecycle()`).

**Failure:** user closes the window while a Run is pending. `Execute` still runs (correct), then the
callback chain `_handle_run_complete` -> `Dispatcher.BeginInvoke` -> `_finish_run_complete` calls
`self.Show()` / `self.Activate()` on a closed `Window`. That throw is caught at `:2584-2588`, but
`_enter_results_mode(report)` at `:2589` is not — so a batch of circuits gets created and the user
is shown nothing about it. Nothing resets `_run_in_progress` either.

**Fix:** add `Closing="window_closing"` to the XAML; set a `self._closed` flag checked at the top of
every `_finish_*` callback; add and call `gateway.dispose()` which calls `ExternalEvent.Dispose()`;
warn (or refuse) on close while `_run_in_progress`.

---

### H5. A document switch silently destroys every staged Circuit Grouper edit

- `.../cg_window.py:739-776` (`window_activated`), `:726-737` (`_document_key`)

```python
if current_key != self._document_key:
    replacement = show_modeless(snapshot, self._gateway, activate=True)
    ...
    self.Close()
```

Two problems stacked:

1. **The document key is wrong.** `_document_key` returns `str(doc.GetHashCode())` — the CLR
   reference hash. Close Model A and reopen the same file: new `Document` object, new hash, so the
   tool treats the same file as a different document. `_raise_navigation:1688` duplicates the same
   hash logic inline, so the two can disagree.
2. **A detected change rebuilds from a fresh snapshot and closes the old window.** All hand-edited
   Load Names, panel selections, breaker ratings, schedule notes, manual regroupings and drag-drop
   moves are discarded; the replacement is seeded from `_init_group_defaults` only, reading the
   *other* document's live selection (`script.py:81-96`).

**Failure:** engineer spends ten minutes grouping 400 devices and setting panels, alt-tabs to a
second open model to check a panel schedule, alt-tabs back and clicks the grouper. Everything is
gone; the status line says "Target document updated."

**Fix:** use `"{}|{}".format(doc.PathName, doc.Title)` from a single shared helper used by both call
sites; subscribe to `Application.DocumentClosing` and `revit.events` `doc-changed` with teardown in
the new `Closing` handler; on a document change disable the window behind an invalid-context overlay
(Wire Tools' `set_invalid_context` is the established pattern) rather than auto-replacing it.

---

### H6. The Circuit Manager dockable panel is fully constructed at startup for every user

- `CED ElecTools.extension/startup.py:184`, `:195-196`

```python
panel_module = imp.load_source("ced_electools_circuit_manager_panel", panel_path)
...
if not forms.is_registered_dockable_panel(panel_cls):
    forms.register_dockable_panel(panel_cls, default_visible=False)
```

`imp.load_source` compiles and executes a 6,456-line module whose imports (`CircuitBrowserPanel.py:48-80`)
pull in `Snippets.revit_helpers`, `categories`, `_elecutils`, `circuit_ui_actions`, ~10
`CEDElectrical` modules and the whole `UIClasses` package (~4,500 lines). Registering the pane then
instantiates it, so `__init__` parses the 43 KB `CircuitBrowserPanel.xaml` and `_try_apply_theme`
(`:2479`) merges all 19 resource dictionaries.

**Cost:** ~0.5–1.5 s on every Revit start, for every user, whether or not they ever open Circuit
Manager (`default_visible=False`).

**Fix:** register a thin placeholder pane class defined *in* `startup.py` — no CEDLib imports, no
XAML, an empty `ContentControl` in `SetupDockablePane`. On its first `Loaded` / `IsVisibleChanged`,
import `CircuitBrowserPanel`, build the real content, assign it.

**Preserve:** the panel correctly defers all Revit DB queries to `panel_loaded` / `IsVisibleChanged`.
Do not regress that.

---

### H7. SHA-1 of every `.py` in CEDLib.lib is computed on every Revit start

- `CED ElecTools.extension/startup.py:63-83`, called from `:127`

```python
digest = hashlib.sha1()
for current_root, dirs, files in os.walk(lib_root):
    ...
        with open(file_path, "rb") as source:
            while True:
                chunk = source.read(65536)
```

The hashing is trivial; the cost is the full `os.walk` plus a synchronous open of every file. Note
also that on a cold process `script.get_envvar` returns nothing, so `reset_modules` is `True` on
every Revit start (`:132-135`) — the cache-invalidation branch never saves work at process start,
only within a persistent engine session.

**Cost:** ~20–60 ms on a local SSD; **seconds** if `CEDLib.lib` is deployed from a network share or
an ACC-synced folder. Grows linearly as the library grows.

**Fix:** fingerprint on `(relative_path, st_mtime, st_size)` from `os.walk`'s own stat data — same
invalidation semantics, no file reads — or gate the check behind a `CED_LIB_DEV` env var, since it
only matters for developers with two checkouts.

---

### H8. Calculating one circuit rebuilds every row in the Circuit Manager; the cheap path is dead code

- `.../Circuit Manager.pushbutton/CircuitBrowserPanel.py:3693-3707` (`_load_items_full`)
- `:3709-3733` (`_load_items_fast`)
- `:2847`, `:2872`, `:2877` (`_on_operation_complete`)
- `:3130` (`_safe_load_items`, `fast=False` default)

Each `CircuitListItem.__init__` (`:773-905`) does ~8 name-based parameter resolutions via
`_lookup_param_value` (`LookupParameter` is a linear name scan of the element's parameter set) plus
a `json.loads` of the `Circuit Data_CED` payload. `_load_items_full` runs that for every circuit in
the document, on the WPF UI thread, with no progress and no cancel — and it is called from
`_on_operation_complete`, so a single-circuit calculate rebuilds all of them.

`_load_items_fast`, which reuses existing items, exists and is correct. **No call site ever passes
`fast=True`** — every caller of `_safe_load_items` omits the argument.

**Cost:** a 5,000-circuit model means ~40,000 parameter-name scans and 5,000 JSON parses per open,
repeated after every calculate, every action, every view activation and every document open.

**Fix:** call the existing `_refresh_items_by_circuit_ids(request.circuit_ids)` (`:4465`, already
used by the edit-properties path) from `_on_operation_complete`; resolve the shared-parameter
`Guid`s once and use `get_Parameter(Guid)` as `CircuitBranch._get_param_value:2548` already does;
wire up `fast=True` or delete it.

---

### H9. The Circuit Manager discards the staged calculation and runs the whole thing twice

- `.../CircuitBrowserPanel.py:2879-2924` (`_show_calculation_preview`)
- compare `.../Calculate Circuits.pushbutton/script.py:182-193`

The backend returns two preview shapes: `preview_contract: 'rerun_original_operation'` without a
stage for compound callers (`calculate_circuits_operation.py:266-273`), and a full
`staged_calculation` for standalone callers (`:330-337`). The ribbon button handles both. The pane
always re-raises the original operation key, discarding `result["staged_calculation"]`.

**Cost:** Calculate All from the pane with preview on, 3,000 circuits — pass 1 does collection, lock
partitioning, 3,000 `CircuitBranch` constructions, the full engineering sequence and a validation
snapshot, then throws it away and repeats. Exactly double the runtime of the ribbon button, and the
stale-detection contract added by the hardening commit (`_validate_staged_item`, `:1095`) is never
exercised from the pane.

**Fix:** mirror `script.py:182-193` — when `result.get("staged_calculation")` is a dict, raise
`apply_calculated_circuits` with `{"calc_preview_decision": decision, "staged_calculation": staged}`;
otherwise fall back to the re-run.

---

### H10. `doc.Regenerate()` runs once per tag inside the Tag Homeruns loop

- `.../Wire Tools.pushbutton/lib/wire_tools_logic.py:2964-3033`, with `Regenerate` at `:2723` and `:2839`

`tag_homeruns` loops over every wire id, opens a `SubTransaction` per wire, and calls
`_create_wire_tag` -> `_place_leader_tag` / `_place_no_leader_tag`, each of which regenerates the
document so `tag.get_BoundingBox(view)` is valid for clearance refinement.

**Failure:** Select Homeruns on a typical lighting plan returns 400–800 open-ended wires. Tag
Homeruns then performs 400–800 full document regenerations plus as many sub-transaction commits,
inside `ExternalEvent.Execute` — 30 s to 4 min of frozen, unresponsive Revit with no progress bar
and no cancel.

**Fix:** create all tags in one pass, `Regenerate()` **once**, then run the placement-refinement
pass. Drop the per-wire `SubTransaction` and collect failures instead. Add progress/cancel above a
few dozen wires.

---

### H11. FilterableComboBox does a full layout pass and a per-item visual-tree walk on every keystroke

- `CEDLib.lib/UIClasses/filterable_combo_box.py:1155-1187`, `:1224-1241`, called from `:682` -> `:1090`

```python
self.combo.UpdateLayout()
count = int(self.combo.Items.Count)
generator = self.combo.ItemContainerGenerator
for index in range(count):
    container = generator.ContainerFromIndex(index)
    text_blocks = self._visual_descendants(container, TextBlock)
```

There is no debounce anywhere in `UIClasses` — no `DispatcherTimer`, no coalescing. `_begin_invoke`
at Background priority (`:1422-1437`) does not coalesce either, so N keystrokes queue N full
refreshes. `_visual_descendants` collects the entire subtree even though only `[0]` is used, and
`_move_candidate` (`:741-768`) re-runs `_apply_filter` on every arrow key.

**Cost:** combined with C6, typing five characters into a 2,000-item picker performs
5 x (2,000 container realizations + 2,000 subtree walks + 2,000 `Inlines` rebuilds + a full
`UpdateLayout`). Visible per-character lag and dropped keystrokes above a few hundred items.

**Fix:** fix C6 first; then coalesce highlight refreshes behind a `_highlight_pending` flag, add a
~150 ms `DispatcherTimer` debounce between `TextChanged` and `_apply_filter`, drop `UpdateLayout()`,
and do highlighting with a value converter or attached property in the `DataTemplate` rather than by
walking containers.

---

## MEDIUM

### M1. Shared ElementId helper silently drops numeric input

- `CEDLib.lib/Snippets/revit_helpers.py:11-22` (`get_elementid_value`)
- `CEDLib.lib/Snippets/categories.py:18-28` (`category_id_value`)

A plain `int` has neither `.Value` nor `.IntegerValue`, so `get_elementid_value(5)` returns the
fallback, not `5`. `category_id_value` — documented as accepting "Category/ElementId-like inputs" —
routes non-`ElementId` input straight into it at `categories.py:28`, yielding
`_INVALID_CATEGORY_ID_VALUE` (-1). `coerce_elementid_value` (`revit_helpers.py:25-42`) exists for
exactly this and handles both forms correctly.

**Consequence:** pass a numeric category id (from a saved config, an external-event payload, or
`category_id_values`) to `unique_category_ids` / `build_category_set` and the category is silently
dropped — a parameter binding or filter comes back empty with no error.

Three tools have already written local workarounds, which AGENTS.md explicitly forbids:

- `.../SheetVisibilityQC.pushbutton/script.py:185-209` — with the comment *"The shared helper only
  understands ElementId-like objects and otherwise returns its fallback, which made valid selected
  rows look like id 0."*
- `.../Let there be Light.pushbutton/script.py:143-148`
- `.../Wire Tools.pushbutton/lib/wire_tools_logic.py:105-111`

And one live failure: `cg_apply.py:110-114` does `value = revit_helpers.get_elementid_value(token)`
then `if value:` — falsy for an integer token, so the documented numeric fallback returns `None`.

**Fix:** point `categories.category_id_value` at `coerce_elementid_value`, delete the three local
helpers, fix `cg_apply.py:110-114`, and add a test with plain `int` input (AGENTS.md requires
coverage of both numeric and native `ElementId` forms).

---

### M2. Search-box caret jumps to end on every keystroke

- `CEDLib.lib/UIClasses/structured_search_box.py:569-591` (`_set_input_text`)

```python
if self._input.Text != text:
    self._input.Text = text
self._input.CaretIndex = len(text)   # unconditional
```

The `Text` assignment is correctly guarded; the `CaretIndex` assignment is not, and it runs on every
`TextChanged`.

**Failure:** the user clicks into the middle of an existing search string to fix a typo; the next
character teleports the caret to the end and the text comes out scrambled. Reads to users as
"the search box is broken."

**Fix:** move the `CaretIndex` assignment inside the `if`, and preserve the caret offset when the
state did not rewrite the string.

---

### M3. `select_filter` returns the wrong token

- `CEDLib.lib/UIClasses/structured_search.py:734-747`

```python
new_token = SearchFilterToken(definition)
self._tokens.append(new_token)
self._organize_tokens()
self._active_token_index = self._tokens.index(new_token)
...
return self._tokens[-1]
```

`_active_token_index` is computed correctly by identity; the return value is `_tokens[-1]`, a
different token whenever the new one sorts before an existing one. `tools/tests/test_structured_search.py:222-231`
walks exactly this path but asserts only token order. The WPF control
(`structured_search_box._select_current_suggestion:1144-1151`) only null-checks the result, so it is
accidentally immune — but `select_filter` is the documented API.

**Fix:** `return new_token`.

---

### M4. Filter installed on the shared default CollectionView

- `CEDLib.lib/UIClasses/filterable_combo_box.py:1017-1039`, `:1071-1096`

`ItemsControl.Items` always exposes `Filter`/`Refresh`, so the fast path at `:1017-1039` always wins.
When `ItemsSource` is set, `Items` delegates to the **default** `CollectionView` for that source —
shared by every `ItemsControl` bound to the same list. `_apply_filter` then installs
`lambda value: self._matches(value)` on it.

`.../Let there be Light.pushbutton/script.py:2929` creates one `FilterableComboBox` per DataGrid row,
all against the same shared options list.

**Consequence:** (a) typing in one row's editor filters every other row's dropdown; (b) the closure
holds a strong reference to the behavior -> ComboBox -> window from a collection that outlives the
window. `dispose()` (`:332-351`) clears it, but only one of three consumers calls it
(`Theme.pushbutton/script.py:294` and one Light call site do not).

**Fix:** give each combo its own view (`CollectionViewSource.GetDefaultView` over a per-control
`ObservableCollection`), or filter into a per-control collection directly. Make `dispose()`
mandatory from `Window.Closed`.

---

### M5. `IsViewValidForElementIteration` called with one argument

- `.../SheetVisibilityQC.pushbutton/script.py:541-553`

```python
validity_check = getattr(DB.FilteredElementCollector, "IsViewValidForElementIteration", None)
if validity_check is not None:
    try:
        if not bool(validity_check(view.Id)):
```

The Revit API signature is static `IsViewValidForElementIteration(Document document, ElementId viewId)`.
The one-argument call raises `TypeError`, swallowed by the bare `except Exception: pass`.

**Failure:** for every schedule or legend on a selected sheet the intended clean skip never happens;
`FilteredElementCollector(doc, view.Id)` at `:555` throws `ArgumentException`, bubbles to
`_advance_scan_state:792-798`, and the user gets a raw exception string in the warnings list on every
scan.

**Fix:** `validity_check(doc, view.Id)`, and narrow the `except` so an arity error is not mistaken
for an old Revit version.

---

### M6. SheetVisibilityQC scan finalization is unbounded and uncancellable

- `.../SheetVisibilityQC.pushbutton/script.py:727-767`, `:611-619`, invoked from `:807-812`

The per-view scan is correctly chunked, but the finalize builds a full row DTO for every candidate
element in the document — one `doc.GetElement` plus category/family/type/level reads each — inside a
single `ExternalEvent.Execute`, with no progress tick and no cancel check.
`_collect_candidate_ids` is document-wide, not sheet-scoped, and `state["total"]` counts the finalize
as one tick (`:723`).

**Failure:** the progress bar reaches 100%, then Revit hard-freezes for minutes with Cancel doing
nothing.

**Fix:** slice `sorted(candidate_values)` into batches of ~500 per `scan_step`, tick progress against
`len(candidates)`, and honor `gateway.cancel_requested` between batches.

---

### M7. Let there be Light placement has no progress or cancel, and the rollback path can misfire

- `.../Let there be Light.pushbutton/script.py:1482-1543`, `:1299-1328`, `:3825-3827`

```python
transaction.Commit()
self._normalize_world_z_bulk(normalization_records, report)
group.Assimilate()
except Exception as ex:
    report.fatal_error = _safe_text(ex)
    try:
        transaction.RollBack()
```

The whole placement loop plus `_build_existing_index`'s document-wide collector runs in one
`ExternalEvent.Execute`. `_place_clicked` only sets a status label. If `group.Assimilate()` throws,
`transaction.RollBack()` is called on an already-committed transaction.

**Failure:** a 5,000-row CSV import freezes Revit into "not responding" and the user force-kills it
mid-transaction.

**Fix:** chunk placement across repeated external events (the resumable-state pattern
SheetVisibilityQC already uses), and guard with
`if transaction.HasStarted() and not transaction.HasEnded():` as
`RevisionManager.pushbutton/script.py:614-624` correctly does.

---

### M8. Let there be Light runs 13 full collectors before the window appears

- `.../Let there be Light.pushbutton/script.py:1712-1734`, called from `main()` at `:3915` before `window.Show()`

```python
categories = collect_category_options(doc)
for category in categories:
    target_options_by_category[int(category.category_id)] = collect_host_fixture_options(
        doc, category.category_id,
    )[0]
```

`SUPPORTED_CATEGORY_SPECS` (`:81-95`) has 13 entries including `OST_GenericModel`. Each iteration
runs a `FilteredElementCollector.OfClass(FamilySymbol)` (`:1581-1586`) and, per symbol,
`_collect_type_parameter_values` (`:354-381`), which enumerates and stringifies every type parameter.

**Failure:** 10–60 s freeze after the ribbon click with no window and no progress on a large model.
Repeats on every Refresh (`:2689`).

**Fix:** build only the default category's options eagerly; collect the rest lazily on category
selection. Drop `_collect_type_parameter_values` from the eager path.

---

### M9. ExternalEvents are never disposed in three modeless tools

- `.../SheetVisibilityQC.pushbutton/script.py:996`, close handler at `:1748-1770`
- `.../Let there be Light.pushbutton/script.py:1749`, close handler at `:2762-2773`
- `.../Tag by Example.pushbutton/lib/tag_by_example_events.py:641`

`ExternalEvent` implements `IDisposable`; none of the three close paths calls `Dispose()`.
SheetVisibilityQC is worst: its single-instance guard depends on `Application.Current`
(`:2856-2871`), which can be null in some Revit/pyRevit hosts — then `_existing_window()` returns
`None` every time and each ribbon click leaks a whole window, gateway, handler and ExternalEvent.

**Fix:** dispose the event and null the handler in each close handler; back the SheetVisibilityQC
single-instance check with a `__revit__`- or `EXEC_PARAMS`-scoped registry, as Let there be Light
does with `_pyrevit_engine_id` (`:3876-3886`).

---

### M10. Theme tool subscribes a bare bound method to a session-lifetime static event

- `.../Theme.pushbutton/script.py:162-175`

```python
InputManager.Current.PreProcessInput += self._on_unlock_input
```

Every other tool in the repo stores the delegate (`Let there be Light/script.py:1796`,
`tag_by_example_events.py:653-656`). In IronPython, `-=` on a bound method is not reliably matched to
the delegate `+=` created, and the removal is wrapped in a bare `except: pass` (`:171-174`).

**Failure:** one failed detach leaves a Python callback inspecting every keyboard event in Revit for
the rest of the session, holding a strong reference to a closed window. The user sees typing latency
in every dialog, with no error anywhere.

**Fix:** store `self._unlock_input_handler = PreProcessInputEventHandler(self._on_unlock_input)` and
subtract that exact object; log the exception instead of swallowing it.

---

### M11. Hardcoded developer path shipped inside the runtime library

- `CEDLib.lib/UIClasses/Resources/Themes/theme_audit.py:8-14`

```python
ROOT = r"C:\Users\Aevelina\CED_Extensions\CEDLib.lib\UIClasses\Resources\Themes"
```

It also names `CEDTheme.DarkAlt.xaml`, which does not exist in the repo, so the audit reports an
all-blank DarkAlt column even on the original machine. This is a developer script living in a folder
pyRevit puts on `sys.path`.

Related: `CLAUDE_HANDOFF.md:219` pins
`C:\Users\Aevelina\AppData\Local\Programs\Python\Python310\python.exe` as the documented interpreter.

**Fix:** `ROOT = os.path.abspath(os.path.dirname(__file__))`, derive `THEME_FILES` from
`resource_loader.THEME_RELATIVE_PATHS`, and move the file to `tools/`.

---

### M12. Documentation viewer has no disable on the release path

- `.../CED Tools.panel/Docs.pushbutton/script.py:48-71`

No `__context__` guard, no feature flag, no `bundle.yaml` in the folder, and the folder is not
`_`-prefixed — pyRevit's two disable mechanisms. `main()` unconditionally constructs
`DocumentationViewerWindow`.

**Failure:** with the deployed `docs/user-guide/` tree or `catalog.json` absent,
`resolve_documentation_root` (`Documentation/pathing.py:58-60`) raises and the user gets
"The documentation viewer could not open" — a visibly broken button in a release build.

The viewer is otherwise well built: no file walk at startup (`catalog.py:80-90` reads one prebuilt
`catalog.json`), renderer import deferred to `Loaded` at background priority
(`viewer.py:185-199`), path containment enforced (`pathing.py:95-96`), catalog entries validated
(`catalog.py:72-75`), `.md`-only navigation (`viewer.py:459-460`, `:561-568`), non-http schemes
rejected (`:545-551`).

**Fix:** add `bundle.yaml` with the button disabled (or rename to `_Docs.pushbutton`). Minor:
`_inside`'s IronPython 2.7 fallback at `pathing.py:24` is case-sensitive, so a differently-cased
install path rejects valid pages — casefold on Windows.

---

### M13. Documentation search re-concatenates every page body on every keystroke

- `CEDLib.lib/Documentation/viewer.py:645-650`, `catalog.py:117-131`

```python
cached = (title, keywords, headings, metadata, content)
item["_search_fields"] = cached
title, keywords, headings, metadata, content = cached
searchable = " ".join((title, keywords, headings, metadata, content))
```

The tuple is cached but the joined string is rebuilt per document per call. `_search_changed` calls
`_refresh_results()` synchronously with no debounce (only the *page* highlight is debounced at
`:668-676`), and `_refresh_results` rebuilds a fresh `DataTable` plus three `TextBlock`s with
per-segment `Run` children for every result row (`:299-316`).

**Fix:** cache the joined string as `item["_search_blob"]`, and route `_search_changed` through the
same `DispatcherTimer` debounce already used for highlighting.

---

### M14. Modal `forms.alert` raised from inside `ExternalEvent.Execute`, unowned, with a TransactionGroup open

- `cg_apply.py:644-650`; `move_circuits_to_panel_service.py:616`, `:991`, `:1024`, `:1036`
- `calculate_circuits_operation.py:170`, `:175-181`, `:241`; `_print_staged_report` `output.show()` at `:1280-1285`

Raised from the API thread with no `Owner` set to the Revit main window, while
`run_clicked` has already called `self.Hide()` (`cg_window.py:2294`) and the compound operation's
`TransactionGroup` (e.g. `mark_existing_and_recalculate_operation.py:76`) is still live.

**Failure:** Run against a panel with no panel schedule view. The grouper hides itself, an unowned
WPF dialog may z-order behind the Revit main window, and Revit is blocked inside `Execute` waiting
on it — a frozen, unresponsive Revit with no visible dialog and no grouper. An unattended dialog also
holds the transaction group open indefinitely. Drives C1's cancel path.

**Fix:** hoist these decisions into the UI before raising the event (panel names are known at
plan-build time; the pane already prompts at `CircuitBrowserPanel.py:5501-5508` for >300 circuits).
At minimum set the dialogs' owner to the Revit main window handle, and disable the grouper rather
than hiding it.

---

### M15. Persistent `sys.path` mutation from pushbutton scripts

- `.../RevisionManager.pushbutton/script.py:53-58`
- `.../Tag by Example.pushbutton/script.py:18-20`

```python
if LIB_ROOT in sys.path:
    sys.path.remove(LIB_ROOT)
sys.path.insert(0, LIB_ROOT)
if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)
```

pyRevit's IronPython engine is shared and long-lived, so the reordering and the bundle directory at
index 0 persist for the whole session.

**Failure:** after Revision Manager runs once, `THIS_DIR` shadows every subsequent `import` in that
engine; a name collision surfaces as a bug in an unrelated tool.

**Fix:** append rather than insert, and prefix bundle-local module names
(`ced_revision_manager_core`), matching `_CIRCUIT_MANAGER_MODULE_NAME = "ced_electools_circuit_manager_panel"`
in `Theme.pushbutton/script.py:49`.

---

## Startup budget

Per Revit launch, blocking the UI thread, before the ribbon is usable. "Bad case" is a
responsive-but-slow Desktop Connector plus `CEDLib.lib` on a network path.

| Block | Where | Typical | Bad case |
|---|---|---|---|
| Imports, release metadata, pyRevit telemetry config | `AE pyTools/startup.py:1-1417` | 30–80 ms | 200 ms |
| **ACC route discovery, scoring, state write** | `telemetry_route.py:198-516` | 150–600 ms | 4–16 s |
| CEDLib.lib walk + full-content SHA-1 | `ElecTools/startup.py:63-83` | 30–80 ms | 1–3 s |
| Circuit Manager import graph (~11,000 lines) | `ElecTools/startup.py:184` | 250–700 ms | 1.5 s |
| Panel XAML + 19 resource dictionaries | `ElecTools/startup.py:196` | 300–900 ms | 1.5 s |
| **Total added to Revit startup** | | **0.8–2.4 s** | **8–22 s** |

**Confirmed clean:** no Revit DB query, no `FilteredElementCollector`, no network I/O and no
subprocess launch happens at startup in either `startup.py`. Nothing scales with model size. The
cost is entirely filesystem and XAML parsing, which is why it is all recoverable.

**Check first, cheap if true:** all three `extension.json` files
(`AE pyTools.extension/extension.json:1-19`, `CED ElecTools.extension/extension.json:1-18`,
`CED MechTools.extension/extension.json:1-18`) wrap their metadata in `{"extensions": [ {...} ]}` —
the schema of the root *catalog* file (`extensions.json`, which is correct as-is). pyRevit's
in-extension `extension.json` is a **flat** object. If pyRevit reads these as flat manifests it finds
no recognized keys and `"rocket_mode_compatible": "True"` is never applied. Rocket mode is the
largest pyRevit-side startup and per-command speedup available. Verify against the installed
pyRevit's parser, then flatten and use real JSON booleans rather than `"True"` strings.
*Unverified — pyRevit source is not in this repo.*

---

## What makes a tool window slow to open, ranked

1. **Parsing 19 XAML resource dictionaries from disk, uncached, per window** (C5). ~200 KB, 96 styles,
   55 control templates, every time, used or not. Fixed cost that scales with the design system
   rather than the tool — so it grew exactly as the shared UI work landed. Fixing this alone recovers
   most of the penalty and requires no tool changes.
2. **Non-virtualizing ComboBox popups** (C6). Does not slow the window open, but freezes the first
   dropdown and every keystroke after it.
3. **FilterableComboBox construction and per-keystroke highlight refresh** (H11). When a tool builds
   combos at window load — one per DataGrid row in Let there be Light — the constructor chain
   (`filterable_combo_box.py:290-298`) lands on the open path once per row.
4. **Theme resolution per window** — a full `pyRevit_config.ini` read and parse plus three
   `script.get_config` calls, and a config *write* on first run inside window construction.
   `ui_bases.py:41-84`, `:252-284`, `:334-338`; `theme_manager.py:210-233`.
5. **Redundant path resolution** — `resolve_ui_context` walks ancestors twice (`pathing.py:134-157`,
   `:82`) and `prioritize_syspath` runs `os.path.realpath` on every `sys.path` entry (`:20`). Small
   locally, meaningful on a file share.
6. **Tool-specific eager collection** (M8).

---

## Suggested order

Ordered by benefit per unit of risk, not severity alone.

1. **Verify the `extension.json` schema.** Ten minutes. If rocket mode is not being applied, this
   beats every other performance item here.
2. **Fix the compound-operation rollback condition (C1).** Three files, one inverted condition each.
   The only finding that silently corrupts a model the user believes is untouched.
3. **Cache or app-level-merge the base resource dictionaries (C5).** One file, no tool changes,
   recovers most of the per-window launch penalty. Then trim the default list.
4. **Restore virtualization in the three ComboBox templates (C6).** Three lines of XAML.
5. **Guard the two settings transactions (C2) and add read-only/storage-type checks to the writers (H1).**
6. **Take ACC route discovery off startup (C4) and make the Circuit Manager pane lazy (H6).**
7. **Bring Circuit Grouper and Alerts Manager up to the Wire Tools modeless contract (C3, H4, H5).**
8. **Point `category_id_value` at `coerce_elementid_value` and delete the three workarounds (M1).**

---

## Coverage and caveats

**Not reviewed** (several findings note where a conclusion depends on one of these):
`Snippets/_elecutils.py`, `Snippets/circuit_ui_actions.py`, `Snippets/design_options.py`,
`Snippets/tag_geometry.py`, `Tag by Example/lib/tag_api_compat.py`,
`Tag by Example/lib/tag_host_adapters.py`,
`CEDElectrical/Infrastructure/Revit/repositories/panel_schedule_repository.py`,
`.../distribution_equipment_repository.py`,
`CEDElectrical/Infrastructure/Revit/external_events/circuit_operation_event.py`, and the
`AE pyTools.extension/hooks/` directory.

**pyRevit hooks are worth a separate pass** — they fire on every command and document event, so
anything in there adds latency to every button press.

**Repo weight:** `CEDLib.lib/OLD CHECKPOINTS/` holds ~60 MB of YAML prototype snapshots, and the git
object store carries packs of 237 MB, 142 MB and 102 MB. None of it is parsed at startup, but it all
ships in the pyRevit library path and the MSI. Worth moving out of the deployed tree.

**Open questions flagged but not confirmed** (need a live Revit to settle):

- `electrical_qc_service.py:1159` and `:420` call `int()` on values that may be native
  `DB.ElementId` if `distribution_equipment_repository` returns them that way — would raise
  `TypeError` in IronPython and kill the transformer-secondary-length QC check. Use
  `coerce_elementid_value`.
- `wire_tools_logic.py:646-653`, `:711-718`, `:739-746` pass `system_type=None` to
  `_elecutils.get_circuits_from_selection` / `filter_circuits`. This is only correct if `None` means
  "no filter" there; in `wireutils.get_wire_circuit_id` (`Snippets/wireutils.py:171`, `:203`) `None`
  defaults to `POWER_CIRCUIT_SYSTEM_TYPE`. If the two disagree, the non-power-system-types fix is
  incomplete. Worth a unit test on `filter_circuits(systems, system_type=None)`.
- `Snippets/tag_host_transform.py:127-134`, `:184-203`, `:216-223` — `basis_x` and `basis_y` are
  projected and normalized independently, so they are orthogonal only for a pure rotation about Z.
  For a face-based device on a sloped surface the round trip skews the copied tag-head offset;
  `adapt_target_frame` negating `basis_x` alone produces a left-handed frame. Re-orthogonalize with
  `axis_y = normal.CrossProduct(axis_x)`.
- Tag by Example cross-view copying may mirror the tag-head Z: the example frame is built against
  `owner_view` (`tag_by_example_events.py:112`) and the target frame against the active view
  (`:246`); a floor plan and an RCP have opposite `ViewDirection`. Test: copy a tag from a plan to an
  RCP over the same host and compare `TagHeadPosition.Z`.

**Two things done well, worth defending in review:**

1. The Circuit Manager panel defers every Revit DB query to first show — there is no
   model-size-dependent work at startup anywhere in this repo.
2. `Wire Tools.pushbutton/lib/wire_tools_events.py` is a correct, complete modeless implementation:
   a stable `PathName|Title` document key (`:63`), `DocumentClosing` + `doc-changed` +
   `view-activated` subscriptions with symmetric `detach_lifecycle` (`:166-206`), context
   re-validation at the top of every `Execute` (`:348-357`), stale-payload discard on view change
   (`:246`), and teardown wired to `Window.Closing` (`wire_tools_ui.py:962`). It should be the
   template every other modeless tool is measured against.
