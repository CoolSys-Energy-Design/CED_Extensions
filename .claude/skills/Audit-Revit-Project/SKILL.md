---
name: Audit-Revit-Project
description: Read-only QA/QC audit of ANY Revit project (program-agnostic) via Revit MCP, with optional standards comparison against a known-good reference project. Prompts for the Suspect (audited) project and an optional Good (reference) project, lets the user add custom QC checks per run, and delivers PDF reports with screenshots. Use when Reed asks to audit or QC a Revit model/project that is not covered by a program-specific audit skill (HEB projects have their own Audit-HEB-Projects skill).
---

# Audit-Revit-Project

Read-only audit of a "Suspect" Revit project + optional standards comparison against a "Good"
reference project. Program-agnostic: no client-specific rules live in the numbered checks —
those belong in **Program Rules** at the bottom (or in a dedicated program skill, like
Audit-HEB-Projects).

Everything is READ-ONLY: never start transactions, never save, never sync either model.

## Phase 0 — Setup prompts (ALWAYS, in this order)

1. **Prompt for the SUSPECT project** (the one being audited). Use AskUserQuestion. Offer:
   - A model currently open in Revit (list open docs via MCP first if Revit is reachable)
   - A model the user will open manually
   The Suspect is usually the active document.
2. **Prompt for the GOOD project** (the reference/standard). Same options, plus:
   - Skip comparison (Part 1 audit only) — the default for one-off model QC.
3. **Prompt for the report save location.** The DEFAULT (recommended option) is the shared Teams
   library folder, built portably from the current user's profile — never hardcode a username:
   `%USERPROFILE%\CoolSys Inc\Teams-Coolsys - Tool Development - Documents\MEP AUTOMATION\Project Audits`
   (resolve `%USERPROFILE%` / `$env:USERPROFILE` at runtime). Offer it as the default; the user
   may override by pasting a different full folder path in chat. Before Phase 1:
   - Verify the parent library (`...\MEP AUTOMATION`) exists — it only exists if this user has
     synced the Teams-Coolsys Tool Development library. Create the `Project Audits` leaf if
     missing.
   - If the library is NOT synced on this machine, say so and fall back to a pasted path or the
     Desktop (check for OneDrive redirect: `%USERPROFILE%\OneDrive - CoolSys Inc\Desktop`).
4. **Prompt for which check groups to run.** Multi-select: Model Health / Electrical /
   Documentation & Spelling / Cross-Trade Clash / Naming Conventions. Default: all that apply
   (skip Electrical automatically if the model has no electrical categories populated). Also ask
   "Any additional QC checks to add for this run?" (free-text via Other). If the user adds any:
   - Run them this session as additional audit steps, AND
   - Ask per check: "Save this check into the skill for all future runs?" If yes — is the check
     universal (any project, any client)? Append to **Custom QC Checks**. Is it specific to one
     client/program? Append under that program's heading in **Program Rules** instead, and
     suggest a dedicated Audit-<Program>-Projects skill once a program accumulates 3+ rules.
5. **Both models must be open in the same Revit session before extraction begins.**
   NEVER open a document via `app.OpenDocumentFile` from MCP without explicitly warning the user
   first — it freezes the entire Revit UI for the duration (20+ min for a full store model) and
   blocks the MCP queue. The strongly preferred path: ask the user to open both models through
   the Revit UI, then verify with `get_revit_status` + a doc-list query.
   Cloud models CANNOT be opened detached via API ("Detach option is not valid for cloud model"),
   and file copies of cloud models retain cloud identity and refuse detach too.

## Phase 1 — Suspect-project audit (the "Part 1" report)

Run each check with `execute_revit_code` (IronPython 2). Dump all results as files to the
scratchpad; never rely on the 60-second HTTP response for long operations (see Runtime Rules).
Every check emits findings in a common shape — `{check, rule, element_id, severity, detail,
sheet_or_view}` — so the report builder treats all checks uniformly.

### Group A — Sheet inventory & model health

1. **Sheet inventory** — all `ViewSheet` (skip placeholders): number, name. Classify disciplines
   by number prefix (E-, M-, P-, R-, FP, FA, SEC, MEP, X0 — report whatever prefixes exist; do
   not assume a fixed set).
2. **Warnings** — total count, grouped by warning description, top 10 types with element IDs.
   If a previous audit of the same model is found in the save location, report the delta.
3. **Model hygiene** — in-place families (count + names); imported (not linked) CAD; DWGs in
   views vs model; unplaced/unenclosed/redundant rooms and spaces; linked models and their
   load status; views not placed on any sheet (grouped by view type, excluding view templates,
   schedules, and legends); design options and worksets inventory (report, don't judge).

### Group B — Electrical (skip cleanly if not an electrical model)

4. **Uncircuited items on electrical sheets** — for every view placed on E-* sheets, collect
   FamilyInstances of **Electrical Fixtures, Electrical Equipment, Mechanical Control Devices**.
   Flag any with ≥1 power connector and no circuit. Report **element ID + family + type + the
   sheets it appears on** for every flagged item.
   - Circuit membership: `mep.GetElectricalSystems()` — NEVER `GetAssignedElectricalSystems()`
     (that returns panel-fed systems only and falsely flags everything).
   - Power connector filter: `c.Domain == DomainElectrical and c.ElectricalSystemType in
     {PowerCircuit, PowerBalanced, PowerUnBalanced}` (capital B in UnBalanced, Revit 2024).
5. **Circuit integrity** — walk every `ElectricalSystem`:
   - Circuits with no panel (`BaseEquipment is None`)
   - Circuits loaded past their rating (`ApparentCurrent` converted to amps > `Rating`)
   - Duplicate panel names across Electrical Equipment (`RBS_ELEC_PANEL_NAME`)
   (Reference implementation with 2024/2025/2026 API fallbacks:
   `C:\Users\reed.pinterich\.claude\routines\heb-qaqc\checks\electrical_check.py`.)
6. **User wire/ground sizing overrides** — CED circuit params, when present in the project:
   `CKT_User Override_CED == 1`. Report panel, ckt, load name, hot/ground/neutral sizes, and the
   `Wire Size_CEDT` callout for each. If the CED shared parameters are absent, report the check
   as N/A (non-CED template), not as a pass.
7. **Circuit Manager last calculation** — when the project uses CED Circuit Manager: parse the
   JSON blob in each circuit's `Circuit Data_CED` text parameter; `last_calculation` is a UTC
   ISO timestamp. Report the latest run plus the full list of stamped circuits, and say plainly
   that unstamped circuits predate the feature. N/A if the parameter doesn't exist.

### Group C — Documentation & spelling

8. **Spelling/grammar sweep of ALL sheets** — collect per sheet: sheet name, text notes placed
   on the sheet, view names + text notes in every placed view, schedule titles
   (`ScheduleSheetInstance`), PLUS all circuit load names (they print on panel schedules).
   Build a unique-word list; review it for misspellings (read the whole list — real errors like
   CONDENSOR/DECTECTOR/EXHUAST hide among valid abbreviations). Also scan for doubled words
   (`\b(THE|AND|OF|TO|...)\s+\1\b`) and run-together words (missing spaces). Map every finding
   back to its exact source string and sheet(s). Separate: definite misspellings / abbreviation
   inconsistencies / run-together text (verify visually — a hard line break can sit at the seam).
9. **Untagged equipment on printed views** — for each discipline's plan views placed on sheets,
   flag major equipment (Electrical Equipment, Mechanical Equipment, Plumbing Fixtures where
   applicable) with no tag in any sheet-placed view. Report by element + views checked.
10. **Title block state** — check for stale review stamps ("100% REVIEW — NOT FOR
    CONSTRUCTION", "PRELIMINARY", "NOT FOR CONSTRUCTION") visible in sheet exports; check
    revision/date fields populated on every sheet.

### Group D — Cross-trade clash scan (optional; ask before running on very large models)

11. **In-document category-pair clash check** — bounding-box prefilter
    (`BoundingBoxIntersectsFilter`) then true solid intersection
    (`ElementIntersectsElementFilter`) on the survivors. Default pairs: Pipes×Ducts,
    Pipes×Cable Tray, Ducts×Structural Framing; the user can add pairs at the Phase 0 prompt.
    Report each clash with both element IDs, families, systems, level, and location point.
    (Reference implementation: `C:\Users\reed.pinterich\.claude\routines\heb-qaqc\checks\clash_check.py` —
    pairs are defined at the top; long-running, use the file-watch pattern from Runtime Rules.)
12. **Host-vs-LINKED-model clash check** — `ElementIntersectsElementFilter` cannot cross
    documents, so links need a different technique: for each loaded `RevitLinkInstance`,
    transform each link element's bounding box into host coordinates (all 8 corners through
    `GetTotalTransform` — Min/Max alone is wrong under rotation), prefilter host elements with
    `BoundingBoxIntersectsFilter`, then transform the link element's solids with
    `SolidUtils.CreateTransformed` and test the candidates with `ElementIntersectsSolidFilter`.
    Default pairs: host Pipes/Ducts/Cable Tray × linked Structural Framing/Columns. Linked
    walls/floors are OFF by default (intentional penetrations make them noisy) — add per run
    only if asked. Report unloaded links as NOT checked (cross-reference the Group A link
    status finding). Link-vs-link is out of scope.
    (Ready-to-run implementation: `scripts\clash_check_linked.py` in this skill folder — set
    `OUT_PATH`, optionally `PAIRS`/`LINK_NAME_FILTER`, send via `execute_revit_code`, and poll
    for the output file per Runtime Rules.)
    **Interpreting results — open-web joist noise (validated on Tomball 2026-08-26: 1,364 of
    1,900 raw clashes were sprinkler branch lines crossing K-series bar joists):** joist
    families are modeled as simplified solids, so pipes legitimately routed between joist webs
    read as clashes. Split the report: crossings whose structural member name matches
    `\b(EXIST\.?\s*)?\d+(K|KCS|G|LH|DLH)\d*` (K-series/girder joists) go in a one-line summary
    count per system, NOT itemized; clashes with solid members (wide-flange beams like
    "16 x 24", columns, girders, framing without a joist designation) are the real findings and
    get the full table + screenshots. A pipe DIPPING BELOW the joist bottom chord or running
    along (not across) a joist is real even in the joist bucket — flag crossings where the pipe
    axis is within ~15 degrees of the joist axis.

### Group E — Naming conventions

13. **Naming-standard scan** — driven entirely by the **Program Rules** section below. If a
    program heading matching this project exists, apply its regex rules to view names, sheet
    names/numbers, family names, worksets, and panel names. If NO rules exist for this project's
    program, do not invent standards: instead report the *observed* patterns (dominant prefix
    conventions, outliers like lowercase ad-hoc names, accidental duplicates with `1`/`.0001`
    suffixes) so the user can decide what the standard should be — and offer to save it as a
    Program Rule.

## Phase 2 — Standards comparison (the "Part 2" report; skip if no Good project)

Extract the same datasets from the Good project, then diff locally (Python 3 on the dumped files):

1. **Sheet set** — numbers only in one project; same number with different name; naming-convention
   differences (e.g. partial-plan "- A/B/C/D" suffixes). Distinguish real standards drift from
   expected building differences and list those separately as verified-expected.
2. **Vocabulary diff** — unique words per project; pair near-matches (edit distance ≤2) between
   the two unique sets to catch variant spellings of the same term (ACOUSTICAL/ACCOUSTICAL,
   TANDEM/TANDUM). Shared note blocks drifting apart (typo in only one project) is a key finding.
3. **Load-name conventions** — count abbreviation variants (REC/RECEPT/RECEPTACLE/RCPT/RECPT,
   LTG/LIGHTING, etc.) in both projects' circuit load names. Report the variant table; only
   flag a "correct" form if a Program Rule defines one.
4. **Family/type naming** — dump `FamilySymbol` inventories per category from both; list
   families in one-not-the-other for key categories; flag naming-standard violations (lowercase
   ad-hoc names, missing standard prefixes, accidental duplicates with `1`/`.0001` suffixes).
5. **Panel-schedule layout** — map panels to sheets via `PanelScheduleSheetInstance.OwnerViewId`
   model-wide (per-sheet collectors are unusably slow). Identical layout across projects is a
   conformance win worth reporting.
6. **Warning-count comparison** — Suspect vs Good totals and top warning types; a Suspect type
   absent from the Good model is a stronger signal than raw counts.

## Phase 3 — Screenshot evidence

For each headline finding, export the sheet (from BOTH models when comparing) and present
side-by-side:
- `ImageExportOptions`, `ExportRange.SetOfViews`, `ZoomFitType.FitToPage`; PixelSize 3000 for
  overview, 8000 when a crop must be readable. Works on background (non-active) documents.
- Crop evidence regions with Pillow (Python 3, `PIL.Image.crop`). View the 3000px export first
  to locate the region, scale coordinates to the 8000px image.

## Phase 4 — Deliverables

PDFs saved to the folder resolved in Phase 0 step 3 (default: the shared Teams library
`%USERPROFILE%\CoolSys Inc\Teams-Coolsys - Tool Development - Documents\MEP AUTOMATION\Project Audits`):
- **"<YYYYMMDD> - <Suspect> Audit Part 1.pdf"** — suspect-project audit with full ID/sheet
  tables. Date prefix is the run date, e.g. "20260826 - Store 214 Audit Part 1.pdf".
- **"<YYYYMMDD> - <Suspect> Audit Part 2.pdf"** — standards comparison with side-by-side
  screenshots and a separate expected-building-differences table. (Only when a Good project
  was compared.)
Build as HTML (tables, red/green highlights, embedded `file:///` images downscaled to ~1500px)
and convert: `msedge.exe --headless --disable-gpu --no-pdf-header-footer --print-to-pdf=...`.
Verify element/circuit numbers and quoted strings against the dumped JSON before citing them.

**Reporting brevity**: when a check passes, do NOT enumerate every correct item — one line
stating the check passed (with the count checked, e.g. "All 62 circuits within rating — PASS")
is enough; move straight to the next check. Detailed tables are reserved for findings/failures.
Applies to both the chat summary and the PDF reports.

## Runtime rules (hard-won — do not relearn these)

- **IronPython 2** in `execute_revit_code`: `print` statement, `except Exception, ex:`, no
  f-strings. `System` must be imported before use.
- **Sanitize non-ASCII before json.dumps/write**: drawing text contains Ø (0xD8), ° (0xB0),
  ¼½¾ (0xBC-BE) as raw bytes in .NET strings — they crash json/codecs with UnicodeDecodeError
  and the output file is left 0 bytes. Map them to (DIA)/(DEG)/fractions, everything else >126
  to `?`. Apply to EVERY string leaving Revit (text notes, family names, schedule names).
- **Long operations**: the MCP HTTP call times out at 60 s but the code keeps running. Write
  results to a scratchpad file (flush per item, JSONL for sweeps), end with a done-marker line,
  and poll for the marker from Bash. Errors after timeout are silent — wrap everything and write
  the exception into the marker file.
- **Version tolerance (Revit 2024/2025/2026)**: ElementId `.Value` with `.IntegerValue`
  fallback; `GetElectricalSystems()` with deprecated `ElectricalSystems` property fallback;
  guard unit conversions; try/except every section into an `errors` list — a nonzero error
  count on one Revit version usually means an API changed, fix the script.
- **`Element.Name` is hidden on FamilySymbol/ElementType**: read via
  `DB.Element.Name.__get__(sym)`.
- **Transactions**: this skill is read-only — never call `Transaction.Start()` yourself under
  any circumstances. Note: MCP wrapper behavior has differed across versions (some auto-wrap
  snippets in a transaction, some don't) — the read-only guarantee is that the check code
  itself never modifies the model, so no borrow/sync can occur either way. `uidoc.ActiveView =
  view` can throw inside a call — don't switch views from check code.
- Address each model explicitly by iterating `app.Documents` and matching `Title` — never assume
  which document is active.
- Poll Revit progress from PowerShell when it looks hung: two `Get-Process Revit` CPU samples
  ~15 s apart; rising CPU = still working. WorkingSet collapse just means Windows trimmed it
  (check commit charge before assuming a leak).

## Program Rules (client/program-specific standards; append under a program heading)

Rules here apply ONLY when the Suspect project belongs to that program (match by project name,
template, or ask the user). Keep each rule in the numbered format used by Custom QC Checks.
When a program accumulates 3+ rules, consider promoting it to a dedicated
`Audit-<Program>-Projects` skill (HEB already has one — do not duplicate its rules here).

*(none yet)*

## Custom QC Checks (universal, user-added; append new checks here, numbered)

*(none yet)*
