---
name: Audit-HEB-Projects
description: Read-only QA/QC audit of an HEB Revit project via Revit MCP, plus a standards comparison against a known-good reference project. Prompts for the Good (reference) project and the Suspect (audited) project, lets the user add custom QC checks per run, and delivers PDF reports with screenshots to the Desktop. Use when Reed asks to audit an HEB project, compare project standards, or QC sheets/circuits.
---

# Audit-HEB-Projects

Read-only audit of a "Suspect" HEB project + standards comparison against a "Good" reference project.
Everything is READ-ONLY: never start transactions, never save, never sync either model.

## Phase 0 — Setup prompts (ALWAYS, in this order)

1. **Prompt for the GOOD project** (the reference/standard). Use AskUserQuestion. Offer:
   - A model currently open in Revit (list open docs via MCP first if Revit is reachable)
   - A model the user will open manually
   - Skip comparison (Part 1 audit only)
2. **Prompt for the SUSPECT project** (the one being audited). Same options. The Suspect is
   usually the active document.
3. **Prompt for the report save location.** The DEFAULT (recommended option) is the shared Teams
   library folder, built portably from the current user's profile — never hardcode a username:
   `%USERPROFILE%\CoolSys Inc\Teams-Coolsys - Tool Development - Documents\MEP AUTOMATION\HEB Project Audits`
   (resolve `%USERPROFILE%` / `$env:USERPROFILE` at runtime). Offer it as the default; the user
   may override by pasting a different full folder path in chat. Before Phase 1:
   - Verify the parent library (`...\MEP AUTOMATION`) exists — it only exists if this user has
     synced the Teams-Coolsys Tool Development library. Create the `HEB Project Audits` leaf if
     missing.
   - If the library is NOT synced on this machine, say so and fall back to a pasted path or the
     Desktop (check for OneDrive redirect: `%USERPROFILE%\OneDrive - CoolSys Inc\Desktop`).
4. **Prompt for extra QC checks**: ask "Any additional QC checks to add for this run?" (free-text
   via Other). If the user adds any:
   - Run them this session as additional audit steps, AND
   - Ask per check: "Save this check into the skill for all future runs?" If yes, append it to the
     **Custom QC Checks** section at the bottom of this file via Edit (keep the numbered format).
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

1. **Sheet inventory** — all `ViewSheet` (skip placeholders): number, name. Classify disciplines
   by number prefix (E-, M-, P-, R-, FP, FA, SEC, MEP, X0).
2. **Uncircuited items on electrical sheets** — for every view placed on E-* sheets, collect
   FamilyInstances of **Electrical Fixtures, Electrical Equipment, Mechanical Control Devices**.
   Flag any with ≥1 power connector and no circuit. Report **element ID + family + type + the
   sheets it appears on** for every flagged item.
   - Circuit membership: `mep.GetElectricalSystems()` — NEVER `GetAssignedElectricalSystems()`
     (that returns panel-fed systems only and falsely flags everything).
   - Power connector filter: `c.Domain == DomainElectrical and c.ElectricalSystemType in
     {PowerCircuit, PowerBalanced, PowerUnBalanced}` (capital B in UnBalanced, Revit 2024).
3. **Case-controller neutral check** — circuits serving Mechanical Control Devices must list a
   neutral: `CKT_Include Neutral_CED == 1` and `CKT_Wire Neutral Size_CEDT` not blank/`-`.
   Report panel/ckt/load name/wire callout for every circuit, PASS/FAIL.
4. **Circuit Manager last calculation** — parse the JSON blob in each circuit's
   `Circuit Data_CED` text parameter; `last_calculation` is a UTC ISO timestamp. Only circuits
   recalculated since the timestamp feature shipped carry one — report the latest run plus the
   full list of stamped circuits, and say plainly that unstamped circuits predate the feature.
5. **User wire/ground sizing overrides** — `CKT_User Override_CED == 1`. Report panel, ckt,
   load name, hot/ground/neutral sizes, and the `Wire Size_CEDT` callout for each.
6. **Spelling/grammar sweep of ALL sheets** — collect per sheet: sheet name, text notes placed
   on the sheet, view names + text notes in every placed view, schedule titles
   (`ScheduleSheetInstance`), PLUS all circuit load names (they print on panel schedules).
   Build a unique-word list; review it for misspellings (read the whole list — real errors like
   CONDENSOR/DECTECTOR/EXHUAST hide among valid abbreviations). Also scan for doubled words
   (`\b(THE|AND|OF|TO|...)\s+\1\b`) and run-together words (missing spaces). Map every finding
   back to its exact source string and sheet(s). Separate: definite misspellings / abbreviation
   inconsistencies / run-together text (verify visually — a hard line break can sit at the seam).

## Phase 2 — Standards comparison (the "Part 2" report; skip if no Good project)

Extract the same datasets from the Good project, then diff locally (Python 3 on the dumped files):

1. **Sheet set** — numbers only in one project; same number with different name; naming-convention
   differences (e.g. partial-plan "- A/B/C/D" suffixes). Distinguish real standards drift from
   expected building differences (mezzanine, tenant space, canopy, generator vs fire pump,
   hydronic vs DX refrigeration) and list those separately as verified-expected.
2. **Vocabulary diff** — unique words per project; pair near-matches (edit distance ≤2) between
   the two unique sets to catch variant spellings of the same term (ACOUSTICAL/ACCOUSTICAL,
   TANDEM/TANDUM). Shared note blocks drifting apart (typo in only one project) is a key finding.
3. **Equipment terminology** — grep both text sets for key equipment phrases (twist tie, case
   power, gondola, checkstand, anti-sweat, air curtain, ...) and compare exact naming.
4. **Load-name conventions** — count abbreviation variants (REC/RECEPT/RECEPTACLE/RCPT/RECPT,
   LTG/LIGHTING, etc.) in both projects' circuit load names.
5. **Family/type naming** — dump `FamilySymbol` inventories per category from both; list
   families in one-not-the-other for key categories; flag naming-standard violations (lowercase
   ad-hoc names, missing CED prefixes, accidental duplicates with `1`/`.0001` suffixes).
6. **Panel-schedule layout** — map panels to sheets via `PanelScheduleSheetInstance.OwnerViewId`
   model-wide (per-sheet collectors are unusably slow). Identical layout across projects is a
   conformance win worth reporting.
7. **Title block state** — check for stale review stamps ("100% REVIEW — NOT FOR CONSTRUCTION")
   visible in sheet exports.

## Phase 3 — Screenshot evidence

For each headline finding, export the sheet from BOTH models and present side-by-side:
- `ImageExportOptions`, `ExportRange.SetOfViews`, `ZoomFitType.FitToPage`; PixelSize 3000 for
  overview, 8000 when a crop must be readable. Works on background (non-active) documents.
- Crop evidence regions with Pillow (Python 3, `PIL.Image.crop`). View the 3000px export first
  to locate the region, scale coordinates to the 8000px image.

## Phase 4 — Deliverables

PDFs saved to the folder resolved in Phase 0 step 3 (default: the shared Teams library
`%USERPROFILE%\CoolSys Inc\Teams-Coolsys - Tool Development - Documents\MEP AUTOMATION\HEB Project Audits`):
- **"<Suspect> Audit Part 1.pdf"** — suspect-project audit with full ID/sheet tables.
- **"<Suspect> Audit Part 2.pdf"** — standards comparison with side-by-side screenshots and a
  separate expected-building-differences table.
Build as HTML (tables, red/green highlights, embedded `file:///` images downscaled to ~1500px)
and convert: `msedge.exe --headless --disable-gpu --no-pdf-header-footer --print-to-pdf=...`.
Verify circuit numbers and quoted strings against the dumped JSON before citing them.

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
- **`Element.Name` is hidden on FamilySymbol/ElementType**: read via
  `DB.Element.Name.__get__(sym)`.
- **The MCP wrapper auto-opens a Transaction** around every snippet — never `Transaction.Start()`
  yourself, and `uidoc.ActiveView = view` always throws inside a call.
- Address each model explicitly by iterating `app.Documents` and matching `Title` — never assume
  which document is active.
- Poll Revit progress from PowerShell when it looks hung: two `Get-Process Revit` CPU samples
  ~15 s apart; rising CPU = still working. WorkingSet collapse just means Windows trimmed it
  (check commit charge before assuming a leak).

## Custom QC Checks (user-added; append new checks here, numbered)

1. **Receptacle abbreviation standard is "RECS" — MANDATORY every run** (added by Reed
   2026-08-07). Sweep every circuit load name AND every sheet text string in the Suspect project
   for receptacle references. The ONLY acceptable form is `RECS`. Flag EVERY instance of any
   other variant — `REC`, `RECEPT`, `RECEPTS`, `RCPT`, `RCPTS`, `RECPT`, `RECPTS`, `RECEPTACLE`,
   `RECEPTACLES`, `RECEP`, `RECEPS` — with panel/circuit number (for load names) or sheet number
   (for sheet text), the full source string, and the suggested `RECS` replacement. Report as its
   own section with a total count per variant. Regex guide: match word-boundary tokens
   `\b(REC|RECEPTS?|RECEPS?|RCPTS?|RECPTS?|RECEPTACLES?)\b` case-insensitive; exclude legitimate
   non-receptacle words (RECEIVING, RECESSED, RECORD, RECOVERY, RECIRCULATING, RECYCL*).
   RECEP/RECEPS were found in the wild on Buda 2026-08-07 — do not narrow the pattern.

2. **Fire Protection set — spelling & consistency** (added by Reed 2026-08-14). The FP-* sheets
   get their own dedicated spelling/inconsistency pass in addition to the all-sheets sweep:
   collect every text string on FP-* sheets (notes, view titles, schedules, legends), review the
   word list, and report FP findings as their own section.

3. **FP pipe schedule — ONLY Schedule 30 or Schedule 40 (HARD RULE)** (Reed 2026-08-14). Sweep
   all FP sheet text (general notes especially) for pipe schedule references: `SCH`, `SCHED`,
   `SCHEDULE` followed by a number, and bare `SCH.XX` forms. The ONLY acceptable schedules are
   **30 and 40**. FLAG EVERY instance of any other schedule number (10, 5, 80, etc.) with sheet
   and full source string. Also flag Sch 30/40 statements that contradict each other (e.g. one
   note assigning Sch 30 to a size range another note assigns Sch 40). Scope: fire protection
   sheets/notes — plumbing sheets legitimately use other schedules (e.g. Sch 80 PVC), so limit
   the hard rule to FP-* sheets and FP general notes.

4. **FP sprinkler schedule — types & K-factors** (Reed 2026-08-14). Extract the sprinkler
   schedule (FP legend/schedule sheets): every head listed with its type, orientation, response,
   temperature, and K-factor. Verify each K-factor is consistent between schedule, notes, and
   design criteria table. Report the schedule as a table with any mismatches flagged.

5. **FP design criteria table — NFPA references & code year** (Reed 2026-08-14). Locate the
   design criteria table; verify the NFPA 13 chapter/section references exist and match the
   cited NFPA edition year, and that the code year is consistent everywhere it appears in the FP
   set (criteria table, general notes, cover sheet code summary). Flag chapter/section numbers
   that don't exist in the cited edition (NFPA 13 renumbered heavily in 2019+ editions — a 2016
   section number cited against a 2019+ edition is a classic miss).

6. **FP sprinkler applications by area** (Reed 2026-08-14). Verify via sheet text + plan
   symbols/tags: **Wareroom/Receiving areas use EC-25.2** extended-coverage heads;
   **Water/Paper Gondola areas use K-16.8** heads. Flag any other head type tagged in those
   areas, and flag if the schedule lists EC-25.2 / K-16.8 but the plans tag something else.

7. **FP sprinkler elevations vs RCP** (Reed 2026-08-14). Verify sprinkler elevations are shown
   where required and coordinated with reflected-ceiling-plan areas: extract elevation callouts
   from FP plans and compare against ceiling heights in the RCP/architectural data where
   readable; screenshot any area where heads have no elevation note or the elevation conflicts
   with the ceiling. This check is part-visual: export the FP plan sheets and review.

8. **FP freezer/cooler areas** (Reed 2026-08-14). For every freezer/cooler footprint: verify
   sprinkler elevations AND locations are shown (heads below insulated ceiling, correct
   clearances). Dry-type or listed heads expected — flag wet heads inside freezers. Screenshot
   each freezer/cooler FP area for the report.

9. **FP dry system — TXBY & curbside canopy** (Reed 2026-08-14). Verify the dry-pipe system
   scope covers the TXBY (Texas BBQ) and curbside canopy areas: dry compressor identified
   (location + circuit — cross-check an electrical circuit exists for it), dry valve shown, and
   the dry-system boundary matches the unheated areas. Flag wet pipe shown in unheated space.

10. **Fire pump verification** (Reed 2026-08-14). From FP-2.x (fire pump sheets) + one-line +
    schedules, extract and report: pump rating (GPM @ PSI), motor horsepower, **starter type —
    must be SOFT START**, and ATS requirements/usage (is the fire pump on an ATS, is it shown on
    the electrical one-line, does the controller listing match). Cross-check HP against the
    electrical circuit/feeder serving the pump controller. Flag any mismatch between FP sheets
    and electrical sheets.

11. **FP dry valve detail** (Reed 2026-08-14). Locate the dry valve detail; verify it is
    present, complete (trim, air supply, priming, drain, alarm connections labeled), and
    consistent with the system design (valve size/model matches riser diagram and schedule).
    Screenshot the detail for the report; flag missing components or model mismatches.
