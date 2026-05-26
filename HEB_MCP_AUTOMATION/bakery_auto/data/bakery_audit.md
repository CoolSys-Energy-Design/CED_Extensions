# HEB BAKERY Replication Audit

Generated: 2026-05-19 (post-pipeline; both areas finished ~16:14)
Source view: `Power Callout - BAKERY - L1`
Target doc: `CED_HEB Template_MEPR_R24_OkmntProfCorr` (verified via Revit MCP `get_revit_status` / `get_current_view_info`)
Mirrors the prior pharmacy audit methodology.

---

## 1. Executive summary

The bakery replication is **clean**. Every element the pipeline is *designed* to place was placed
with zero failures across all categories:

- 75/75 source Electrical Fixtures placed (`status:"matched"` for all, 0 anchor-missing, 0 ambiguous).
- 53/53 in-scope circuits created (`made=53, failed=0`).
- All supporting annotation collections placed 1:1: keynotes 23/23, wires 74/74, textnotes 16/16,
  fixture_tags 42/42, wire_tags 60/60 (all reports `skipped=0`, `removed=0`).

The only deltas are **expected-by-design exclusions**, not bugs:

1. 3 source panelboards (`Electrical Equipment`) intentionally not placed.
2. ~99 source circuits out of scope because their devices/panels are outside the bakery callout
   (only panel `BA`, 53 circuits, is in scope).

**One real risk — not visible in the data files:** the just-edited
`place_circuits.set_rating_poles` (skill mtime **16:31**) was applied by a circuiting run whose
report was written at **16:13** — i.e. the current target circuits were created by code that
ran *before* the latest edit. Whether the in-model breakers carry correct ratings or the old
20 A / 1-pole default **cannot be determined from the JSON** and requires an in-Revit re-run +
verification. See §5.

---

## 2. Side-by-side inventory

| Item | Source | Placed | Delta | Explanation (code citation) |
|---|---:|---:|---:|---|
| **Electrical Fixtures** (total) | 75 | 75 | 0 | All `status=="matched"`; placed when `category=="Electrical Fixtures"`. `place_relative.py:87` `if r["category"] != "Electrical Fixtures": continue`; placement gate `place_relative.py:299` `if x["status"] != "matched": skipped += 1; continue`. |
| **Electrical Equipment** (panelboards) | 3 | 0 | **-3** | **Expected-by-design.** Excluded by the same `place_relative.py:87` filter (only `"Electrical Fixtures"` enters `rows`). IDs `7855685` (`zPanelboard - 208V/3Ph`, Mark 24340610), `7855687` (`...IG Bus`, 24340612), `10473209` (`zPanelboard - 208V/3Ph`, 24340682). |
| Circuits (source total) | 152 | 53 | -99 | Only circuits whose placed devices sit on a panel that exists in target. `place_circuits.py:32-43` builds `OK_PAN` from panels of placed devices ∩ target panels. In-scope = panel `BA` only (53). |
| Circuits (in-scope = panel BA) | 53 | 53 | 0 | `place_circuits_report.json`: `made=53, failed=0, updated=0`. |
| Keynotes | 23 | 23 | 0 | `place_keynotes_report.json` `placed=23, skipped=0`. |
| Wires (arrows) | 74 | 74 | 0 | `place_wires_report.json` `placed=74, skipped=0, removed=0`. |
| Textnotes | 16 | 16 | 0 | `place_textnotes_report.json` `placed=16, skipped=0`. |
| Fixture tags | 42 | 42 | 0 | `place_fixture_tags_report.json` `placed=42, skipped=0, removed=0`. |
| Wire tags | 60 | 60 | 0 | `place_wire_tags_report.json` `placed=60, skipped=0, removed=0`. |

`place_relative_map.json` device map = 75 entries; every source fixture id is present, no
fixture id missing. `wire_arrows` map has 74 entries (matches wires).

### Source Electrical Fixtures by family :: type (all 75 placed)

| Family :: Type | Count |
|---|---:|
| EF-U_Balanced Power Connector_CED :: 120V/1P | 6 |
| EF-U_Balanced Power Connector_CED :: 208V/3P | 2 |
| EF-U_Disconnect Switch_CED :: Fused - 100A | 1 |
| EF-U_Junction Box_CED :: Floor | 6 |
| EF-U_Junction Box_CED :: Wall - No Stem | 3 |
| EF-U_Motor Rated Switch_CED :: Motor Rated Switch - 120V, 1 Pole | 4 |
| EF-U_Receptacle_CED :: Cord Drop - 120V/1Ph | 2 |
| EF-U_Receptacle_CED :: Cord Drop - 120V/1Ph 2 IG | 5 |
| EF-U_Receptacle_CED :: Cord Drop - 208V/1Ph | 2 |
| EF-U_Receptacle_CED :: Cord Drop - 208V/3Ph | 4 |
| EF-U_Receptacle_CED :: Duplex Wall | 3 |
| EF-U_Receptacle_CED :: Duplex Wall - GFCI | 15 |
| EF-U_Receptacle_CED :: Duplex Wall - GFCI Horizontal | 12 |
| EF-U_Receptacle_CED :: Duplex Wall - Isolated Ground | 8 |
| EF-U_Receptacle_CED :: Specialty Wall - 208V/3Ph | 2 |
| **Total** | **75** |

Equipment-relative anchoring metrics (`place_relative_report.json`):
`unique_match=75`, `anchor_missing=0`, `no_anchor=0`, `ambiguous=0`,
`equip_shift_ft` min/avg/max = 0.0 / 0.59 / 2.85.

---

## 3. Gap-by-gap explanation

### Gap A — 3 panelboards not placed (EXPECTED-BY-DESIGN)

`place_relative.py:86-88`:
```python
for r in hosts:
    if r["category"] != "Electrical Fixtures":   # devices only for the test
        continue
```
Only `Electrical Fixtures` are ever appended to `rows`. The 3 `Electrical Equipment`
panelboards (`EE-U_Panelboard_CED`) are filtered out before matching, so they never reach
the placement loop. This is intentional — panelboards are target-side fixtures and are matched
by name in the circuiting step (`place_circuits.py:24-30` reads existing target panels).
**Not erroneous.**

### Gap B — 99 source circuits not created (EXPECTED-BY-DESIGN)

Source `circuits.json` has 152 circuits across panels: `RC`=59, `BA`=58, `FS`=32, `LDP4`=2,
`LDP3`=1. The circuiting step only builds groups for placed devices on in-scope panels:

`place_circuits.py:38-43`:
```python
for r in hosts:
    if str(r["id"]) in dev_map:
        for c in r.get("circuits", []):
            if c.get("panel"): src_pan.add(c["panel"])
OK_PAN = set(p for p in src_pan if p in tgt_panels)
```
Every **placed** bakery device's circuits reference panel **BA only** (74 device→circuit links,
all panel `BA`). Circuits on `RC`/`FS`/`LDP*` either belong to the excluded panelboards
(3 circuits are members-all-Electrical-Equipment) or to devices outside the bakery callout, so
they are never grouped. 53 distinct `BA` sys_ids → 53 groups → 53 circuits. **Not erroneous;**
it is the documented scope rule (devices placed AND source panel exists in target).

### Gap C — none for annotations

Keynotes/wires/textnotes/fixture_tags/wire_tags each map 1:1 with `skipped=0` and (where
applicable) `removed=0`. No gaps.

### Un-placed receptacles/devices

**There are none.** All 75 source Electrical Fixtures (including every receptacle, junction box,
balanced power connector, motor-rated switch, disconnect, cord drop) were placed. The report
shows `status:"matched"` for all 75 rows, `anchor_missing=0`, `no_anchor=0`. No row carries
`no-src-equip`, `anchor-id-not-in-target`, or `guid-mismatch` (the failure statuses defined at
`place_relative.py:93,105,108`). No `place_error` was recorded.

---

## 4. Circuiting analysis (`place_circuits_report.json`)

```
made=53  updated=0  failed=0  rows=[]  groups=53
target = CED_HEB Template_MEPR_R24_OkmntProfCorr   apply=true
```
- **Made: 53** new `ElectricalSystem` objects, one per source `sys_id` on panel `BA`.
- **Failed: 0** — `failed` list is empty; no `FAIL:` rows.
- **Panels in scope:** `BA` only. Absent (out of scope): `RC` (59 src ckts), `FS` (32),
  `LDP4` (2), `LDP3` (1). Reason: their devices/panelboards are not placed (§3 Gap B).
- All 53 in-scope circuits have all members placed (`circuits with ALL members placed = 53`).

Screenshot corroboration: circuit labels `BA-1` … `BA-84` appear throughout the view
(e.g. `BA-25,27,29`, `BA-44,46,48`, `BA-50`, `BA-73`), consistent with 53 panel-BA circuits
including multi-pole groupings.

---

## 5. Breaker-rating assessment (CRITICAL — verify in Revit)

Source `circuits.json` carries `rating`/`poles`/`frame` per circuit. Distribution for the
**53 in-scope (panel BA)** circuits:

| Rating | Count | | Poles | Count |
|---|---:|---|---|---:|
| 20 A | 43 | | 1-pole | 39 |
| 45 A | 3 | | 3-pole | 12 |
| 60 A | 3 | | 2-pole | 2 |
| 25 A | 2 | | | |
| 30 A | 1 | | | |
| 80 A | 1 | | | |

14 of 53 in-scope circuits are **non-default** (not 20 A / 1-pole). Examples:

| Panel/Ckt | Rating | Poles | Frame | Load |
|---|---|---|---|---|
| BA 50,52,54 | 45 A | 3P | 400 A | DOUGH DIVIDER - BAKERY |
| BA 68,70,72 | 45 A | 3P | 400 A | SPIRAL DOUGH MIXER - SHOCK BLOCK |
| BA 65,67,69 | 45 A | 3P | 400 A | FRENCH BREAD MOULDER - BAKERY |
| BA 26,28 | 60 A | 2P | 400 A | TORTILLA MACHINE 1 - BAKERY |
| BA 29,31 | 60 A | 2P | 400 A | TORTILLA MACHINE 2 - BAKERY |
| BA 56,58,60 | 80 A | 3P | 400 A | PAN WASHER - BAKERY PREP |
| BA 74,76,78 | 60 A | 3P | 400 A | DONUT FRYER |
| BA 62,64,66 | 30 A | 3P | 400 A | ALTO SHAAM OVEN |
| BA 71,73,75 | 25 A | 3P | 400 A | MULTI DECK CASE |
| BA 80,82,84 | 25 A | 3P | 400 A | MULTIDECK CASE |
| BA 59,61,63 / 44,46,48 / 47,49,51 / 53,55,57 | 20 A | 3P | 400 A | PROOFER / RACK OVENS / PLANETARY MIXER |

`place_circuits.set_rating_poles` (`place_circuits.py:81-100`) is **intended** to push these
source ratings/poles/frame onto each created circuit:
```python
pr = es.get_Parameter(DB.BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES); pr.Set(int(_amps(pv)))
... RBS_ELEC_CIRCUIT_RATING_PARAM ... RBS_ELEC_CIRCUIT_FRAME_PARAM ...
```
and is called on both create and update paths (`place_circuits.py:120, 133`).

**Honest status — cannot be confirmed from data files:**

- The data files (`circuits.json`, `place_relative_map.json`, `place_circuits_report.json`)
  record source values and create counts, **not the in-model breaker parameters** that were
  actually written. The report has no rating/poles fields.
- Timing: `skills/place_circuits.py` mtime = **16:31** (and it is an *untracked* file — not in
  git). `place_circuits_report.json` mtime = **16:13**. The circuiting run that produced the
  current target circuits therefore executed with code as it stood **before the 16:31 edit**.
- **Conclusion:** I cannot assert that the 14 non-default breakers in the live model carry the
  correct 25/30/45/60/80 A multi-pole values. Older runs are known to have left everything
  20 A / 1-pole. If the `set_rating_poles` fix had not yet been applied when the 16:13 run
  executed, the in-model breakers are likely still the 20 A / 1-pole default — wrong for these
  14 circuits and a code-violation risk for the 60–80 A equipment.

**Recommendation:** Re-run the circuiting step with the current (16:31) `place_circuits.py`,
then in Revit verify breaker rating/poles/frame for at least the 14 non-default `BA` circuits
above (spot-check PAN WASHER 80 A/3P, DONUT FRYER 60 A/3P, the two TORTILLA 60 A/2P).
Also commit `skills/place_circuits.py` so the fix is tracked.

---

## 6. Erroneous vs Expected-by-design classification

| Finding | Classification | Evidence |
|---|---|---|
| 3 panelboards not placed | **EXPECTED-BY-DESIGN** | `place_relative.py:87` category filter; panelboards handled by circuiting via target-side lookup. |
| 99 source circuits not created (RC/FS/LDP) | **EXPECTED-BY-DESIGN** | `place_circuits.py:38-43` scope rule; placed devices reference only panel BA. |
| 75/75 fixtures placed, 0 failures | **CORRECT** | `place_relative_report.json` all `matched`, no error statuses. |
| 53/53 in-scope circuits made, 0 failed | **CORRECT** | `place_circuits_report.json` `made=53 failed=0`. |
| Annotations 1:1 (kn/wire/text/fixtag/wiretag) | **CORRECT** | All `place_*_report.json` `skipped=0`. |
| 14 non-default breaker ratings — applied? | **UNVERIFIED RISK** (potentially ERRONEOUS) | Cannot confirm from data; report predates 16:31 skill edit. Requires re-run + in-Revit check. §5. |
| `skills/place_circuits.py` untracked in git | **PROCESS RISK** | `git status` shows `?? skills/place_circuits.py`; fix not committed/reproducible. |

No erroneous *placements* were found (no misplaced, duplicated, or wrong-type fixtures;
`dups`/`place_error` not triggered; `ambiguous=0`).

---

## 7. Recommendations

1. **Re-run circuiting** with current `place_circuits.py` (16:31) so `set_rating_poles` is
   applied by the running code, then **verify in Revit** the 14 non-default BA breakers (§5).
2. **Commit `skills/place_circuits.py`** (currently untracked) so the breaker-rating fix is
   version-controlled and the run is reproducible.
3. Confirm RC/FS/LDP exclusion is acceptable for this deliverable (those circuits belong to
   devices/panels outside the bakery callout — expected, but confirm scope with the engineer).
4. No remediation needed for fixtures or annotations — placement is complete and clean.

## 8. Screenshot

Captured `Power Callout - BAKERY - L1` from target `CED_HEB Template_MEPR_R24_OkmntProfCorr`
(view_id 8141848, Electrical Power Plan, 1/4" scale). Shows placed devices (red symbols),
wire arrows, BA-## fixture/circuit tags, numbered keynotes, and `(ON)`/`DOOR HEAT` text notes —
visually consistent with 75 fixtures + 53 panel-BA circuits + all annotations.
