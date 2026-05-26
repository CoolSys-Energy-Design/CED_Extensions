# HEB Pharmacy Replication — Audit Report

**View:** Power Callouts - PHARMACY
**Source doc:** RunUpdateProfilesHEBCarrollton95PercentCLOUDMODEL (collected snapshot: source = "Equip" / RunUpdate model)
**Target doc:** CED_HEB Template_MEPR_R24_OkmntProfCorr (against ARCH link 00000_Oakmont_v24_HEB_ARCH)
**Date:** 2026-05-19

---

## 1. Inventory — Side-by-Side

| Item | Source count | Placed count | Delta | Classification |
|---|---|---|---|---|
| **Devices — Electrical Fixtures** | 71 | 71 | 0 | OK (100%) |
| **Devices — Electrical Equipment** | 2 | 0 | -2 | EXPECTED (excluded by design) |
| Devices total (host_elements) | 73 | 71 | -2 | EXPECTED |
| Keynotes | 16 | 16 | 0 | OK |
| Textnotes | 12 | 12 | 0 | OK |
| Wires | 57 | 57 | 0 | OK |
| Wire arrows | (derived from wires) | 57 | 0 | OK |
| Fixture tags | 18 | 18 | 0 | OK |
| Wire tags | 31 | 31 | 0 | OK |
| **Circuits** | 89 source rows / 33 in-scope | 32 made + 1 FAIL | -1 | 1 ERRONEOUS-ish (see §4) |

### Source device breakdown (host_elements.json — 73 total)

| Category | Family :: Type | Count |
|---|---|---|
| Electrical Equipment | CED-E-MISC EQUIPMENT :: LIGHTING CONTROL PANEL ( LARGE LCP) | 1 |
| Electrical Equipment | EE-U_Panelboard_CED :: zPanelboard - 208V/3Ph, IG Bus | 1 |
| Electrical Fixtures | CED-E-STUB-UP :: Stub Up | 4 |
| Electrical Fixtures | EF-U_Junction Box_CED :: Wall - No Stem | 3 |
| Electrical Fixtures | EF-U_Motor Rated Switch_CED :: Motor Rated Switch - 208V, 2 Pole | 1 |
| Electrical Fixtures | EF-U_Receptacle_CED :: Duplex Wall | 14 |
| Electrical Fixtures | EF-U_Receptacle_CED :: Duplex Wall - GFCI | 2 |
| Electrical Fixtures | EF-U_Receptacle_CED :: Duplex Wall - Isolated Ground | 3 |
| Electrical Fixtures | EF-U_Receptacle_CED :: Duplex Wall - Tamper Resistant | 1 |
| Electrical Fixtures | EF-U_Receptacle_CED :: Quad Wall | 2 |
| Electrical Fixtures | EF-U_Receptacle_CED :: Quad Wall - Isolated Ground | 41 |

All 63 receptacles + 4 stub-ups + 3 junction boxes + 1 motor-rated switch = 71 fixtures, **all placed**.

---

## 2. Why the 2 "missing" devices were NOT placed (EXPECTED BY DESIGN)

Not placed:
- id **7855682** — Electrical Equipment, EE-U_Panelboard_CED :: zPanelboard - 208V/3Ph, IG Bus
- id **9688811** — Electrical Equipment, CED-E-MISC EQUIPMENT :: LIGHTING CONTROL PANEL ( LARGE LCP)

Root cause — `skills/place_relative.py` only ever considers Electrical **Fixtures**; the placement loop skips every other category:

```python
# place_relative.py line 86-88
for r in hosts:
    if r["category"] != "Electrical Fixtures":   # devices only for the test
        continue
```

`devices_total` in the report is likewise defined as fixtures only (`place_relative.py:128`). Panelboards and the lighting control panel are Electrical Equipment and are intentionally excluded — the pipeline replicates field devices, not the upstream distribution gear (which is expected to already exist in the target template). **No bug. Expected.**

Confirmation from the placement report (`place_relative_report.json`): `devices_total = 71`, `unique_match = 71`, `anchor_missing = 0`, `no_anchor = 0`, `ambiguous = 0`. All 71 fixture rows have `status = "matched"` — there were zero "anchor missing" / "no-src-equip" / "guid-mismatch" skips. Equipment-relative shift: min 0.93 ft, avg 1.96 ft, max 6.13 ft.

---

## 3. Receptacles: all placed (no gap)

Every receptacle in the source callout (63 of them across Duplex/Quad/GFCI/IG/TR types) was placed. `place_relative_map.json["devices"]` has exactly 71 entries, one per source Electrical Fixture id; the set of map keys exactly equals the set of source fixture ids (verified — 0 missing, 0 extra). There is **no receptacle placement gap**.

---

## 4. Circuiting Failure Analysis

`place_circuits_report.json`: `groups = 33`, `made = 32`, `updated = 0`, `failed = 1`.

The 89 rows in `circuits.json` decompose as:
- **33** circuits where every device member is a placed fixture → in scope (32 made + 1 failed).
- **30** "no-members" circuits (panel/feeder-level systems with no field devices) → not creatable, nothing to group.
- **25** circuits whose members are not in the PHARMACY callout (devices live in other areas) → out of scope.
- **1** equipment-only circuit (members are excluded Electrical Equipment) → out of scope.

This decomposition matches `place_circuits.py` exactly: it groups only devices that are in `dev_map`, on a panel that exists in the target (`OK_PAN = src panels ∩ target panels`), then creates one `ElectricalSystem` per source `sys_id` (`place_circuits.py:47-56, 100-110`). Placed circuits by panel: PH=27, CS=3, CL=2 (= 32).

### The single FAIL

```
{ "panel": "PH", "src_sys": "10612412", "status": "FAIL:The panel and circuit do not match." }
```

Source circuit **10612412 = "EWH-1 - PHARMACY"** — panel PH, ckt "21,23", **2 poles**, **208 V**, 20 A.
Members (both placed successfully as fixtures):
- 10612405 — EF-U_Junction Box_CED :: Wall - No Stem
- 10612409 — EF-U_Motor Rated Switch_CED :: Motor Rated Switch - 208V, 2 Pole

Its source `base_equipment` is id 7855682 = the `zPanelboard - 208V/3Ph, IG Bus` — i.e. the very panelboard that is excluded as Electrical Equipment (§2).

**Why it fails while 32 others succeed:** every one of the 32 successful circuits is **1-pole / single-phase / 120 V** (verified: poles distribution of placed circuits = {1: 32}). Circuit 10612412 is the **only 2-pole 208 V circuit** in scope. The exception text `"The panel and circuit do not match."` is raised by the Revit API call `es.SelectPanel(tp[g["panel"]])` (`place_circuits.py:103`): the target panel named "PH" cannot accept a 2-pole 208 V multi-wire branch circuit (panel phase/voltage configuration in the target template does not match the circuit's voltage/poles). The code catches it and records the FAIL row (`place_circuits.py:108-110`); it does not crash the pipeline.

---

## 5. Erroneous vs Expected — Classification

| Finding | Verdict | Evidence |
|---|---|---|
| 2 Electrical Equipment (panelboard + LCP) not placed | **EXPECTED by design** | `place_relative.py:86-88` filters to "Electrical Fixtures" only |
| All 71 fixtures placed, all status "matched" | **OK** | `place_relative_report.json` (unique_match=71, anchor_missing=0) |
| 30 "no-member" + 25 out-of-area + 1 equipment-only circuits not created | **EXPECTED by design** | `place_circuits.py:47-56` scope rules; members absent from callout |
| Circuit 10612412 "EWH-1 - PHARMACY" FAIL ("panel and circuit do not match") | **DATA/CONFIG issue, not a code bug** | Only 2-pole/208 V circuit in scope; target panel "PH" config rejects it (`place_circuits.py:103`). The devices ARE placed; only the electrical connection is missing. Worth fixing manually. |
| Keynotes/textnotes/wires/wire arrows/fixture tags/wire tags | **OK — 100%** | Each report `skipped=0`, placed = source count |

There are **no genuine software bugs** in this run. The only real action item is the one un-circuited load (EWH-1), which is a target-panel phase/voltage mismatch — its physical devices were placed correctly, only the `ElectricalSystem` could not be assigned to panel PH.

---

## 6. Recommendations

1. **EWH-1 / circuit 10612412:** Manually circuit the placed Junction Box (tgt id of src 10612405) + Motor Rated Switch (src 10612409) as a 2-pole 208 V circuit on the appropriate target panel. Verify the target "PH" panel's phase/voltage; the source panelboard was 208 V/3-phase but the target panel as named may be single-phase 120 V. If a different panel hosts 208 V 2-pole loads in the target, assign there.
2. **Equipment (panelboard + LCP):** Confirm these already exist in the target template at the correct location; they are intentionally not replicated by the pipeline. If they do not exist, place them manually (out of pipeline scope).
3. **Pipeline enhancement (optional):** `place_circuits.py` could pre-check circuit poles/voltage against the target panel's electrical config and emit a clearer, classified report row (e.g. "skipped: 2-pole 208V load, panel PH is 120V 1-phase") instead of relying on the generic Revit exception string.
4. No action needed for the 55 out-of-scope source circuits (no members / members in other areas / equipment-only) — expected behavior.

---

## 7. Phase 2 — Screenshot status

**SKIPPED.** The long-running Revit pipeline did not signal completion within the
coordination budget (~15 poll iterations of the task output file
`...\tasks\bb8fz3k7l.output`, which never showed "PIPELINE COMPLETE" or "TIMEOUT").
Per the read-first coordination constraint, no `mcp__revit__*` tool was called
(doing so while the pipeline holds an open Revit transaction would corrupt it).
No Revit screenshot of "Power Callouts - PHARMACY" was captured. All Phase-1
findings above are derived purely offline from the JSON data + skill source and
are complete and cross-verified.
