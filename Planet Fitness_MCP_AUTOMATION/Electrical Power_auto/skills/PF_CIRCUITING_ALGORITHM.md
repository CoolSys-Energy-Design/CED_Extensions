# Planet Fitness Power Plan — Circuiting Algorithm

Reverse-engineered from Swannanoa NC reference circuiting. Use this as Pass 2 after PF_POWER_PLAN_PLAYBOOK.md places all the fixtures (Pass 1).

Apply this algorithm AFTER all electrical fixtures + keynotes are placed. Each yaml profile already sets `CKT_Panel_CEDT` on the fixture; this algorithm decides circuit numbers (single fixture per circuit vs. grouped, single-pole vs. multi-pole) and creates `ElectricalSystem` objects + `Wire` homeruns.

## Step 0 — Inventory

```python
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.DB.Electrical import ElectricalSystem, Wire

# Collect all electrical fixtures, group by their CKT_Panel_CEDT param
fixtures_by_panel = {}  # "L1": [(elem_id, load_name, x, y, va, voltage, poles), ...]
```

For each fixture, read:
- `CKT_Panel_CEDT` (which panel) — already set by yaml
- `CKT_Load Name_CEDT` (what kind of load)
- `Apparent Load Input_CED` (VA)
- `Voltage_CED` (120, 208, 240)
- `Number of Poles_CED` (1, 2, 3)
- Location.Point (X, Y for geographic clustering)

## Step 1 — Decide circuit grouping per panel

### Panel L1 — Cardio Equipment (1 fixture per circuit, dedicated)

Every fixture on L1 gets its own dedicated single-pole 20A circuit. **No grouping.**

| Pattern | Rule |
|---|---|
| TREADMILL / STAIRMASTER / POWERED BIKE | 1 fixture → 1 circuit, single-pole 20A |

In Swannanoa: 42 fixtures → 42 circuits (1..42).

### Panel L2 — Lighting, TV TRUSS, GEN receps, HVAC support

| Load pattern | Grouping | Breaker |
|---|---|---|
| `TV TRUSS` | **3 fixtures per circuit** (adjacent in the truss row, X-grouped) | 1P 20A |
| `GYM GEN. RCPTS` | **4–6 per circuit**, geographic cluster | 1P 20A |
| `ROOF MAINT. RCPTS` | 4 per circuit (HVAC pre-placed) | 1P 20A |
| `EF-1 & 2 RESTROOMS`, `EF-3 & 4 BCS & IT ROOM` | 2 per circuit (HVAC pre-placed) | 1P 20A |
| `CLOCK`, `PF SIGN` | Single fixture each | 1P 20A |
| Lighting (no load name) | Often very large groups (20–30 fixtures per zone) | 1P 20A |
| `TV - BLACK CARD SUITE` (FTMR variant) | Single fixture | 1P 20A |

**TV TRUSS specifically**: in Swannanoa 24 TV truss receps were grouped 3-per-circuit across ckts 1, 3, 5, 7, 9, 11, 13, 15 (only odd-numbered ckts — even numbers used for lighting). Each group covers 10 ft of X-span (e.g. ckt=1 = X=24..34, ckt=3 = X=39..49).

### Panel L3 — Front-of-house specialty + back-of-house utility

**Single-fixture dedicated circuits** (1 fixture → 1 circuit, 1P 20A):
- `VENDING MACHINE` (each vending)
- `IT SERVER RACK` / `IT RACK` (each)
- `HAND DRYER` (each)
- `HWH` (each)
- `CIRC. PUMP`
- `TMAX TIMER`
- `TV - WOMEN LOCKER ROOM`, `TV - MEN LOCKER ROOM` (43tv, each)
- `TV - BLACK CARD SUITE` (BCS 65tv)

**Grouped circuits** (multi-fixture, 1P 20A):
| Load | Group size |
|---|---|
| `GYM GEN. RCPTS` | 3–4 per circuit |
| `BCS GEN RCPTS` | 4–6 per circuit |
| `BREAKRM GEN RCPTS` | 5 per circuit |
| `RECEPTION GEN. RCPTS` | 3 per circuit |
| `CHECK-IN TABLETS` | 6 per circuit |
| `CHECK-IN COUNTER` | 3 per circuit |
| `MENS RESTRM SINKS`, `WOMENS RESTRM SINKS` | 3 per circuit (by gender) |
| `MENS RESTRM GEN. RCPTS`, `WOMENS RESTRM GEN. RCPTS` | 4 per circuit (by gender) |
| `MENS VANITY`, `WOMENS VANITY` | 2 per circuit (by gender) |
| `BATHROOM SINK` | 4 per circuit |
| `HYDROMASSAGE RECEPT - 103A` (the touchscreen GFCIs from HM63C profile) | 4 per circuit (all 4 hydromassage chairs grouped) |
| `BACKWRAP RCPTS` | 3 per circuit |
| `TVs & RADIANCE MONITOR - CHECK-IN 102` | 3 per circuit |
| `RECEPT - BCS` | 1 per circuit (or per room) |

**Multi-pole circuits on L3**:
- `BIG FAN` → 2-pole 30A, format "ckt=31,33" (odd-only — same phase? or 2 adjacent slots)
- `ECH-1 (VESTIBULE HEATER)` → 2-pole 20A, format "ckt=39,41"

### Panel L4 — 240V heavy spa equipment

Every L4 circuit is a multi-pole breaker for one piece of heavy equipment. **One fixture per circuit, no grouping.**

| Load | Breaker | Format |
|---|---|---|
| `STAND-UP TANNER - <room>` (BCS_Upright Tanning) | 3-pole 40A | "ckt=2,4,6" (3 adjacent even slots) |
| `LAY-DOWN TANNING - <room>` (PF Tanning 42-3) | 3-pole 40A | "ckt=8,10,12" |
| `HYBRID TANNER - <room>` (HM62E) | 3-pole 40A | "ckt=14,16,18" |
| `TLT TANNING - <room>` (BCS_Tanning Bed 42.4) | 3-pole 40A | "ckt=21,23,25" |
| `HYDROMASSAGE - 103A` (HM63C Specialty 240V/1Ph) | 2-pole 30A | "ckt=5,7" (2 adjacent odd slots) |
| `CRYOLOUNGE - 103` (BCS_Cryolounge) | 1-pole 20A | single (e.g. "ckt=1", "ckt=3") |
| `MASSAGE CHAIRS - 103` (SmarteCarte) | 1-pole 20A | single, **2 chairs grouped on 1 circuit** (e.g. "ckt=39") |
| `RED WAVE - 103E` (BCS_Tanning Red Wave) | 3-pole 40A | "ckt=27,29,31" |

**Multi-pole circuit format**: Revit shows multi-pole breakers as comma-separated slot numbers (e.g. `"2,4,6"` for a 3-pole on slots 2/4/6, or `"5,7"` for a 2-pole on 5/7). When you call `ElectricalSystem.CircuitNumber`, you write the comma-separated string yourself.

### Panel MDP-1 / MDP-2 — Feeders (don't touch)

These have high-amp circuits (200A–1200A) feeding sub-panels and existing house service. Pre-circuited by template. Skip.

## Step 2 — Voltage / phase encoding

Revit stores voltage in volts but in some internal display unit you can ignore. Cross-reference:

| Display VA | Actual nominal voltage |
|---|---|
| 1291.67 | 120 V (1Ph 2W) |
| 2238.89 | 208 V (3Ph) |
| 2583.34 | 240 V (1Ph 3W or 3Ph) |

When creating an `ElectricalSystem`, set voltage by linking to the panel's voltage spec — the system inherits from the panel.

## Step 3 — Geographic clustering for grouped circuits

For multi-fixture circuits (GYM GEN, RESTRM GEN, BCS GEN, etc.), group fixtures by spatial proximity:

1. Filter all fixtures in the room/zone with the same load name.
2. Sort by X (or Y, whichever has more spread).
3. Take the first N (group size from table above), assign to circuit K.
4. Take next N, assign to circuit K+2 (skip 1 for next phase if 3-phase panel, or K+1 for single-phase).
5. Continue until all fixtures circuited.

**Phase balancing**: in a 3-phase panel, alternate circuits between phases by skipping numbers. The pattern in Swannanoa: GYM GEN RCPTS used ckts 17, 19 (odd) — both on same phase; BCS GEN RCPTS used 11, 13; BREAKRM used 14; etc. The pattern isn't strict phase-balancing — just sequential within available odd or even slots.

## Step 4 — Circuit number assignment

Within each panel, number circuits sequentially starting from 1. Reserve slot ranges for known equipment types (e.g. on L4, reserve ckts 1, 3 for cryolounges, 5/7 for first hydromassage, 8/10/12 for first tanning bed, etc.). 42-circuit panels use slots 1..42.

For multi-pole breakers, use 2 or 3 adjacent slots:
- 2-pole: slots N, N+2 (same phase column in a 3-phase panel — even-only or odd-only)
- 3-pole: slots N, N+2, N+4 (one slot per phase)

## Step 5 — Create the ElectricalSystem + wires

```python
from Autodesk.Revit.DB.Electrical import ElectricalSystem
from Autodesk.Revit.DB import Transaction

t = Transaction(doc, "Create circuit")
t.Start()
# Create the system with the fixtures
es = ElectricalSystem.Create(doc, fixture_ids, system_type)
# Set the panel
es.SelectPanel(panel_element)
# Set circuit number (Revit assigns automatically, or override via CKT_Circuit Number_CEDT param)
t.Commit()
```

`system_type` is from `Autodesk.Revit.DB.Electrical.ElectricalSystemType` — typically `PowerCircuit`.

Then create `Wire` objects between fixtures and from first fixture to panel — these are the visible Arc wires on the floor plan. Revit can auto-route via `Wire.Create` with `WireType.WiringType.Arc` and specifying endpoints.

## Step 6 — Verify via panel schedules

After circuiting:
- Open the panel schedule view for L1/L2/L3/L4
- Each circuit should show its connected fixtures and computed load (VA)
- Total panel load should be reasonable (not overloaded)

## Quick-reference table — per-load circuit policy

| Load name | Panel | Group? | Breaker | Notes |
|---|---|---|---|---|
| TREADMILL / STAIRMASTER / POWERED BIKE | L1 | 1/ckt | 1P 20A | Dedicated |
| TV TRUSS | L2 | 3/ckt | 1P 20A | Odd ckts only (1,3,5...15) |
| GYM GEN. RCPTS | L2 or L3 | 4-6/ckt | 1P 20A | Geographic |
| BCS GEN RCPTS | L3 | 4-6/ckt | 1P 20A | Geographic |
| BREAKRM GEN RCPTS | L3 | 5/ckt | 1P 20A | Single circuit usually |
| RECEPTION GEN. RCPTS | L3 | 3/ckt | 1P 20A | Single circuit |
| CHECK-IN TABLETS / USB | L3 | 6/ckt | 1P 20A | Single circuit |
| CHECK-IN COUNTER | L3 | 3/ckt | 1P 20A | Single circuit |
| MENS / WOMENS RESTRM SINKS | L3 | 3/ckt by gender | 1P 20A | Separate ckt per gender |
| MENS / WOMENS RESTRM GEN. RCPTS | L3 | 4/ckt by gender | 1P 20A | Separate ckt per gender |
| MENS / WOMENS VANITY | L3 | 2/ckt by gender | 1P 20A | Separate ckt per gender |
| BATHROOM SINK | L3 | 4/ckt | 1P 20A | |
| HYDROMASSAGE RECEPT (GFCI touchscreens) | L3 | 4/ckt | 1P 20A | All 4 HM63C chairs grouped |
| HYDROMASSAGE - 103A (Specialty 240V) | L4 | 1/ckt | 2P 30A | "5,7" format |
| HAND DRYER | L3 | 1/ckt | 1P 20A | Dedicated |
| HWH | L3 | 1/ckt | 1P 20A | Dedicated |
| VENDING MACHINE | L3 | 1/ckt | 1P 20A | Each vending dedicated |
| IT SERVER RACK | L3 | 1/ckt | 1P 20A | Each rack dedicated |
| BIG FAN | L3 | 1/ckt | 2P 30A | "31,33" format |
| ECH-1 (VESTIBULE HEATER) | L3 | 1/ckt | 2P 20A | "39,41" format |
| TV - BLACK CARD SUITE / TV - locker | L2 or L3 | 1/ckt | 1P 20A | |
| STAND-UP TANNER (BCS_Upright Tanning HM62B) | L4 | 1/ckt | 3P 40A | "2,4,6" format |
| LAY-DOWN TANNING (PF Tanning 42-3) | L4 | 1/ckt | 3P 40A | |
| HYBRID TANNER (HM62E) | L4 | 1/ckt | 3P 40A | |
| TLT TANNING (BCS_Tanning Bed 42.4) | L4 | 1/ckt | 3P 40A | |
| RED WAVE - 103E (BCS_Tanning Red Wave) | L4 | 1/ckt | 3P 40A | |
| CRYOLOUNGE - 103 (BCS_Cryolounge) | L4 | 1/ckt | 1P 20A | |
| MASSAGE CHAIRS - 103 (SmarteCarte) | L4 | 2/ckt | 1P 20A | 2 chairs grouped |
| CLOCK, PF SIGN | L2 | 1/ckt | 1P 20A | |
| ROOF MAINT. RCPTS | L2 | 4/ckt | 1P 20A | HVAC pre-placed |
| EF-1 & 2 RESTROOMS, EF-3 & 4 BCS & IT ROOM | L2 | 2/ckt | 1P 20A | HVAC pre-placed |

## Pass-2 Workflow

1. **Load reference table** above into your script.
2. **Iterate all fixtures** by `CKT_Panel_CEDT`.
3. **For each panel**: bucket fixtures by load name, then apply the group-size rule.
4. **Sort each bucket by X then Y** (geographic order) before chunking into groups.
5. **Create ElectricalSystem** for each circuit (or each group within a load).
6. **Assign sequential circuit numbers** within the panel, using the multi-pole format for ≥2P breakers.
7. **Create Wire arcs** from panel to first fixture, then fixture-to-fixture for grouped circuits.
8. **Verify**: count circuits per panel matches expected (Swannanoa: L1=42, L2=42 max, L3=42, L4=42).

## What NOT to circuit

- **MDP-1, MDP-2, TR-L4, (E) PANELs** — these are feeders and pre-existing house service. Don't add circuits to these.
- **Pre-placed HVAC fixtures** (RTU disconnects, exhaust fan JBs) — usually already circuited by the template.
- **Lighting fixtures** (`OST_LightingFixtures`) — out of scope for power-plan circuiting; the lighting designer handles those.
- **Data Devices** (`OST_DataDevices`) — handled by the LV/data designer, not power.

## Open questions for future projects

- The exact phase-balancing convention isn't strictly enforced in Swannanoa (odd/even slot mixing isn't consistent). May need user confirmation per project.
- TV TRUSS groups-of-3 — is it always 3 even in projects with different TV truss lengths? Confirm.
- HYDROMASSAGE RECEPT groups-of-4 — assumes 4 hydromassage chairs. Confirm count per project.

---

# Pass-2 Add-on: Circuit Description (Load Name) + Load Classification + Breaker Sizing

Charlotte (Central) NC was used as the canonical reference (re-confirmed against Lindenhurst on 2026-05-21). Where Charlotte and the older Swannanoa data disagree, **prefer Charlotte**.

## CB sizes and pole counts — pull from yaml, not the table

**Do NOT hard-code 20A.** Each fixture's yaml profile sets `CKT_Rating_CED` (breaker amps) and `Number of Poles_CED`. Read those from the placed fixture and pass them through:

```python
rating = fixture.LookupParameter("CKT_Rating_CED").AsDouble()
poles  = fixture.LookupParameter("Number of Poles_CED").AsInteger()
# after es = ElectricalSystem.Create(...):
rp = es.LookupParameter("Rating")          # writable
if rp and not rp.IsReadOnly and abs(rp.AsDouble() - rating) > 0.01:
    rp.Set(rating)
# Pole count and Voltage come from the connector — don't try to set them on the ES.
```

L4 in Lindenhurst surfaced 4 different ratings (20/30/35/40A). Pulling from yaml avoids the all-20A trap.

**Yaml-vs-doc reconciliations seen in Lindenhurst** (yaml wins):

| Fixture | Algorithm doc said | Yaml actually says |
|---|---|---|
| `BIG FAN` | 2P 30A | **3P 20A** |
| `HYBRID TANNER` | 3P 40A | **2P 40A** |
| `STAND-UP TANNER` | 3P 40A | **2P 40A** |
| `TLT TANNING` (TANNING BED) | 3P 40A | **2P 40A** |
| `RED WAVE` | 3P 40A | **2P 35A** |

These may differ in other markets — always read the fixture, don't assume.

## Circuit description (ES `Load Name` parameter) — canonical Charlotte names

When you create an `ElectricalSystem`, its `Load Name` defaults to the fixture's connector Load Classification abbreviation (e.g. `"REC"`, `"NC"`, `"C"`). That is **not** the circuit description PF design conventions expect on the panel schedule. Override `es.LookupParameter("Load Name").Set("<canonical>")` per the table below.

| Fixture `CKT_Load Name_CEDT` | ES `Load Name` (Charlotte canonical) | Class (read-only, family-supplied) | Notes |
|---|---|---|---|
| TREADMILL | TREADMILL | NC | dedicated |
| STAIRMASTER | STAIRMASTER | NC | dedicated |
| POWERED BIKE | POWERED BIKE | NC | dedicated |
| TV TRUSS | TV TRUSS | C | 2–3/ckt |
| WATER FOUNTAIN | EWC | REC | dedicated |
| HWH | HWH-1 - WATER HEATER | REC | dedicated |
| GEN RCPTS - FT / MOBILITY / GYM | GYM MAINT. RCPTS | REC | bucket per Lindenhurst suffix; one ES per suffix bucket |
| GEN RCPTS - BCS | BCS MAINT. RCPTS | REC | 4–6/ckt |
| GEN RCPTS - BREAK | BREAK & MECH RM RCPTS | REC | 5/ckt |
| CHECK-IN TABLETS | CHECK-IN TABLETS | REC | 6/ckt |
| CHECK-IN COUNTER | CHECK-IN DESK RCPTS | REC | 3/ckt |
| TMAX RECEPT | TMAX TIMER | REC | grouped if multiple TMAX cells (Lindenhurst convention) |
| MENS RESTRM SINKS | MENS SINK RCPTS | REC | 3/ckt; split by X (mens X < threshold) |
| BATHROOM SINK | WOMENS SINK RCPTS | REC | Lindenhurst names the women's Sloan trough "BATHROOM SINK"; verify by X cluster |
| HAND DRYER (mens) | MENS HAND DRYER | NC | dedicated; X < gender threshold |
| HAND DRYER (womens) | WOMENS HAND DRYER | NC | dedicated; X ≥ gender threshold |
| TV - MEN/WOMEN LOCKER ROOM (mens) | MENS LOCKRM TV | C | dedicated |
| TV - MEN/WOMEN LOCKER ROOM (womens) | WOMENS LOCKRM TV | C | dedicated |
| TV - BLACK CARD SUITE | TV - BLACK CARD SUITE - 103 | C | dedicated (in BCS). FTMR variant uses `TV - FUNCTIONAL TRAINING`. |
| TV - DIGITAL MEDIA | TV & RADIANCE MONITOR - CHECK-IN 102 | C | grouped with the reception TV+radiance recep |
| TV & RADIANCE MONITOR - CHECK-IN 102 | TV & RADIANCE MONITOR - CHECK-IN 102 | C | 3/ckt with TV - DIGITAL MEDIA |
| HYDROMASSAGE RECEPT - 103A | HYDROMASSAGE RCPTS - 103A | C | 4/ckt (or 3 if site has 3 chairs) |
| IT SERVER RACK | I.T. RACK | REC | dedicated, one ES per rack |
| VENDING MACHINE | VENDING MACHINE | NC | dedicated |
| BIG FAN | BIG FAN | C | dedicated, **3P** per yaml |
| SIGNAGE (IF NEEDED) | STORE SIGN | NC | dedicated |
| RECEPT - BCS | BCS MAINT. RCPTS | REC | 6–8/ckt — preferred over Swannanoa "RECEPT - BCS" name |
| BACKWRAP RCPTS | BACKWRAP RCPTS | REC | 4/ckt (skip if no backwrap on site) |
| HYDROMASSAGE - 103A | HYDROMASSAGE - 103A | NC | dedicated, **2P 30A** |
| CRYOLOUNGE - 103 | CRYOLOUNGE | C | dedicated, 1P 20A |
| MASSAGE CHAIRS - 103 | MASSAGE CHAIRS - 103 | NC | 2 chairs grouped per circuit |
| HYBRID TANNER - <room#> | HYBRID TANNER - <room#> | NC | dedicated, **2P 40A**; preserve unique room # per tanner |
| STAND-UP TANNER - <room#> | STAND-UP TANNER - <room#> | NC | dedicated, **2P 40A** |
| TLT TANNING - <room#> | **TANNING BED - <room#>** | NC | dedicated, **2P 40A** — Charlotte renames "TLT TANNING" → "TANNING BED" |
| RED WAVE - <room#> | RED WAVE - <room#> | NC | dedicated, **2P 35A** |
| SAUNA - <room#> | SAUNA - <room#> | NC | dedicated, 2P 40A (skip if no sauna) |
| ECH-1 (vestibule heater) | ECH-1 | NC | dedicated, 2P 20A |
| P-1 / CIRC PUMP | P-1 - CIRC. PUMP | NC | dedicated |
| FACU junction box | F.A.C.U | C | dedicated 1P 20A; placed near L1/L2 cluster with keynote #6 |

### Gender split heuristic

For Lindenhurst-style projects, partition gendered fixtures (HAND DRYER, TV - MEN/WOMEN LOCKER ROOM, BATHROOM SINK, MENS RESTRM SINKS) by an X-coordinate threshold. Pick the threshold by sampling positions of any one gendered cluster — the midpoint between the two cluster centroids works. Mens X < threshold, womens X ≥ threshold. In Lindenhurst the threshold landed at ~575 ft.

## Load Classification — read-only on the ES

`ElectricalSystem.LookupParameter("Load Classification")` is **read-only** — it aggregates from the connector(s) of the connected fixture(s). Setting `"Load Classification"` directly on the ES throws `The parameter is read-only.`

Lindenhurst showed `TREADMILL` connector class = `REC`; Charlotte's TREADMILL connector class = `NC`. They use **different family connector definitions**, even though the canonical Charlotte circuit description is "TREADMILL". To match Charlotte's classifications, the fix is at the **family connector level**, not on the ES:

1. Open the family RFA.
2. Edit the electrical connector → Load Classification.
3. Reload family into project.

Alternative if family edits are out of scope: live with the Lindenhurst-family classifications and only override the `Load Name` (description). Note this in the QAQC for the project so the user can choose whether to swap the family.

## What I CAN write on the ES (from this work)

| Param | Writable? | Source |
|---|---|---|
| `Load Name` | ✓ | canonical Charlotte map above |
| `Rating` (breaker amps) | ✓ | fixture's `CKT_Rating_CED` |
| `Schedule Circuit Notes` | ✓ | optional notes field |
| `Load Classification` | ✗ read-only | family connector |
| `Voltage` | ✗ read-only | family connector |
| `Number of Poles` | ✗ read-only | family connector |
| `CircuitNumber` | ✗ read-only | Revit auto-assigns from panel's next slot when `SelectPanel()` is called |

## Pass-2 ordering (revised)

1. **Inventory** by `CKT_Panel_CEDT`.
2. **Group** per the table above (bucket by suffix where applicable; geographic-cluster the bigger groups).
3. **For each ES**: `ElectricalSystem.Create(doc, ids, PowerCircuit)` → `SelectPanel(panel)` → set `Rating` from fixture's yaml → set `Load Name` per canonical map.
4. **Defer** Wire homerun arcs (Pass 2b) until the user confirms circuit groupings + descriptions + ratings are correct.
5. **Defer** circuit number tagging (Pass 3) — Revit auto-assigns slot numbers; tags read those.

## Verification snippet

```python
for pname, pid in [("L1",...),("L2",...),("L3",...),("L4",...)]:
    panel = doc.GetElement(DB.ElementId(pid))
    systems = list(panel.MEPModel.GetAssignedElectricalSystems())
    for es in systems:
        ln = es.LookupParameter("Load Name").AsString()
        if ln in ("SPARE","SPACE"): continue
        print(pname, es.CircuitNumber, ln, es.LookupParameter("Rating").AsDouble(), "P=", es.LookupParameter("Number of Poles").AsInteger())
```

---

# Pass 2c: Home runs + tagging

Synthesized from Charlotte (Central) NC E101 Power Plan, applied to Lindenhurst on 2026-05-21.

## Key insight: don't draw a homerun for every circuit

Charlotte's E101 has **~225 circuited fixtures but only 7 wires**. The convention is:
- **Every fixture** gets a `Panel & Circuit Number` tag (text reads `L1/5`, `L3/27,29,31`, etc.) — that is how the circuit assignment is documented for the bulk of fixtures.
- **A small set of representative wires** are drawn to indicate homerun routing TYP for whole equipment rows or for multi-pole specialty loads.

## When to draw a wire homerun (apply these rules)

Draw a wire (`Autodesk.Revit.DB.Electrical.Wire`) only for:

1. **One sample per cardio equipment ROW × TYPE.** Cluster cardio fixtures by Y coordinate; for each row, draw one wire per equipment type present (TREADMILL / STAIRMASTER / POWERED BIKE). Charlotte = 2 treadmill rows × 1 wire each + 1 stairmaster + 1 bike = 4 wires. Lindenhurst = 1 row × 3 types = 3 wires.
2. **Each multi-pole specialty load.** Charlotte drew a wire for each BIG FAN (3-pole) so the wire tag could display `L3/27,29,31`. Same for tanners with multi-pole breakers if you choose to wire them.
3. **One representative per grouped circuit type** that benefits from visualizing the homerun (e.g., 1 wire on one TV TRUSS group). Optional — fixture tags alone are enough for most groups.

Everything else (L3 sinks, restroom receps, locker rooms, IT racks, vending, hand dryers, etc.) is documented **with fixture tags only — no wire**.

## Tag types in this project (Lindenhurst, but same family library)

| Purpose | Family | Type | Lindenhurst ElementId |
|---|---|---|---|
| Fixture Panel/Ckt | `EF-Tag_Electrical Fixtures_CED` | `Panel & Circuit Number` | 2789668 |
| Equipment label | `EE-Tag_Electrical Equipment_CED` | `Panel Name` | 1784154 |
| Wire homerun | `E_Tag - Wire` | `Homerun - Panel & Circuits - Slash` | 4959576 |
| Special comments (e.g. "60A/3 NF", "F.A.C.U") | `EF-Tag_Electrical Fixtures_CED` | `Comments` | 2801981 |

`Panel & Circuit Number` tag reads from the circuited fixture's ES and produces strings like `L1/21`, `L3/81,83,85`. The wire tag reads from the bound `Wire.MEPSystem` and produces the same format.

## Wire creation API (the working incantation)

```python
from Autodesk.Revit.DB.Electrical import Wire, WiringType
from System.Collections.Generic import List

# Signature: Wire.Create(doc, wireTypeId, viewId, wiringType, IList<XYZ>, startConnector, endConnector)
# NOTE: third arg is VIEW Id, not LEVEL Id (older docs say level — they are wrong).

def get_electrical_connector(fixture):
    for c in fixture.MEPModel.ConnectorManager.Connectors:
        if c.Domain.ToString() == "DomainElectrical":
            return c

def make_charlotte_style_homerun(fixture, jog_ft=4.0, arc_ft=14.0):
    """Three-vertex Charlotte homerun: V0=fixture conn, V1=V0+(0,+jog),
    V2=V1+(-arc, 0). Jog NORTH (perpendicular to equipment row), arc WEST."""
    fc = get_electrical_connector(fixture)
    o = fc.Origin
    v0 = DB.XYZ(o.X, o.Y, o.Z)
    v1 = DB.XYZ(o.X, o.Y + jog_ft, 0)
    v2 = DB.XYZ(v1.X - arc_ft, v1.Y, 0)
    pts = List[DB.XYZ]()
    pts.Add(v0); pts.Add(v1); pts.Add(v2)
    # endConnector=None → wire becomes a homerun stub.
    return Wire.Create(doc, wire_type_id, view.Id, WiringType.Arc, pts, fc, None)
```

The wire **automatically binds to the fixture's existing `ElectricalSystem`** (so `wire.MEPSystem` is set and the wire tag picks up the panel/ckt text). No explicit `ConnectTo` is needed because the start-connector is already part of a system.

### Critical: three vertices, not two

Charlotte's wires use **3 vertices** to draw an L-shape:
1. **V0**: at the fixture's electrical connector
2. **V1**: short jog (~4 ft for cardio, ~1.5 ft for BIG FAN, ~3 ft for TV TRUSS) perpendicular to the equipment row — typically NORTH (toward the ceiling/joist direction in plan)
3. **V2**: longer arc (~14 ft for cardio, ~2 ft for BIG FAN, ~7 ft for TV TRUSS) along the row direction — Charlotte arcs **WEST** by convention

A 2-vertex wire (straight line from fixture toward panel) **does NOT match** Charlotte's TYP style. The L-shape is what makes it read as a homerun in plan view.

### Pick the right sample fixture

For cardio rows, pick the **easternmost fixture** of each equipment type so the westward arc visually traces through the row toward the panel side. For BIG FAN, pick each fan (one wire per multi-pole circuit). For TV TRUSS, pick the easternmost fixture of the easternmost truss group.

### Jog/arc lengths (Charlotte sample, copy these)

| Load | Jog (ft, +Y) | Arc (ft, -X) |
|---|---|---|
| TREADMILL | 4 | 14 |
| STAIRMASTER | 4 | 14 |
| POWERED BIKE | 4 | 4 (short — only 2 bikes) |
| BIG FAN | 1.5 | 2 |
| TV TRUSS | 3 | 7 |

Direction signs: jog `+Y` (north) and arc `-X` (west) are conventions that match Charlotte's drawings. If the project orientation is rotated, rotate the directions accordingly — but always keep the L-shape (perpendicular jog, then parallel arc).

## Wire tag placement

Place wire tags near the **unconnected (panel-side) end** of the wire — that's where the homerun arrow points. Get the open endpoint via:

```python
for c in wire.ConnectorManager.Connectors:
    if not c.IsConnected:
        tag_anchor = c.Origin
        break
tag = IndependentTag.Create(doc, WIRE_TAG_TYPE, view.Id, Reference(wire), False, TagOrientation.Horizontal, XYZ(tag_anchor.X, tag_anchor.Y + 1.0, 0))
```

The tag text auto-fills from `wire.MEPSystem` (e.g., `L1/21`, `L3/81,83,85`).

## Fixture tag placement

Default to **1.5 ft north of the fixture** (Y+) with horizontal orientation, no leader. Adjust per fixture if the tag overlaps a keynote — the user will visually fix anything that collides; don't over-engineer the algorithm.

```python
loc = fixture.Location.Point
tag_pos = DB.XYZ(loc.X, loc.Y + 1.5, loc.Z)
IndependentTag.Create(doc, EF_TAG_TYPE, view.Id, Reference(fixture), False, TagOrientation.Horizontal, tag_pos)
```

## Comments tags (deferred; tag manually or in Pass 2d)

Charlotte applies the `Comments` tag type (14 instances in E101) to:
- Specialty equipment to show **disconnect/fuse rating** (e.g., BIG FAN → `30A/3 NF`; tanners → `60A/3 NF`)
- Special labels: `F.A.C.U`, `ECH-1`, plug type `L6-30R`

The text comes from the fixture's `Comments` parameter (NOT the canonical circuit description). Setting `Comments` on the fixture instance then placing this tag type renders the value. Not required for first-pass circuiting; flag as Pass 2d if the user asks.

## Charlotte vs Lindenhurst wire count summary

| Project | Fixtures | Circuits | Wires drawn | Wire tags | Fixture tags |
|---|---|---|---|---|---|
| Charlotte E101 | ~225 | ~225 | 7 | 7 | 225 |
| Lindenhurst | 136 | 85 | 6 | 6 | 136 |

The 6 Lindenhurst wires: 1 TREADMILL + 1 STAIRMASTER + 1 POWERED BIKE + 2 BIG FAN + 1 TV TRUSS.
