# Planet Fitness E101 — Training synthesis

Source view: **E101 - Power Plan** (view id 2243920), project "Planet fitness North Asheville, NC". All raw JSON dumps in this folder.

## Files in this folder

| File | Contents |
|---|---|
| `hosts.json` | 191 Electrical Fixtures + 9 Electrical Equipment, full instance + type params, location, host, level, facing/hand orientation |
| `keynotes.json` | 49 Generic Annotation keynote symbols, with `CED-G-NOTE #` value |
| `textnotes.json` | empty — no TextNote elements in E101 |
| `wires.json` | 6 THWN arc homerun wires, 3 vertices each, with MEPSystem link + Panel/Circuits params |
| `fixture_tags.json` | 212 fixture tags (`IndependentTag`), 119 distinct strings, with `tagged_elements` link |
| `wire_tags.json` | 6 wire tags, all `E_Tag - Wire :: Homerun - Panel & Circuits - Slash` |
| `circuits.json` | 140 ElectricalSystem circuits, with `panel.id` link, voltage, poles, rating, load name, elements-on-circuit |

## Distribution topology

```
(E) UTILITY XFMR 750 kVA  (exterior NW, near grid 3/A)
        │ (no DS shown — utility primary)
        ▼
DP  Switchboard  480Y/277V  (top, electrical room)
 ├── TR-L1  150 kVA  480→208Y/120V ─► L1  600A 208Y/120V  ─► L2 (sub-feed cct 55,57,59) ─► (L3?)
 ├── TR-L4  112.5 kVA  480→240Δ    ─► L4  240V Δ
 └── Fused 200A 480V equipment switch
```

Voltage sanity: L1–L3 circuits report 208 V / 120 V on `ABC` phase — 208Y/120V. L4 circuits report 240 V on `ABC` — 240V Δ. Matches the rule "L1–L3 always 208V Y; L4 always 240V Δ."

## Panel → load zoning (the PF rule)

| Panel | Voltage | Building zone (E101) | Loads |
|---|---|---|---|
| **L1** | 208Y/120V, 600A | top (electrical room) | All cardio — treadmills + ellipticals. Each machine = its own 1-pole 120V branch. 46 circuits in view. |
| **L2** | 208Y/120V, 125A | top (electrical room) | Strength side (E elev), ECH-1 vestibule heater (L2/42), locker/restroom loads, some misc. 48 circuits. |
| **L3** | 208Y/120V, 125A | **bottom (IT room in template, per user)** | Tanning rooms, reception/check-in, spa, 120V back-of-house. 21 circuits. |
| **L4** | 240V Δ | **bottom (IT room in template, per user)** | 240V dedicated: hybrid tanning, red wave, hydromassage. Each load = dedicated NEMA 6-30 on 60A/3 NF disconnect. 14 circuits. |
| **DP** | 480Y/277V | top (electrical room) | Service entrance switchboard, 8 feeder circuits (to transformers + L4). |

**Replication rule:** when moving panels in the new template, keep this zoning:
- DP, TR-L1, L1, L2 → electrical/storage room near utility XFMR
- L3, L4, TR-L4 → IT room (per user's explicit instruction)

## Wire / homerun pattern

Only 6 wires in E101 — one homerun arc per equipment row. Style: 3-vertex arc, `Arc` wiring type, THWN. Wire tag sits on the kink and shows `Panel/circ,circ,circ,...` (slash + comma list).

| Wire id | Tag text | What it serves |
|---|---|---|
| 7815601 | `L1/25,27,29,31,33,35,37,39,41,43` | N cardio row, odd legs (treadmills) |
| 7815828 | `L1/26,28,30,32,34,36,38,40,42,44` | N cardio row, even legs |
| 7815991 | `L1/1,3,5,7,9,11,13,15,17,19,21,23` | S cardio row, odd legs |
| 7816149 | `L1/2,4,6,8,10,12,14,16,18,20,22,24` | S cardio row, even legs |
| 7837403 | `L2/1,3,5,7,9,11,13,15` | Strength row |
| 7940318 | `L2/42` | ECH-1 vestibule heater |

## Fixture-tag patterns

3 distinct tag-type purposes layered on the same fixture:

1. **Panel/circuit** — `L3/25`, `L1/3` — most common (~118 tags). Auto-populated from the circuit assignment. Tag type: `E_Tag_Electrical Fixtures_CED :: Panel & Circuit`.
2. **Elevation (inches)** — `30"`, `39"`, `46"` — for cardio recep heights (different machines need different recep AFF). Tag type: `EF-Tag_Electrical Fixtures_CED :: Elevation (Inches)`.
3. **Letter / disconnect callout** — `U`, `TV`, `60A/3\nNF` — equipment-type identifier or disconnect spec. Multi-line.

Same physical fixture can carry **all three** tag types simultaneously (orphan count: 0; every tag has a tagged element).

## Keynote pattern

49 `Manual Key Note- All Shapes` symbols in two type variants (`Square`, `Square - Typical`). The keynote number lives in instance param **`CED-G-NOTE #`** (string `"1"` .. `"20"`). Companion text param `CED-G-NOTE TEXT` is unfilled in this project — the legend image is the source of truth.

Keynote → equipment context mapping (inferred from positions in E101 + your legend image):

| # | Legend meaning | Where it appears in E101 |
|---|---|---|
| 1 | GFCI faceplate recep above counter | Check-in / Reception counter |
| 2 | Hand dryer JB @ 42" | Restrooms, near Janitor / WCs |
| 3 | GFCI duplex below countertop | Reception under-counter |
| 4 | Hydromassage unit, dedicated 208V NEMA L6-30 + 12V touchscreen | Hydromassage room (the "4" cluster) |
| 5 | Tanning/hybrid bed 240V/30A NEMA 6-30 + 60A/3 NF disconnect | Hybrid Tanning, Red Wave, Tanning rooms |
| 6 | Relocated FACU | Electrical / Storage room |
| 7 | Isolated ground for check-in IT recep | Check-in (paired with kn 8) |
| 8 | Phone / RJ45 jack for FA/MGR; data wire to tanning | Reception, near data |
| 9 | Data rack / AV vendor coord | Reception area |
| 10 | Coord backwrap + check-in recep with tenant | Black Card / Check-in |
| 11 | TV mounting height/connection | TV locations |
| 12 | Mount recep behind drinking fountain | (none seen in this view — TBD) |
| 13 | Power & data raceway + recessed JB for floor routing | Cardio area (raceway base) |
| 14 | Coord club comm with low-voltage vendor | Reception |
| 15 | Mount recep/USB ports horizontally above counter | Reception |
| 16 | Beverage cooler — dedicated recep per mfr | Reception kiosk area |
| 17 | Power for future signage | Front facade |
| 18 | Planet Fitness clock, mount per tenant | Cardio / common areas |
| 19 | Coord with fan mfr | Cardio (ceiling fans?) |
| 20 | False column raceway — drop wire down false column for cardio raceway | Cardio rows (every row has this) |
| 21 | Power for ECH-1, see L1 schedule | Vestibule (L2/42 wire; legend says L1 but circuit lives on L2 in this project) |

(Discrepancy noted: keynote 21 legend says "L1 schedule" but ECH-1 is on L2-42 in this model. Worth flagging during replication.)

## Family / type catalog observed (CED library)

**Equipment families:** `EE-U_Panelboard_CED`, `EE-U_Transformer_CED`, `EE-U_Switchboard_CED`, `EE-U_Utility Transformer_CED`, `EE-U_Equipment Switch_CED`

**Fixture families (top instances):** `EF-U_Receptacle_CED :: Duplex Wall`, `EF-U_Junction Box_CED :: Wall - With Stem` (and likely 240V dedicated recep / TV / data variants — full list in `hosts.json`)

**Tag families:** `E_Tag - Wire :: Homerun - Panel & Circuits - Slash`, plus `EF-Tag_Electrical Fixtures_CED` family with multiple tag types

**Keynote family:** `Manual Key Note- All Shapes :: Square` (and `Square - Typical`)

## Replication checklist (for the new project — Pass 1 = recepts + fixture tags only)

1. Switch active view via MCP to the target plan view in the new project.
2. Inventory the linked DWG: enumerate blocks (name, insertion point, rotation). Identify equipment by block name pattern; ask the user when ambiguous.
3. **Move (don't place new)** the existing panels per PF zoning:
   - DP, TR-L1, L1, L2 → electrical/storage room near utility XFMR
   - L3, L4, TR-L4 → IT room
4. For each piece of equipment in the new CAD background, place the matching CED fixture family from the catalog at the correct relative position (offset behind/under/above per E101 conventions), at the right mounting elevation.
5. Tag each fixture with its appropriate tag types (elevation, letter code, panel/circuit). Panel/circuit will be blank until circuits are assigned — that is a later pass.
6. Report back to user with placement summary + ambiguities. **STOP before assigning circuits, drawing wires, or placing keynotes.**

## Things still uncertain (ask user when relevant)

- Should the new project's keynote NUMBERS match this project's numbering, or will the new project have its own legend?
- Is the new project on the same Revit/CED family library? (User implied yes when picking "different Revit project entirely" without flagging family availability.)
- How tightly should I match elevation tags? Cardio fixtures had varied heights (30/39/46"); the new equipment may differ by manufacturer.
- For multi-tagged fixtures (3 stacked tags), should I always place all three, or only the ones the legend implies the fixture needs?
