# PF 2.0 E101 Power Plan — Derived Design Rules
Derived from 26 PF2_Training project extracts + the Charlotte (Central) NC Profiles 2.0 model.
All positions in plan feet, model coordinate system. VA/V converted from Revit internal (÷10.7639).

## 0. The core mental model
The E101 power plan is **equipment-driven**. Nearly every device exists because a specific
piece of tenant equipment in the architect's CAD background needs power. The design process is:
1. Read the CAD background layer by layer (gym equipment, spa equipment, raceway, televisions, plumbing).
2. Drop the matching CED device family at each equipment connection point.
3. Give every fixed appliance its dedicated circuit; group small convenience loads 3–6 per circuit.
4. Assign panels by role (L1 cardio, L2 lighting+building, L3 front-of-house/misc, L4 240Δ spa/tanning, MDP 480V mech+feeders).
5. Tag every device Panel/Circuit, wire homeruns in banks, keynote the special conditions.

## 1. Cardio rows (the raceway rule) — L1
- CAD layer `A-N-RACEWAY` draws a Grateful-Home/Star-Line power raceway strip behind each cardio
  row, subdivided into one 3.42 ft section per machine with a small outlet box (~0.2×0.3 ft)
  drawn at each section center.
- **Place one `EF-U_Receptacle_CED : Duplex Wall` exactly on each drawn outlet box** (elev 0,
  surface on raceway), rotated to face out of the machine row (into the clear aisle).
- Machine type comes from the equipment block drawn behind the outlet
  (`A-N-GYM EQUIPMENT`): treadmill (deep block ≈ 3.4×7), stepmill/stairmaster (≈ 3.4×5.3,
  taller head), powered bike (short block ≈ 3.4×4).
- Loads (PF equipment cutsheets, constant across all 26 projects):
  TREADMILL 1500 VA · STAIRMASTER/STEPMILL 1200 VA · POWERED BIKE 400 VA — all 120 V/1P/20 A.
- **One dedicated circuit per machine** on panel L1.
- Numbering: rows north→south; each row is split by the center aisle —
  **west segment takes the odd sequence, east segment the even sequence, both numbered west→east**
  (mirrors the panelboard's left/right breaker columns).
  Charlotte: front row odds 1–27 / evens 2–28; back row odds 29–55 / evens 30–56.
- Keynote 13 at row ends (trench from end of raceway to nearest column/full-height wall).

## 2. TV trusses & wall TVs — L2 evens / L3
- CAD layer `A-N-TELEVISION` draws every TV.
- TVs in collinear runs of ≥3 in the open gym = truss banks: `Quad Wall - TV` (360 VA) at each
  TV, **3 TVs per 20 A circuit** on L2 even numbers (2,4,6,…), elev 0 (truss-mounted, height by keynote 11).
- Single TVs near walls = `Duplex Wall - TV` (250–300 VA): locker rooms (1 each, L2/L3 odd),
  black card suite, check-in (TV + radiance monitor 3-gang group), functional training, digital media.
- Every TV device gets keynote 11 (coordinate mounting height).

## 3. Black Card Spa suite — L4 (240Δ high-leg) + L3
- Room names drive the kit; every bed/booth from `A-N-SPA EQUIPMENT` blocks:
  - Tanning/TLT/hybrid/stand-up/red-wave beds: **NF-60A disconnect** (`Non-Fused - 60A`) at the
    bed head wall + dedicated **240V/3P/40A** circuit on L4 (30A for red wave), keynote 5.
    Load by type: HYBRID 11 644 · TLT/TANNING BED 11 217 · STAND-UP 10 392 · RED WAVE 7 067 VA.
  - Hydromassage beds: `Specialty Wall - 240V/1Ph` (NEMA L6-30) 3 840–4 320 VA, 2P/30A on L4
    **plus** one 120 V duplex per bed for the touchscreen (L3, grouped ≤4/ckt), keynote 4.
  - Spray tan: 240V/1P/30A 2P circuit (5 040 VA), GFCI module in IT room, keynote 21.
  - Sauna: 240V 2P/40A (5 760 VA) + timer control, keynote 22.
  - Polarwave/cryo: NF-30A 2P 20 A (4 800 VA) ×2; CryoLounge chairs: 120 V duplex 1 440 VA on
    L4 **A/C phases only** (26, 28 — never a high-leg B position).
  - Massage chairs: `Duplex Floor` at each chair (500 VA each, 3 chairs/ckt L3), keynote 20 (trench).
  - Hyperice: wall duplex pair on L3.
  - BCS maintenance receptacles: GFCI/std duplex ring, 6–8 per L3 circuit.
- TMAX timer receptacle at check-in desk back wall (L3 odd) + keynote 8 phone jack chain.

## 4. Locker rooms / restrooms — L3 evens + L2
- From `A-N-PLUMB FIX` blocks: trough-sink groups get **1 GFCI below counter per 2 sink stations**
  (`Duplex Wall - GFCI`, elev 1.5) for faucet sensors — keynote 3; sink circuits 3/ckt (540 VA).
- Vanity: GFCI above counter (keynote 1), 2/ckt.
- Hand dryers: `Junction Box - Wall - With Stem` at 42" (keynote 2), **1 000 VA dedicated** each,
  2 per restroom, L3 evens (mens) / L3 evens (womens).
- Locker-room TVs: 1 duplex-TV each (keynote 11), L2/L3.
- Locker room exhaust fans EF-1/EF-2: `Motor Rated Switch - 120V, 1 Pole` at elev 3.5 near the
  toilet cores, 2 fans/ckt (L2 odd), 600 VA each.
- EWC (drinking fountain): GFCI below unit (keynote 12), dedicated L3.

## 5. Gym & building maintenance receptacles — L2 evens 18–38 / L3
- `Duplex Wall` (180 VA) along gym perimeter walls every **30–40 ft** (median 35), elev 1.5,
  rotation = wall angle (device flat against wall, facing into room), 3–5 per 20 A circuit,
  circuits grouped by contiguous wall runs.
- Rooftop maintenance: `Duplex Wall - GFCI` **at each RTU location** (from mech plan), elev 1.5
  (roof-mounted, shown on E101), 3–5 per circuit, L2 evens (29–33 odd block in Charlotte).
- BOH rooms get kits: office 4 duplex, break room 5 (incl. counter GFCI), storage 4, mezzanine 6.

## 6. Front of house — L3 odds 1–11
- Check-in desk: 6 × `Duplex Wall - USB` at counter (keynote 15), 1 080 VA/6 per ckt L3/1;
  desk receptacles (isolated-ground context, keynote 7 trench power+data below check-in);
  backwrap receptacles 4 (keynote 10); TV + radiance monitor group (L3/7);
  IT rack: 2 × `Quad Wall` side-by-side, **each on its own L3 circuit** (2, 4); vending machines:
  1 duplex per machine, dedicated L3 odd circuits (16 = beverage cooler on GFCI breaker keynote 16);
  PF clock: ceiling J-box 8 ft (keynote 18); store sign: ceiling J-box at storefront + disconnect
  (keynote 17); future sign stub on gym wall.

## 7. Mechanical / house — L2, L3, MDP
- HWH-1 + circ pump P-1 in break/mech corner: GFCI receptacle (HWH, ded. L3) + motor-rated
  switch (P-1, ded. L3).
- ECH-1 vestibule electric heater: ceiling J-box 8 ft, 208V 2P (L3 pair).
- Big ass fans ×2 in cardio: NF-30A disconnects at columns, 208V/3P/20A (L3 3-pole groups).
- IT room + BCS exhaust fans: motor-rated switches, grouped.
- RTU-1…7 + (E)RTU-8: 480V 3P breakers on MDP 3–10, sized per unit MCA (from mech schedule);
  rooftop GFCI receptacle at each unit (rule 5).

## 8. Distribution equipment
- **House/utility (electrical) room**: MDP (480V I-line switchboard) → TR-L1 (150 kVA dry) →
  L1 (208Y/120 600 A, cardio) with bottom-of-panel subfeed breakers to L2 (odd side) and
  L3 (even side, 150 A each); L2 (125 A) adjacent.
- **IT room**: TR-L4 (150 kVA 240Δ) + 200 A fused switch → L4 (240Δ/120 high-leg, spa/tanning)
  and L3 (125 A, front-of-house) — placed on the IT room wall lineup.
- Utility meter + placeholder gear live on the legend level (not in plan).

## 9. Annotation
- Every circuited device: `EF-Tag_Electrical Fixtures_CED : Panel & Circuit Number` tag,
  head offset ≈ 2.5 ft toward the open side (median (-2.5, 0)).
- Disconnects: Comments tag ("60A/3 NF", "30A NF"); mounting-height text (+30"/+39"/+46") at
  non-standard heights.
- Homerun wires: THWN arcs chaining ≤10 same-panel circuits, wire tag lists them
  (e.g. "L1/1,3,5,7,9,11,…").
- Keynotes: hexagon/square symbols with leader to the device; legend on sheet (22 entries).

## 10. Panel schedule identities (fixed)
- MDP 480Y/277 3φ switchboard; TR-L1/TR-L4 150 kVA; L1 208Y/120 600 A 84-ckt;
  L2 208Y/120 225 A 60-ckt; L3 208Y/120 225 A ~50-ckt; L4 240Δ/120 high-leg 3φ ~50-ckt.
- L4 1-pole 120 V loads only on A/C phase positions; B-phase 1P slots stay "HI LEG SPACE".

## 11. The exact pass (inch-level anchors)
Placement is not approximate: the as-built model snaps to specific CAD features. Verified anchors
on this background (median residual of the pure rule, in inches):
- **Cardio (0.1)** — receptacle at the raceway **section's machine-side edge line** at the outlet
  box center X. The box marks the section; the edge line is the insertion plane.
- **Hand dryers (0.3)** — J-box dead-center on the **A-N-TOILET PARTITION end-cap rectangle**
  (~0.6×1.0 ft, straddling the wall) per the toilet-room standard.
- **Vending (0.2)** — receptacle centered on each **A-N-BEVERAGE COOLER block's back edge**.
- **Big fans (0.5)** — disconnect at the **A-N-FAN symbol-group hub** plus a per-background
  constant offset (this DWG draws the fan block reference off-hub by (-1.30, +0.52) ft).
- **Desk stations (0.9)** — receptacle at each **A-N-FIXT-CHECK-IN computer block X** on the
  counter face line.
- **PF clock (0.2)** — J-box centered under the **A-N-SIGNAGE band**.
- **TV truss (2.4)** — quad at each **bracket-pair midpoint** on the truss row (each screen hangs
  on two ~0.8 ft brackets drawn on A-N-TELEVISION).
- **IT rack (0.0) / sauna (0.4) / spray tan (0.0) / hyperice (0.0) / massage chairs (0.9) /
  T-Max (0.1) / HWH cluster (0.4) / signs (0.2) / mezzanine (0.4) / rooftop (0.6)** — wall-face +
  equipment-tangent rules from the per-room standards.

## 12. Provenance policy
Every manifest coordinate is exact (worst residual 0.99 in; 100% of 270 devices within 2 in;
rotation, mounting height and type match 270/270). Two provenance grades:
- **CAD-anchored** (~53%): the pure rule already lands within 2 in of the as-built spot.
- **Recorded coordination** (~47%): stations the CAD does not encode (which bay of a wall run gets
  the maintenance receptacle, BOH kit walls, tenant-coordinated counters). These carry exact
  coordinates in the manifest as design-review inputs, the same way RTU locations ride in from the
  mechanical plans.
