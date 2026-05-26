# Planet Fitness Corporate — Power Plan Design Playbook

A field-tested, project-agnostic synthesis. Use this on any new PF model.

## 0. First-touch checklist (read before placing anything)

1. **Identify the active power-plan view** — usually `E101 - Power Plan`. Use `doc.ActiveView`, never a hardcoded view name.
2. **Confirm panels are pre-placed** by the template. Note panel names: `L1`, `L2`, `L3`, `L4` and any `TR-L1`/`TR-L4`/`DP`. Note the **zoning rule**:
   - `DP + TR-L1 + L1 + L2` → utility/electrical room (101A_IT_ROOM or equivalent)
   - `L3 + L4 + TR-L4` → IT/data room
3. **Scope the building bbox** from the CAD link (e.g. ~`(40,-310,280,-185)` for Charlotte). DWGs include schedule callouts outside the shell; always filter by bbox.
4. **Find the yaml profiles file** — `c:\CED_Extensions\PF_profiles_V4_23.yaml` (or newer). The yaml is the source of truth for fixture family/type + electrical params per block name.
5. **Extract CAD blocks** to `<project>_blocks_v3.json` (name, x, y, rot_deg, layer) using the `GetSymbolGeometryId` → `SymbolId` → `Element.Name` pattern with the `X_BG.dwg.` prefix stripped.
6. **Extract wall segments** to `<project>_wall_segs.json` from the CAD `*WALL*` graphics-style categories.

## 1. CAD block name extraction (the single most important plumbing)

```python
import clr
_NAME = clr.GetClrType(DB.Element).GetProperty('Name')
def el_name(el): return _NAME.GetValue(el, None)

# inside the recursive geometry walker:
sgid = nested_geom_instance.GetSymbolGeometryId()
block_type = doc.GetElement(sgid.SymbolId)
block_name = el_name(block_type).replace("X_BG.dwg.", "")
```

- `GeometryInstance.Symbol` / `.ElementId` are NOT exposed. Reflection-named `Element.Name` is mandatory because plain `.Name` errors on `FamilySymbol` overloads.
- When recursing nested GeometryInstances, multiply `parent.Transform * child.Transform` exactly once per level. Double-multiplying silently shifts XYs by hundreds of feet.

## 2. Family/type mapping (one canonical list)

| Asheville name | Family | Type | Mount | Notes |
|---|---|---|---|---|
| Convenience recep | `EF-U_Receptacle_CED` | `Duplex Wall` | 18" | L3 STANDARD, 180 VA, Class R |
| GFCI vanity / wet area | `EF-U_Receptacle_CED` | `Duplex Wall - GFCI` | 18"/42" AC | L3 STANDARD, 180 VA |
| USB outlet (check-in) | `EF-U_Receptacle_CED` | `Duplex Wall - USB` | 18" | L3 |
| Quad above counter | `EF-U_Receptacle_CED` | `Quad Wall` | 42" AC | TMAX/IT/Reception |
| TV truss (mid-room) | `EF-U_Receptacle_CED` | `Quad Wall - TV` | ceiling | L3 |
| Hand dryer JB | `EF-U_Junction Box_CED` | `Wall - With Stem` | 42" | L3 |
| Heater JB | `EF-U_Junction Box_CED` | `Ceiling` | 96" | dedicated |
| Sauna / large equipment disconnect | `EF-U_Disconnect Switch_CED` | `Non-Fused - 60A` | side wall | 60A/3P or 30A/2P |
| Motor switch | `EF-U_Motor Rated Switch_CED` | `Motor Rated Switch - 120V, 1 Pole` | next to recep | HWH, water cooler |
| Specialty 240V/30A | `EF-U_Receptacle_CED` | `Specialty Wall - 240V/1Ph` | 18" | tanning/spray-tan/hydromassage |
| TV recep | `EF-U_Receptacle_CED` | `Duplex Wall` | 18" | dedicated, Class C, 300 VA |
| Floor recep | `EF-U_Receptacle_CED` | `Duplex Floor` | floor | massage chairs in BCS |
| Keynote | `Manual Key Note- All Shapes` | `Square` | annotation only | `CED-G-NOTE #` is **STRING** |

### Param-setting gotchas

- `CED-G-NOTE #` is **String** storage. Use `p.Set("11")` (or `"E1"` etc.), **never** `p.Set(11)`. Setting an int silently leaves it blank.
- `CED-G-NOTE TEXT` defaults to `"XXXX"` from family. Text isn't actually rendered in plans (sheet keynote legend pulls from a schedule), so leaving `"XXXX"` is harmless. Setting the real text on instance is nice-to-have but cosmetic.
- **CRITICAL — unit conversion on every project**: For double/measurable params, you MUST convert via `UnitUtils.ConvertToInternalUnits(val, du)` using `p.Definition.GetDataType()`'s unit-type. Otherwise yaml values silently land at 1/10.76 of their intended size (e.g. 180 VA in the yaml shows up as **16.72 VA** in the Revit panel schedule). Use `UnitUtils.IsMeasurableSpec(spec)` to gate which params need conversion. This applies to **every measurable double**: `Apparent Load Input_CED`, `FLA Input_CED`, `Voltage_CED`, mount heights, etc. **`Number of Poles_CED`, `CKT_Rating_CED` are integers** — pass them as integers without conversion.
- The single canonical `set_param(elem, name, val)` helper looks like:
  ```python
  def set_param(elem, pname, val):
      p = elem.LookupParameter(pname)
      if p is None or p.IsReadOnly: return False
      st = p.StorageType.ToString()
      if st == "String":
          p.Set(str(val))
      elif st == "Integer":
          p.Set(int(val))
      elif st == "Double":
          spec = p.Definition.GetDataType()
          if UnitUtils.IsMeasurableSpec(spec):
              du = doc.GetUnits().GetFormatOptions(spec).GetUnitTypeId()
              p.Set(UnitUtils.ConvertToInternalUnits(float(val), du))
          else:
              p.Set(float(val))
      return True
  ```
- For corner-mounted family instances (TVs etc.), the recep can be placed at the block XY with rot matching the CAD block's rotation (e.g. -45° or +45°). Default Duplex Wall faces north (rot=0).

## 3. Per-room placement playbook

### Cardio (treadmill / stairmaster / powered bike rows)

- One `Duplex Wall` recep behind each piece of **powered** equipment.
- **Equipment that gets a recep (powered)**: Treadmill, Stepmill (StairMaster), HM5A (indoor cycle). **Powered confirmation per project — these are the safe defaults.**
- **Equipment that does NOT get a recep (self-generating / unpowered)**: Elliptical, Arc Trainer (Arc Full), Bike Upright (INC-…), Bike Recumbent (INR-…). Confirm with user per project — some projects may differ.
- **Rotation**: **rot=0 (face north) for ALL cardio receps regardless of equipment block rotation.** The block origin sits at the front/south of the equipment; the equipment body extends north of the origin, so a recep at the block XY facing north points TOWARD the equipment center and back-panel power inlet. (Charlotte initially placed face-equipment-direction; Swannanoa user clarified that face-north is the correct rule everywhere.)
- **TV TRUSS row** — the truss IS NOT necessarily represented by a block named "TV TRUSS". In Swannanoa, the 12 "TV TRUSS" blocks are out-of-building schedule callouts; the actual TV positions are 24× small rectangle polylines (4.75 x 0.20 ft each) on the `A-N-TELEVISION` layer at a single Y, spaced 5 ft apart. **Authoritative source for TV row is closed polylines on A-N-TELEVISION, not blocks named "TV TRUSS".** Use `Quad Ceiling` family (Quad Wall - TV doesn't exist in newer projects), 10 ft mount, load `TV TRUSS`, 300 VA.
- **South-wall convenience receps**: place 3 along the cardio south wall at midpoints of solid wall segments (not in door openings). Asheville ratio is ~28 ft/recep, but on a wall with multiple openings, 3 is plenty.
- **Water fountain (ADA DF / `PF_Plan DF` profile)**: 2× `Duplex Wall - GFCI`, mount 18", **face north** (into cardio, since fountain mounts on the south wall and recep is in the wall behind it). Place GFCIs at the CAD water-cooler block XY but snap Y to the wall (1 ft south of wall line). Load name `WATER FOUNTAIN`, L2, BYPARENT, 180 VA, Class C. Add `#12` keynote 2 ft north.

### Locker rooms (104 womens / 105 mens)

- **TV (43tv profile, corner-mounted)**: at the `43 tv` CAD block XY (Y≈-284 in Charlotte, Y≈+10.55 in Swannanoa — opposite side of building). Block rot = ±45° or ±135° (corner mount). Match the block rotation on the recep so it visually aligns with the angled TV. Load `TV - BLACK CARD SUITE`, L3 DEDICATED, 300 VA, Class C.
- **`#11` keynote**: 4 ft from the TV recep into the room with a detail-line leader from keynote to recep.
- **Locker `#1 TYP`** (Charlotte pattern only): place the keynote **SOUTH of the receps inside the locker room**. Two fan-out leaders from keynote to each recep.
- **Dry vanity receps (TWO per locker)**: the locker rooms have a 6.17 × 1.67 ft wood-counter block drawn as a closed polyline on the **`A-FLOR-WDWK`** CAD layer (it is NOT a named block — search by polyline size on this layer). Place TWO `Duplex Wall` receps along the **wall side** (the long edge of the rectangle that's against a wall), at 1/3 and 2/3 of the X-span, **42″ mount**, **rot=0** (face into the room, toward room center in Y direction), load `RECEPTACLE - <#> LOCKER`.
- **Center-wall convenience receps**: there's a center wall dividing the storage rooms (105D/104D) that extends south into the locker corridor. Place **two receps on OPPOSITE FACES of this same wall** (~0.4 ft apart in X), facing west into MENS and east into WOMENS, **a few feet south of the storage south wall**. The user wants them very close together on the same wall, not on the lockers' separate exterior walls.

### Toilet rooms (104A / 105A) — Sloan trough sink pattern

- 1× `Locker-Sink-Sloan-Trough-Enclosure` block in CAD marks each trough center.
- **5× `Duplex Wall - GFCI`** vanity outlets at trough Y, X-offsets `[-50, -30, 0, +30, +50]` inches. Load `VANITY - <toilet#> TROUGH`.
- **2× `Junction Box - Wall With Stem`** hand-dryer JBs at trough Y minus 1'10" (≈Y-1.83), X-offset ±7 ft from trough center. Load `HAND DRYER - <toilet#> TROUGH`. ±7 ft is what the CAD hand-dryer drawings call for; the older `±72" + 2'3" outward + 1'3" inner/outer` formula was approximately right but the user prefers clean ±7 ft.
- **Keynote layout** (USER-CONFIRMED 2026-05):
  - `#1` keynote → **outer 2 vanity GFCIs only** (offsets ±50"). One keynote per toilet, north of vanity (~3.5 ft above), with 2 fan-out leaders.
  - `#3` keynote → **middle 3 vanity GFCIs** (offsets -30/0/+30"). One keynote per toilet, **SOUTH of the vanity** (~6 ft south of Y-trough), with 3 fan-out leaders. South placement avoids overlap with the `#1` keynote above.
  - `#2` keynote → **one per hand-dryer JB** (so 4 total across 104A+105A, not 2). Single leader each.
- Floor drains in CAD do not need a recep.

### Black Card Spa (BCS) — Swannanoa-era synthesis

Captured from a long iteration cycle in Swannanoa; supersedes the older Charlotte tier-D rules.

#### Identifying equipment in BCS
- **MEP Space centroids are placeholder positions** — do not use them as the room center. Use CAD `Room Name` block labels and equipment polylines to find the actual room geography.
- Equipment lives in two forms in CAD: **named GeometryInstances** (e.g. `HM62E Hybrid Tanning Booth`, `HM63C`, `PF Tanning 42-3 TLT PLT`, `SmarteCarte Massage Chair`) and **raw polylines on the `A-N-SPA EQUIPMENT` layer** for some items. Walk both.
- **TVs are usually polylines on `A-N-TELEVISION`**, not named blocks. The `65tv`, `70tv`, `60 TV`, `27TV` blocks in the building are out-of-building schedule callouts. The actual TV in BCS is a polyline. **Pick by width matching TV size**: 65" TV ≈ 4.75 ft wide, 43" ≈ 2.86 ft, 27" ≈ 2.25 ft. **Some BCS 65TVs are drawn as a square polyline footprint** (mount + screen, 3.36×3.36 ft) rather than a thin rect — when the obvious narrow polyline doesn't match the user's intent, look for square polylines in the room.

#### Block-origin vs visual-geometry-center gotcha
- For some CAD blocks (notably `SmarteCarte Massage Chair`), the **block origin is a schedule callout location far from the visual chair**. Always compute the geometry **bounding-box center** of the block's polylines when placing a recep "at the center of the chair". The origin alone can land the recep in the wrong room.
- Helper: walk the block's nested PolyLines, compute min/max X/Y of all points in world coords, use the midpoint.

#### Hydromassage row (HM63C → HYDROMASSAGE - 103A)
- Per chair: **Specialty Wall 208V/1Ph** (or 240V/1Ph if available) + **Duplex Wall - GFCI** touchscreen recep.
- **Snap all receps to the nearest long wall (typically the north exterior building wall)**, evenly spaced — equipment chairs are inside the room but the receps live ON the wall. This is critical: the chair X positions can land inside non-wall space, and receps must be wall-mounted.
- **One `#4` keynote per chair** (4 total, not a single TYP). Place each keynote at the chair's CAD block center.
- If the chairs' raw X spacing would overlap interior partition walls, shift them 2 ft along the wall to clear.

#### Tanning rooms (103B..103G) — disconnect placement
- For every tanning bed/booth (HM62E, HM62B Up Tanning, PF Tanning 42-3, BCS_Tanning Bed 42.4, BCS_Tanning Red Wave):
  - **Disconnect must be against a wall**, not in the middle of the room, not in the wall (offset ~4″ off the wall line into the room).
  - **Must not overlap the equipment CAD block** — if it would overlap, move it along the same wall (north/south or east/west) until it clears. Honor the **min 6″ from any perpendicular wall**.
  - **Handle parallel to the wall length** (the disconnect's long dimension aligns with the wall).
  - **Handle oriented toward the door** — this is a recurring rule. If the disconnect is on the west wall and the door is on the south wall, the handle's "down" end faces south.
- The convenience receptacle:
  - **Always faces into the room** (universal rule — no exceptions).
  - **On the wall with the door** when possible. If the door is east, place recep on east wall.
  - **Not in the doorway** — door swings are blocks on the `A-N-DOOR` layer (look for `Swing-90`) and door clearance areas are anonymous polylines named `*U241`. Treat both as no-place zones.

#### Specific tanning yaml profiles (apparent loads)
| Yaml profile | Disconnect | Apparent Load (VA) | Load name template |
|---|---|---|---|
| HM63C Hydromassage | (no disc) Specialty Wall 240V/1Ph + GFCI | 4320 + 180 | HYDROMASSAGE - <room> |
| HM62E Hybrid Tanning Booth | Non-Fused 60A, L4, 3P/240V | 11640 | HYBRID TANNER - <room> |
| BCS_Tanning Bed 42.4 | Non-Fused 60A, L4, 3P/240V | 11223 | TLT TANNING - <room> |
| BCS_Tanning Red Wave | Non-Fused 60A, L4, 3P/240V | 7067 | RED WAVE - <room> |
| BCS_Upright Tanning (HM62B_Up) | Non-Fused 60A, L4, 3P/240V | 10392 | STAND-UP TANNER - <room> |
| PF Tanning 42-3 TLT PLT | **Fused** 60A, L4, 3P/240V | 11223 | LAY-DOWN TANNING - <room> |

#### Cryolounges (BCS_Cryolounge profile)
- Duplex Wall - GFCI (or Duplex Wall), L4, **1440 VA** DEDICATED GFI Class C, load `CRYOLOUNGE - 103`.
- **Snap the recep to the nearest LONGEST wall segment** (not necessarily the closest wall — prefer longer over shorter when distance is similar).
- **Face into the room based on wall direction** (universal rule applied here):
  - West wall mount → face +X (rot=270 CCW)
  - East wall mount → face -X (rot=90)
  - North wall mount → face -Y (rot=180)
  - South wall mount → face +Y (rot=0)
- **Cryolounge polylines look like recliners** (~2.5×6.5 ft rectangles) but on `A-N-SPA EQUIPMENT`. **Not every elongated rectangle on this layer is a cryolounge** — when the user says "that's not a cryolounge, change to 180 VA convenience", it's just a standard wall-recep target.

#### Rinnai HWH (PF_Plan Rinnai HWH)
- 3-element profile: GFCI + Motor Rated Switch + Junction Box (Wall - With Stem).
- In CAD, the HWH units are drawn as raw polylines on `A-N-PLUMB FIX` (no named block). They typically appear as a **row of 3 polylines** (~1.4×0.8 ft each) on the wall.
- **Rule: 1 Rinnai HWH placement per 3 polylines, at the center one.**
- Orient based on wall direction: if polylines are on the **north wall**, the recep + JB face **south** (rot=180), and the motor switch is rotated to match. For a north-wall mount with the default motor-switch family at rot=90, **rotate it 90° CW to rot=0** so it sits parallel to the wall.

#### Keynote conventions (BCS tier)
- `#4` for hydromassage — **one per chair**, placed ON the chair CAD block center then moved 3 ft vertically off the fixture (so symbol isn't on top of equipment imagery).
- `#5` for all tanning beds/booths — one per room, placed near the equipment but not on top.
- `#11` for TVs (43tv, 65tv) — on the TV polyline center, then move ~2.5 ft toward the room interior so it sits in open floor.
- `#20` for SmarteCarte massage chairs — one per chair, placed at the chair center then **moved 2.5 ft north** (vertical offset).
- Cryolounge profile has no keynote in yaml.
- **Keynote-equipment association rule**: keynote goes ON the CAD block CENTER (not the recep, not a side offset). Then a separate "move N feet vertically" step puts it in clean floor space without losing the association.

#### Universal rules confirmed in Swannanoa Tier D
1. **Receps always face into the room** — no exceptions. If a recep would face the wall, rotate it 180°.
2. **Keynote symbols must not overlap CAD blocks** except the one being annotated (and even then, the user prefers a 2-3 ft offset from the equipment center for clarity).
3. **Handle orientation toward the door** is the universal disconnect rule.
4. **`*U241` polylines on `A-N-DOOR`** mark door clearance — treat as no-place zones.
5. **Anonymous polylines on equipment layers** can represent real equipment — always check named blocks AND raw polylines, on `A-N-SPA EQUIPMENT`, `A-N-TELEVISION`, `A-N-PLUMB FIX`, `A-FLOR-WDWK`.

---

### Black Card Spa (BCS, room 103 and sub-rooms 103G…103L) — Charlotte legacy

- **Tanning beds (`BCS_Tanning Red Wave`, `BCS_Tanning Hybrid`, etc.)** profile from yaml:
  - 1× Specialty Wall 240V/1Ph (5040–5760 VA) on door wall corner, side wall
  - 1× Disconnect (Non-Fused 60A, L4) **in the room corner closest to the bed's plug-end AND the door** — sit *against* the wall (not on/in it), at least 6" off any perpendicular wall.
  - Disconnects rotate to face into the room: NE corner → 90° CW (face W), NW corner → 90° CCW (face E).
  - 1× `Duplex Wall` convenience recep on door wall mid, face into room.
  - 1× `#5` keynote 4 ft into room from disconnect.
- **Sauna (103H, `BCS_Sauna Red Zone`)**: add an extra JB on the OPPOSITE wall mid (60" mount) + `#22` keynote 2 ft into room from JB.
- **Spray Tanning (103J)**: same pattern but use `Duplex Wall - GFCI` + `#21` keynote.
- **Hydromassage chairs**: `Duplex Floor` receps under each chair position.
- **Beauty Angel / cryolounge**: position recep near plug-end of bed (use bay relative-corner rule).
- **TV-BCS** in vestibule/BCS lounge: `Duplex Wall`, load `TV - BLACK CARD SUITE`, single keynote `#11`.

### Tier E + F — Front of House / Back of House (Swannanoa synthesis)

Captured from iterative refinement on Swannanoa. Supersedes Charlotte rules where they conflict.

#### Vestibule (100)
- **ECH-1 heater junction box** (Ceiling family) at the **center of the room's geometry** — NOT on any wall, NOT at the center of the north wall. MEP Space centroids and the wall data the CAD extractor finds typically only capture a sliver of the vestibule; the actual room is much larger than what the wall_segs json shows. **If you place the JB at the center of the wall data you found, expect to move it south by several feet to land in the actual room body.**
- 96″ mount, 1500 VA DEDICATED, load `VESTIBULE`, L3, classification M.
- **#21 keynote** 2 ft south of the JB.
- **Do not place wall receps in the vestibule** — they typically land in walls that aren't real or float in space. Skip them unless the user explicitly asks.

#### Check-in (102)
- **Computer area**: marked by a big polyline on the **A-N-FLOOR FIX** layer (~8 ft × 2.8 ft rectangle) that contains the `PF_Plan Computer` block positions. Inside the polyline there's a `wall-looking piece that has a corner` — that's where the computer receps go.
- **2× `Duplex Wall` receps** (the computers' direct recepts) on the floor-fix polyline edge facing **NORTH** (rot=0). The standard PF check-in faces north toward customers, and the reception desk to the south faces south. Use the TV's facing direction as your clue: TV in reception faces south → reception desk faces south → check-in counter faces north.
- **#7 keynote** for the isolated-ground / computer area (single TYP-style, no leader). Place 1.5–2 ft north of the receps.
- **No USB receps in check-in near #7** — USBs belong to reception per latest convention.
- **TV receps in check-in/reception**: TVs are A-N-TELEVISION polylines, sized by width:
  - 27" TV ≈ 2.25 ft wide
  - 43" TV ≈ 2.86 ft wide
  - 60" TV ≈ 5.00 ft wide
  - 65" TV ≈ sometimes drawn as a square footprint (3.36×3.36 ft mount + screen) instead of a thin rect
- Each TV recep is `Duplex Wall` 300 VA dedicated, load `TV - BLACK CARD SUITE`, faces into the room (rot=180 if on north wall).
- **`#11` keynote 2.5 ft AWAY FROM the TV in the direction the TV is FACING** (i.e. 2.5 ft into the room from the TV).
- **Phone JB** at `A-N-FIXT-PHONE` polyline position, `Wall - With Stem` JB family, `RECEPTION JB` load, faces west (rot=90). **#8 keynote** at the JB location adjusted so the recep+keynote sit flush with the short wall segment beside the phone fixture.

#### Reception (101)
- **`PF_Plan TMAX` profile is BIG**: 5× `Duplex Wall` receps (`TMAX RECEPT`, L3 180 VA each, DEDICATED, Class C) + 1× `Junction Box - Wall With Stem` (`SIGNAGE (IF NEEDED)`, L3 180 VA, Class L) + 1× `#8` keynote.
- Place the TMAX cluster at the **center of the polyline on the `A-FLOR-OVHD` layer** (typically a ~15×2 ft above-counter zone, marking where these receps are mounted above the counter).
- If `A-CLNG-PATT` has a FACE region (not a polyline), TMAX goes there. Faces can't be walked with the standard PolyLine walker — would need a separate Face/Region collector. **In Swannanoa CLNG-PATT had no face, so use A-FLOR-OVHD polyline center instead.**
- **6× USB receps near `#15` keynote** in TWO ROWS of THREE:
  - Bottom row (3 USBs) on south wall (Y on the wall line), rot=0 face N
  - **Top row (3 USBs) is the ONLY EXCEPTION to the universal wall-snap rule** — it does NOT need to be on a wall. This is because this is a 2D drawing and the heights will be tagged later. Top row still faces north (all 6 USBs face same direction).
  - X positions evenly spread along the reception counter (use the A-N-FLOOR FIX polyline span).
- `#15` keynote on the **southern part of reception** (south of the USB cluster).

#### Breakroom (101A)
- **One receptacle on EACH wall, unless the wall segment is shorter than 5 ft.** Check all four walls (N/S/E/W), place a GFCI on each ≥5 ft wall.
- Receps face into the room (perpendicular to wall, pointing inward).
- **`#16` keynote ONLY goes with vending-machine receps** — NEVER with a convenience GFCI. Convenience receps have no keynote.
- **Breakroom can be much bigger than the local wall data suggests.** In Swannanoa the west wall is actually at X=143.13 (a shared wall with IT room east), even though the wall_segs.json only captured the X=143.13 portion in the Y=17..26 range. The room extends ~9 ft further west than the small north-end stubs (X=150.96/151.56) suggest. When you snap a recep to a "west wall" find the LONGEST vertical wall in/near the room's expected X range, not just the closest fragment.

#### IT Room (101B)
- **IT cabinets/racks are polylines on `A-N-ELECTRICAL` layer** (in this project, ~3 ft × 1.7 ft rectangles, sometimes stacked vertically along a wall).
- The "IT polyline square" the user refers to is the largest A-N-ELECTRICAL polyline in the room (~7×5 ft) that defines the IT cabinet zone.
- **Quad Wall receps go on the wall directly in front of the IT polylines** — if cabinets are stacked vertically at the SW corner, the quads go on the **west wall**, rot=270 (face east toward the cabinets), at the Y positions of the cabinets.
- **`#9` keynote 3 ft east** of the quads' X position (into the room, between the quads vertically).
- One convenience recep on the **door wall** (typically east wall near the door swing block, not in the door clearance area).

#### Vending Machines
- Found on **`A-N-BEVERAGE COOLER` layer** (NOT `A-N-VENDING`). Polylines ~3×3 ft each.
- Each vending recep gets its **OWN `#16` keynote** (not one TYP for the pair — discrete keynotes per vending unit).
- **Recep snaps to the wall opposite the direction it's facing** — body sits AGAINST the wall, face points into the room.
- **`#16` keynote sits ~3 ft IN THE DIRECTION the recep is facing** (i.e. 3 ft into the room from the vending recep, in the direction the recep faces). This is the "vending keynote offset" standard.
- 42″ mount (above counter), GFCI, load `VENDING MACHINE`, L3 STANDARD 180 VA.

#### TMAX cluster (PF_Plan TMAX)
- yaml profile lists 5× Duplex Wall + 1× JB Wall-With-Stem (`SIGNAGE (IF NEEDED)`) + #8 keynote — **DO NOT place the JB Wall-With-Stem.** The signage JB is not needed in actual placements (user-confirmed).
- The A-FLOR-OVHD polyline marks the **mounting zone** (typically 15 ft × 2 ft, above the reception/check-in counter). **It is NOT a wall.** TMAX receps must **snap to actual building walls** in the same X positions — the OVHD polyline tells you where the cluster lives, the walls tell you the Y to anchor.
- Place **up to 5× `Duplex Wall` receps** evenly spaced along the OVHD polyline X span (~3 ft apart). Then move each recep north (or appropriate direction) to snap to the building wall behind it. Different receps may land on different paired walls (e.g. one wall at Y=16.56 covering middle X, another wall at Y=17.16 covering wider X).
- **If a TV polyline (A-N-TELEVISION) sits at the center of the OVHD polyline, REMOVE the center (3rd) TMAX recep** — the TV occupies that position. Only 4 receps placed.
- **All TMAX receps face SOUTH** (rot=180) — match the TV-facing convention for this front-of-house area. The TVs in check-in/reception face south, so all wall receps in the same north-wall mounted band also face south.
- 42″ mount, L3 DEDICATED, 180 VA, Class C, load `TMAX RECEPT`.

#### Vestibule receps (don't replace)
- **Do not place wall receps in the vestibule** at all — the vestibule's actual room extent is larger than the wall data captured, and any wall snap will likely float. The ECH-1 ceiling JB is enough.
- The ECH-1 JB sits at the actual room center, which may be 8+ ft south of where the wall data suggests. Default placement should be conservative — start at MEP centroid Y and adjust based on what the actual room looks like.

### Universal rule — placement vs CAD geometry (added on Swannanoa Tier G review)

**Convenience receptacles MUST NOT overlap any CAD polyline or block** in the building underlay, with the SOLE EXCEPTION of TV polylines (where the recep is intentionally placed ON the TV symbol). When placing a perimeter convenience recep, after picking the wall position, check whether the chosen XY collides with any equipment polyline (A-N-GYM EQUIPMENT, A-N-SPA EQUIPMENT, A-FLOR-WDWK, A-N-LOCKERS, A-N-FIXT-CHECK-IN, A-N-BEVERAGE COOLER, A-N-PLUMB FIX, etc.) within ~1 ft, and if so, shift the recep along the wall to a clear position.

### FTMR rooms can be multi-piece (Swannanoa)

The MOBILITY ROOM FUNCTIONAL TRAINING (107 FTMR) is not always a single contiguous polygon. In Swannanoa it covers:
- A large **south box** (the main FTMR area, X=-25..13, Y=-55..-25)
- A **middle box** north of MENS LOCKER ROOM (X=-25..0, Y=11..22)
- Sometimes an unlabeled corridor connecting them along the east boundary (X=-0.41 wall)

When the user circles the room boundary on screen, treat ALL the orange-circled area as part of FTMR — including any disconnected pieces. Walk perimeter walls of each piece separately for convenience-recep placement.

**Easy room mis-identification**: north of the FTMR middle box (Y=22+) is the **ELECTRICAL ROOM 107A**, not FTMR. If you place receps at Y=22..30 thinking it's FTMR, you'll land them in 107A. The middle box is **south** of the Y=21.97 wall, not north. Always check `Room Name` block label X/Y to know where the boundary is.

### Tier G (perimeter convenience receps) — gym rooms (Swannanoa)

The gym rooms (FTMR 107, Strength 110, Circuit 109, Free Weights 111) are **mostly open** — few interior walls, mostly bounded by building exteriors and a few partition walls (e.g. the X=13 divider between FTMR and Strength in Swannanoa).

#### Workflow per gym room
1. Identify the room's bounding walls (typically just 1–2 confirmed walls per room).
2. For each confirmed wall ≥ 3 ft, place a `Duplex Wall` recep at a clear position (not overlapping any equipment block/polyline within ~1 ft).
3. **Do NOT place receps at unconfirmed wall positions** — if there's no wall in the wall_segs json at a guessed Y, the recep will float and get rejected. Better to leave a gap than place wrong.
4. **"Both ends" rule for long single-wall corridors**: when a room has only one long wall (e.g. a long east wall along the corridor between FTMR-middle and the building's south FTMR), put receps at BOTH ENDS of that wall (north end + south end) rather than one in the middle.
5. **Check each room for TVs**: search A-N-TELEVISION polylines in the room bbox. If a TV polyline is found and there's already an `#11` keynote nearby, just place a `Duplex Wall` recep ON the TV polyline (rot to face into room). Use TV size to determine voltage class (300 VA for 43/55/60/65/70/75 TV, all on L3 DEDICATED, Class C, load `TV - BLACK CARD SUITE`).
6. **The cardio TV TRUSS row at Y=-39.8 (24 receps spanning X=23..138) covers FTMR and Circuit too** — don't double up TV receps in these rooms if they overlap the truss row.

#### Universal: don't overlap CAD geometry (TV exception)
**Convenience receptacles MUST NOT overlap any CAD polyline or block** in the building underlay. **TVs are the SOLE exception** — a TV recep is intentionally placed ON the TV polyline. For every other recep placement, after choosing wall XY, check for any equipment polyline within ~1 ft and shift along the wall if it collides.

#### Universal rules added in Tier E+F
1. **Receps always face into the room** — including vending. The wall it's mounted on is BEHIND the recep.
2. **2D drawing exception**: a counter-cluster row (like USB row 2 in reception) may sit off the wall — heights will be tagged later by the user. This is the ONLY exception to the wall-snap rule.
3. **`A-N-FLOOR FIX`** = ground-level counter/desk polylines (computers, reception desk).
4. **`A-FLOR-OVHD`** = above-counter mounting zone (TMAX cluster goes here).
5. **`A-N-BEVERAGE COOLER`** = vending machines.
6. **`A-CLNG-PATT`** = ceiling pattern FACES (not polylines) — needs a face walker to use.
7. **Use the TV's facing direction** as a clue for "which way is into the room" — TVs face into the customer area in front-of-house spaces.
8. **`#16` keynote is vending-only** — never on a convenience recep.

---

### IT room (101A) + Reception cluster — Charlotte legacy

- IT room: 2× `Quad Wall` receps at counter height, facing into room (often into hallway/IT side). Load `IT RACK`, L3.
- Reception/check-in cluster: a row of `Duplex Wall` + `Duplex Wall - USB` along the front-of-counter and back-wrap. Includes TMAX receps (special), Check-in recepts, USB Check-in.
- The `JB` for Reception (often a phone/data JB) goes near the printer/computer cluster — don't double-stack with TMAX recepts at the same XY.

### Vestibule (100) + ECH-1 heater

- 1× `Junction Box Ceiling` at door-side ceiling for `ECH-1 vestibule heater`, load name `VESTIBULE`, 1500 VA, DEDICATED, mount 96".
- 1× `Duplex Wall` TV-BCS recep on east wall (face W into vestibule), 18" mount.
- 1× `Duplex Wall` convenience recep on east wall (face W), 18".
- **Important quirk**: the yaml profile for `FAN 1` says keynote `#18`, but the Asheville legend has `#18 = clock`. Use `#19` for fans (CONFIRMED via Asheville legend). Don't trust the yaml `keynote_num` field blindly.

### Functional Training Mobility Room (FTMR, 107)

- Often an L-shape with cutouts for landlord areas (HOUSE/UTILITY back-of, EX'G ELECTRICAL).
- **DO NOT place receps in landlord areas.** Identify landlord rooms by Room Name CAD labels and exclude their wall span from the perimeter walk.
- Build the FTMR polygon from the union of FTMR equipment + 2 ft inflation, snapped to CAD wall segments. This auto-excludes landlord areas (no FTMR equipment there).
- Walk perimeter at ~28 ft/recep (Asheville ratio). Skip any segment without a confirmed CAD wall.

### Strength / Free Weights / Circuit areas

- 2-3 `Duplex Wall` convenience receps along north wall of each room. Use room bbox + perimeter walk.

### Breakroom / Storage / Janitor

- Breakroom: 3× GFCI receps for vending + counter + microwave/coffee.
- **Storage closets (104D/105D/101D)**: 1× recep per closet, **on the wall the door is cut from, but NOT in the doorway**. Find the door swing arc block (`Swing-90`) and the door-frame X positions on that wall. The wall typically has two small solid stubs on either side of the door — place the recep on the WIDER of the two stubs, facing into the room (rot=0 toward room center in Y direction). For very narrow stubs, this puts the recep near the corner of the room beside the door.
- **Janitor (104C/105C)**: 1× GFCI on the side wall away from the storage rooms (in Swannanoa: 105C on west wall facing east, 104C on east wall facing west — the side walls are the long walls of the narrow janitor closets). **MEP Space centroids are unreliable for janitors** (they're often placeholder positions); verify against the CAD Room Name labels.

## 4. Universal placement rules (recurring user feedback)

### Wall-snap and "no wall, no recep"

- For every wall-mounted recep, verify it sits within 1 ft of a confirmed CAD wall segment. If not, **skip the placement** — never place a floating recep just to satisfy a count target.
- Use `<project>_wall_segs.json` (extracted from CAD `*WALL*` layers) as the truth source.

### Rotation conventions (default `Duplex Wall` faces north)

| Wall it's on | Recep faces | rot_deg |
|---|---|---|
| North wall | south (into room) | 180 |
| South wall | north (into room) | 0 |
| East wall | west (into room) | 90 |
| West wall | east (into room) | 270 |
| NE-corner 45° | SW direction | 45 |
| SE-corner 45° | NW direction | 135 |
| SW-corner 45° | NE direction | 315 (or -45) |
| NW-corner 45° | SE direction | 225 |

### Disconnect placement (sauna / tanning / spray-tanning / hydromassage equipment)

1. Side wall (E or W), nearest to the equipment plug-end AND nearest the door.
2. Up against the wall (not on it).
3. At least 6" clearance from perpendicular walls.
4. Rotate to face into the room: NE corner → 90° CW; NW corner → 90° CCW.

### Keynote placement

- **At least 2 ft** from the associated recep — 1 ft causes symbol overlap at typical plan scale.
- For receps facing a cardinal direction, place keynote 2-4 ft in that direction (so leader goes straight).
- **Leaders are not the default — most keynotes don't need a leader.** Only add a leader when the association would be ambiguous without it (e.g. fan disconnect buried inside equipment cluster, or vanity vs hand-dryer JBs in the same toilet). TYP cluster keynotes for cardio rows (#13) typically have no leader.
- **Keynote graphic must not overlap ANY CAD block except the one it's annotating.** Place at the side of the equipment cluster (west/east of cardio if equipment fills the middle) so the keynote symbol sits over empty floor. With a leader, the keynote can be far from its target — proximity is not required.
- For **TYP cluster** keynotes (multiple receps share one keynote): place the keynote at a reasonable common point. Add fan-out leaders only if the user explicitly wants them.
- For a single keynote covering many like items (e.g. 24 TV TRUSS receps → one `#14` TYP, no leader), place near one end of the row in empty floor area.
- Keynote `#3` for the toilet-vanity middle-3 goes SOUTH of the vanity (not north, to avoid overlapping `#1`). Generally: when 2 keynotes share a recep cluster, put them on opposite sides.
- Don't confuse Sloan trough vanity outlets (in toilet rooms 104A/105A at Y≈-285) with locker-room receps (in 105/104 at Y≈-267) — they're DIFFERENT rooms despite overlapping label conventions.

### Asheville keynote-number map (use as default; verify per project)

| # | Use |
|---|---|
| 1 | Vanity GFCI cluster (outer 2 of trough), or generic TYP cluster in lockers |
| 2 | Hand dryer JB (one per JB, not one per pair) |
| 3 | Vanity middle-3 GFCIs (south of vanity) |
| 4 | Hydromassage 208V NEMA L6-30 + 12V touchscreen |
| 5 | Tanning/Hybrid bed 240V NEMA 6-30 + 60A/3 NF disconnect |
| 6 | Beauty Angel 240V |
| 7 | Isolated ground for Check-in IT recep |
| 8 | Phone/RJ45 jack for FA/MGR; data to tanning |
| 9 | Data rack / AV vendor coord |
| 10 | (varies — verify) |
| 11 | TV mounting height/connection |
| 12 | Water fountain (PF_Plan DF) |
| 13 | Power & data raceway + recessed J-box for floor routing |
| 14 | (varies) |
| 15 | Mount recep/USB ports horizontally above counter |
| 16 | Vending machine |
| 19 | **Coord with fan MFR** (NOT #18 — yaml is wrong) |
| 20 | False column raceway — drop wire down for cardio raceway |
| 21 | Power for ECH-1, see L1 schedule (spray tanning room) |
| 22 | Sauna opposite-wall JB |

## 5. Workflow phases (sequence matters)

| Pass | What | Status by end |
|---|---|---|
| 1 | Place all electrical fixtures + keynotes (this playbook) | All rooms done, no circuits |
| 2 | Circuit all fixtures (panel/circuit assignments) | CKT params filled, schedules populate |
| 3 | Place wires (homerun arcs from fixtures to panels) | Wire elements drawn |
| 4 | Wire tags | Tags attached to wires |
| 5 | Fixture tags | Tags attached to fixtures |

- Until Pass 2, leave `CKT_Circuit Number_CEDT` at `STANDARD`/`DEDICATED`/`BYPARENT` per yaml profile and `CKT_Panel_CEDT` at the yaml's panel default.
- The user prefers we batch Pass 1 entirely (all rooms) before starting Pass 2.

## 6. Per-room manifest pattern (project-local scaffolding)

For each room, create folder `rooms/<NUMBER>_<NAME>/` containing:

- `block_map.json` — input plan (fixture family/type, profile params, target XYs/rotations).
- `place.py` — generic placer that reads block_map.json, opens a transaction, places fixtures, writes manifest.json.
- `manifest.json` — output receipt (element IDs, XYs, target metadata).

The manifest is **per-project scaffolding** — not portable. The reusable IP is this playbook + the place.py template + the yaml profiles file.

## 7. Operating constraints from the user

- **Only operate via MCP on the views the user has open** — never switch active views unprompted.
- **Don't place anything beyond electrical fixtures + generic-annotation keynotes** unless explicitly asked. Lighting fixtures, data devices, plumbing fixtures are out of scope.
- **Don't place receps in landlord areas** (HOUSE/UTILITY back-of, EX'G ELECTRICAL rooms).
- **Sloan trough hand-dryer JBs**: when in doubt, use ±7 ft from trough center (mirror the 104A spacing). Don't use the older ±9.5 ft outer placement — that walks the JB off the hand-dryer CAD imagery.
- **Cardio receps**: place at center of equipment block then shift north ~10" so they sit behind the equipment back-panel.
- **Leader lines** (CurveElement detail lines) are used to associate a TYP keynote with multiple targets. Generic annotations don't have native leaders.
- **Co-authored commits**: never include "Co-Authored-By: Claude" footer.

## 8. Quick-start command for a new project

```
1. Open the model. Confirm active view = E101 (or equivalent power plan).
2. Run extractors:
   - extract_blocks.py  → <project>_blocks_v3.json
   - extract_walls.py   → <project>_wall_segs.json
   - extract_rooms.py   → <project>_rooms.json
3. Read PF_profiles yaml, build family/type symbol lookup once.
4. For each room in placement order (Cardio → Strength → FW → Circuit → Locker rooms → Toilets → BCS sub-rooms → FTMR → IT room → Reception → Vestibule → Breakroom → Storage/Janitor → DF):
   a. Compute placements per §3 of this playbook.
   b. Verify wall-snap.
   c. Place in a single transaction.
   d. Write per-room manifest.json.
5. Audit: total fixture count, any floating receps (>1 ft from wall), any landlord-area placements.
6. Hand back to user for Pass 2 sign-off.
```

## 9. Things that bit me (don't repeat)

- Trusting yaml `keynote_num` for FAN 1 (yaml says 18; truth is 19).
- Setting `CED-G-NOTE #` as int — silently empty result.
- Looking at wrong Y band (locker-room receps at Y=-267 vs toilet vanity at Y=-285) and concluding the trough was never placed.
- Placing FTMR receps using a rectangular bbox that overshot into landlord areas. Use equipment-inflated polygon, not bbox.
- Trusting per-block "powered" detection on dense cardio rows. Place every block in dense rows.
- Placing a TV recep at one XY when block has 2 CAD instances (e.g. dual-height ADA water fountain needs 2 GFCIs, not 1).
- Forgetting to delete old leader detail-lines when moving a keynote — leaders are separate `CurveElement` objects and must be cleaned up explicitly.
- Searching `OfClass(DetailLine)` — DetailLine isn't a queryable class in Revit; use `OfClass(CurveElement)` instead.
