# PF Electrical Power auto-replication — Learnings log

This file records rules, gotchas, and user feedback discovered while building the auto-replicator. Treat it as living. Each rule has a brief justification so the future agent can decide whether the rule still applies.

## Project setup observed (Charlotte Corporate Takeover)

- Active view: `E101 - Power Plan`, view_id ends differing between models — always use `doc.ActiveView`.
- Distribution panels are pre-placed in the template by the user. Zoning rules per user:
  - **DP + TR-L1 + L1 + L2** → electrical/storage room near utility XFMR.
  - **L3 + L4 + TR-L4** → IT room.
  - This project has **no `(E) UTILITY XFMR`** (takeover reusing existing service).
- Spaces in the view are unplaced placeholders (8×6 ft bbox, area=0). **Cannot use `Space.IsPointInSpace` or space bbox for spatial filtering.**
- 0–4 architectural walls in document (the shell lives in an unloaded linked Revit model). **No face-hosted fixture placement available.**
- `EF-U_Receptacle_CED :: Duplex Wall` has `FamilyPlacementType = OneLevelBased` despite its name — use `place_workplane_fixture(doc, sym, level, pt, rotation_rad)`, not face-hosted.

## IronPython gotchas (Revit MCP `execute_revit_code`)

- `DB.Element.Name` is an ambiguous overridden property on `FamilySymbol` etc. — `DB.Element.Name.GetValue(sym)` throws `AttributeError`. Use the reflection pattern:
  ```python
  import clr
  _NAME = clr.GetClrType(DB.Element).GetProperty('Name')
  def el_name(el): return _NAME.GetValue(el, None)
  ```
- `GeometryInstance.Symbol` is **not** exposed in this Revit version (no `.Symbol`, `.ElementId`, or `.Reference` property on the nested `GeometryInstance` returned by `parent.GetSymbolGeometry()`).
- **CORRECTION — block NAMES are reachable** via `GetSymbolGeometryId()` → `SymbolGeometryId.SymbolId` → `doc.GetElement(symbol_id)`. The returned element is an `ElementType` of category `X_BG.dwg` whose **`Name` property is `X_BG.dwg.<BLOCK_NAME>`** (e.g. `X_BG.dwg.TUB-flat-2`). Always strip the `X_BG.dwg.` prefix.
  ```python
  sgid = nested_geom_instance.GetSymbolGeometryId()  # SymbolGeometryId, not ElementId
  block_type = doc.GetElement(sgid.SymbolId)         # ElementType with Name = "X_BG.dwg.<BLOCK_NAME>"
  block_name = _NAME.GetValue(block_type, None).replace("X_BG.dwg.", "")
  ```
  Use this name to match against profile rules (treadmill / elliptical / TV / etc.) — NOT just layer + symbol-id.
- `Arc.GetEndPoint(0)` throws on **unbound** arcs (full circles). Always check `arc.IsBound` first; for unbound, derive bbox from center+radius.
- Top-level `cad.get_Geometry(opts)` returns a wrapper `GeometryInstance` whose `Transform` already encodes world placement. When recursing children, do `parent_world.Multiply(child.Transform)` exactly once per nested instance — applying it twice silently puts coordinates hundreds of feet off.
- Subagents are blocked from writing files via Write/Bash/PowerShell on this project. To bypass, write via Revit MCP `System.IO.File.WriteAllText`. Main-session Write works fine.

## CAD block reading

- Filter by `GraphicsStyleId → GraphicsStyleCategory.Name` (CAD layer name).
- AIA layers in PF projects: `A-N-GYM EQUIPMENT` (new equipment), `A-X-GYM EQUIP` (existing), `A-N-RACEWAY`, `A-N-CHECK-IN`, `A-N-FIXT-CHECK-IN`, `A-N-PLUMB FIX`, `A-N-BEVERAGE COOLER`, `A-N-FAN`, `A-N-TELEVISION`, `A-N-SPA EQUIPMENT`, `A-X-SPA EQUIP`, `A-N-LOCKERS`, `A-N-LOCKER BENCH`, `A-N-FIXT-PHONE`, `A-N-SIGNAGE`.
- Filter inserts to building bounds (~`(40,-310,280,-185)` for Charlotte) — DWGs often contain reference drawings elsewhere.

## Powered vs unpowered gym equipment

- **A "powered" block has a power bar**; unpowered does not. Visually, an unpowered block is just a rectangle without the power-bar feature.
- Programmatic detection rule (`pf_gym_recepts.find_powered_gym_blocks`):
  - Look at small closed polylines inside the block (`SMALL_SHAPE_MIN < max_extent < SMALL_SHAPE_MAX` = 0.05 to 0.7 ft).
  - Powered if any candidate sits on the **lateral centerline** (`|local_y| < 0.5 ft`) at a **moderate forward offset** (`0.3 < local_x < 4 ft`).
  - Also powered if a symmetric pair straddles the centerline at the same local_x (two outlets on one bar).
- **Known false negatives** on the Charlotte job: the lower-left half of the Y=-254 cardio row (X<147). The detector missed 13 blocks here even though the user confirmed they're powered. **For dense cardio rows (≥15 blocks per row), it is safer to place at every block** rather than rely on per-block detection. Open question: figure out what's structurally different about those 13 blocks (geometry fingerprint? layer property?).
- **Known false positives**: my first pass detected 12 "powered" blocks scattered through the strength area that were actually random non-powered equipment. Detector was too greedy on misc geometry shapes. Trust the user's curated zones over per-block heuristics for non-cardio areas.

## Receptacle placement rules (Charlotte project, may generalize)

- **Family/Type**: `EF-U_Receptacle_CED :: Duplex Wall` (`OneLevelBased`).
- **Level**: `L1 - Finished Floor`.
- **Mounting height**: 46″ for cardio receps. (Train more samples for other equipment.)
- **Position**: at the **north edge center** of the block's world bbox — NOT at the block's origin and NOT at the "power bar center square". The square detector is for *discrimination only* — once powered, place at the bbox's north edge center.
- **Rotation**: receptacle face direction in this project is **world +Y (north)** for cardio. For TVs: **world −Y (south)** because the wall is to the north and the TV/recep faces into the gym away from the wall. Rule: recep face points **away from the wall** it's mounted on.
- **Snap-to-wall**: for wall-mounted receps (TVs, GFCIs over counter, etc.), find the closest CAD line on any `*WALL*` layer in the relevant direction and snap recep XY to the wall's room-side face. Wall thickness in this CAD is ~0.5 ft, so e.g., south-face of north wall ≈ north-face minus 0.5 ft.
- **Default rotation 0** with `place_workplane_fixture` leaves the family pointing in its default direction. For this family, that default direction = **+Y world** (north). To face south, rotate by `math.pi`. To face east/west, ± `math.pi/2`.

## Cleanup discipline

- Duplicate detection: any two receps within 0.5 ft of each other in XY are duplicates — keep the earliest, delete the rest.
- When the user provides corrective feedback ("delete the random ones"), interpret as: drop all receps in the affected Y/X band except those they explicitly confirmed.

## Skill files in this folder

- `cad_blocks.py` — `read_cad_blocks(doc, view, wanted_layers, building_bounds)` returns world-coord block inserts.
- `place_fixture.py` — `find_family_symbol`, `list_fixture_types`, `place_workplane_fixture`, `place_face_hosted_fixture_on_wall`, `set_mounting_height_inches`.
- `place_tag.py` — `find_tag_type`, `list_tag_types`, `place_fixture_tag` (IndependentTag.Create with type override).
- `pf_gym_recepts.py` — `find_powered_gym_blocks(doc, view, layer, building_bounds)` returns powered blocks with their north-edge centers as suggested placement points.

## Additional rules discovered during cleanup pass

- **Duplicate-location guard**: my placement pipeline produced one true duplicate (id 8219806 stacked on id 8219752 at (95.29, -250.68)). Always run a dedup pass at the end of any batch — group receps within 0.5 ft XY and delete all but the earliest-created.
- **Functional Training Mobility Room (FTMR)** has gym-equipment blocks at X≈94-95 (west of the cardio row's leftmost block at X=97.6). The blocks here LOOK like powered cardio to the detector but are NOT. **Rule: any A-N-GYM EQUIPMENT block at X < cardio_row_left_edge belongs to FTMR and does not get a recep.** Use the cardio row's leftmost block X-coordinate (97.6 in this project) as the dividing line.
- **TV wall snap**: TVs are wall-mounted. After identifying the 5 candidate locations from `A-N-TELEVISION` blocks (or by detector at the north-edge of the gym area), find the nearest CAD wall line on any `*WALL*` layer (in this project, `A-N-WALL`, `A-WALL`, `A-X-WALL`, etc.) and snap the recep Y to the wall's room-side face. In Charlotte the wall face is at Y=-190.11 (wall is ~0.5 ft thick, so faces are at -190.11 and -189.61).
- **Default fixture orientation** of `EF-U_Receptacle_CED :: Duplex Wall` with `rotation_rad=0` is FACING NORTH (world +Y). Verified empirically: cardio receps placed with rotation_rad=0 face the right direction (north) without correction. To face south for TVs, rotate by `math.pi` (180°).

## Family/type map for PF fixture replication

Discovered in the CED library on this project:

| Equipment / Use case | Family | Type | Mount ht |
|---|---|---|---|
| General receptacle (cardio raceway) | `EF-U_Receptacle_CED` | `Duplex Wall` | 46" |
| GFCI counter recep | `EF-U_Receptacle_CED` | `Duplex Wall - GFCI` | counter |
| Counter outlet for AV / data | `EF-U_Receptacle_CED` | `Duplex Wall - Isolated Ground` | counter |
| USB counter | `EF-U_Receptacle_CED` | `Duplex Wall - USB` | counter |
| Standard TV recep | `EF-U_Receptacle_CED` | `Duplex Wall - TV` | 60" |
| 240V spa equipment (hybrid, red wave, hydromassage, sauna, cryolounge) | `EF-U_Receptacle_CED` | `Specialty Wall - 240V/1Ph` | 24" |
| Hand dryer junction box | `EF-U_Junction Box_CED` | `Wall - With Stem` | 42" |
| Drinking fountain JB | `EF-U_Junction Box_CED` | `Wall - With Stem` | varies |
| Spa disconnect (60A/3 NF) | `EF-U_Disconnect Switch_CED` | `Non-Fused - 60A` | 60" (next to recep) |
| Floor outlet (raceway drop) | `EF-U_Receptacle_CED` | `Duplex Floor` or `Quad Floor` | floor |
| Cord reel (overhead, cardio) | `EF-U_Cord Reel_CED` | `Cord Reel - JBox` / `Drop Cord - Duplex` | ceiling |
| EPO button | `EF-U_General Electrical Box_CED` | `Emergency Power Off Button` | per code |
| Motor switch | `EF-U_Motor Rated Switch_CED` | `Motor Rated Switch - 208V, 2 Pole` (etc.) | varies |

## Placement rules per fixture category (Charlotte project)

- **Cardio receps**: Duplex Wall at north-edge center of each powered A-N-GYM EQUIPMENT block, rotation_rad=0, mount=46".
- **TVs (top of strength area)**: Duplex Wall (or Duplex Wall - TV), snapped to the wall south-face just north of placement, rotation_rad=pi (face south, away from wall), mount=60".
- **TVs in locker rooms** (A-N-TELEVISION blocks at Y≈-284): Duplex Wall - TV at block origin, rotation_rad=block rotation (±45° from CAD), mount=60".
- **Spa equipment** (A-N-SPA EQUIPMENT blocks): Specialty Wall - 240V/1Ph at block origin, rotation_rad=block rotation, mount=24".
- **Hand dryers** (subset of A-N-PLUMB FIX blocks): EF-U_Junction Box_CED :: Wall - With Stem at block origin, mount=42". **Need geometry inspection or user input to discriminate hand dryers from toilets/sinks.**

## Spa equipment placement rule (USER-CONFIRMED)

For each piece of tanning / spa equipment (tanning bed, sauna control unit, hybrid bed, red wave, cryolounge bed, hydromassage, spray tan, BCS equipment):

**Position rule:** put the receptacle on the wall that minimizes `dist(wall, equipment_block_center) + dist(wall, nearest_door_in_room)`. In other words: the wall that's nearest BOTH (a) the tanning-area block in that room AND (b) the room's door.

**Rotation rule:** receptacle faces the wall's room-side perpendicular — i.e., points INTO the room (away from the wall).

**Implementation algorithm (also coded inline; should be factored into a skill):**
1. Gather CAD A-N-DOOR block insertion points within the tanning area.
2. Collect wall line segments from any `*WALL*` layer in the tanning area.
3. For each equipment block (or target room center if no block exists):
   - For each wall segment within ~12 ft of the equipment:
     - Compute `d_eq` = perpendicular distance from wall to equipment.
     - Compute `d_door` = min perpendicular distance from wall to nearest door of the target room.
     - Score = `d_eq + d_door`.
   - Pick the wall with min score.
   - Recep position = projection of the equipment XY onto that wall.
   - Recep facing direction = perpendicular to the wall, in the direction of room center (inward).
   - Rotation_rad = `math.atan2(inward.x, inward.y)` (since family default points +Y; `(sin r, cos r) = inward`).

**Known pitfall** discovered while implementing: two receptacles for adjacent rooms (e.g., 103E and 103H separated only by a shared wall) can end up at the **same wall point** because that wall is closest to both. Fix needed: when projecting onto a wall, verify the wall actually bounds the target room — e.g., require the wall midpoint to be within a small distance of the target room center along a coordinate axis, or that the inward-facing direction points TOWARD this room (not away). TBD: add room polygon detection so we know which side of each wall is which room.

**REFINED rule (USER-CORRECTED): MUST FILTER TO LONG WALLS FIRST.** Without filtering, the algorithm picks 1-2 ft wall segments (door jambs, partial wall pieces) because they happen to be close to both equipment and door. Pre-filter wall segments to those with length > 4 ft (LONG_MIN). In the Charlotte tanning area: 890 total wall segs → 325 long walls. This eliminates door frame pieces, partial wall fragments, etc.

**Implementation update:**
```python
LONG_MIN = 4.0  # ft minimum for a wall segment to be considered "long"
long_walls = [w for w in wall_segs if w['len'] > LONG_MIN]
# Then apply the (room_center, nearest_door) scoring on long_walls ONLY.
```

**Room half-extent constraint**: also constrain candidate walls to those within `room_half_extent` of room center (~7-8 ft for typical tanning rooms, 10-15 ft for SAUNA/CRYOLOUNGE/BLACK CARD SPA). This stops the algorithm from picking a long wall that's far from the room.

**SAME-SIDE rule (USER-CONFIRMED)**: when the equipment block(s) sit off-center within the room, the chosen wall must be on the **same side** of the room as the equipment — never opposite. Otherwise the recep ends up halfway across the room from the actual machine it serves.

Implementation:
```python
# Compute equipment direction from room center
eq_dir = (eq_avg.x - room_center.x, eq_avg.y - room_center.y)
# If equipment is essentially at room center (|eq_dir| < ~1.5 ft), the constraint doesn't apply.
if |eq_dir| < 1.5: skip this filter.
# For each candidate wall:
wall_dir = (wall_projection - room_center)
dot = wall_dir . eq_dir
if dot < 0:  # wall is on opposite side from equipment
    REJECT this wall, try the next one
```

This caught the HYDROMASSAGE failure: room center at (204.33, -294.29), equipment blocks averaged at (198, -282) → NW of room center; my initial pick was a wall ESE of room center → opposite quadrant → would have been rejected by this rule.

**Block-to-room mapping**: when assigning A-N-SPA EQUIPMENT / A-X-SPA EQUIP blocks to rooms, **don't** use nearest-room-center — many blocks sit in the corridor between two rooms and a user-supplied mapping is needed. In Charlotte:
- 246.8, -293.3 → 103E TANNING
- 250.2, -293.3 → 103H SAUNA
- 253.6, -293.3 → 103G CRYOLOUNGE
- 198, -278.5 / 198, -285.7 → 103A HYDROMASSAGE
- 198, -292.8 / 198, -300 → 103A HYDROMASSAGE (existing equip on A-X-SPA EQUIP)

The corridor blocks at Y≈-293.3 mapped to rooms accessible from the corridor (whether N or S).

**MIN-CLEARANCE rule (USER-CONFIRMED)**: the chosen wall must NOT be right up against the CAD block equipment. Without this, my algorithm picked the north wall of 103B RED WAVE (a 7-ft wall at Y=-299.83) which is only **0.13 ft north** of the equipment block at (214.6, -299.7). The wall was *parallel and adjacent* to the equipment — physically the recep would sit on top of the bed.

Implementation: for each candidate wall, compute the minimum distance from the wall segment to ANY equipment block in the room. Reject if `d < MIN_CLEARANCE` (1.0 ft works in practice — covers wall thickness and minimum installation clearance).

```python
MIN_WALL_TO_EQ_CLEARANCE = 1.0  # ft
for w in long_walls:
    if eq_pts:
        d_to_eq = min(dist_pt_seg(eqx, eqy, w.p0.x, w.p0.y, w.p1.x, w.p1.y) for eqx, eqy in eq_pts)
        if d_to_eq < MIN_WALL_TO_EQ_CLEARANCE:
            REJECT this wall
```

Combined acceptance criteria (must all pass):
1. Wall length > 4 ft (LONG wall)
2. Wall within `room_half_extent` of room center
3. Wall NOT on opposite side from equipment (same-side rule, dot >= 0 with equipment direction)
4. Wall NOT right against equipment (dist >= MIN_WALL_TO_EQ_CLEARANCE)

Score remaining candidates by `d(wall, room_center) + d(wall, nearest_door)` and pick lowest.

**PER-AXIS BBOX rule (replaces earlier "no clamped projection" attempt)**: a wall qualifies as bounding a target room only if its projection point (after clamping to the wall segment) falls inside the room's bounding square. Check `|proj_x − room_center.x| ≤ half_ext` AND `|proj_y − room_center.y| ≤ half_ext`. This correctly rejects long walls that "run past" a small room (like SAUNA's 38-ft south wall running west past CRYOLOUNGE's west boundary) while still accepting partial walls whose endpoint lies inside the room (like CRYOLOUNGE's 6.3-ft west wall whose segment ends at the SW corner of the room — that endpoint IS inside CRYOLOUNGE, so the wall is valid).

**PREFER LONGER walls** in scoring: `score = d_room + d_door − LENGTH_BONUS × wall_length` with `LENGTH_BONUS ≈ 0.4`. This makes a 10-ft long wall slightly preferable to a 7-ft short wall for the same distance — pushed bottom-row tanning recepts onto the south building wall (217 ft) rather than internal walls.

**Full rule stack (all must pass) and final scoring:**

1. Wall length > 4 ft
2. `|proj_x − room.x| ≤ half_ext` AND `|proj_y − room.y| ≤ half_ext`  (per-axis bbox)
3. Wall not on opposite side from equipment (dot product ≥ 0; only if equipment > 1.5 ft from room center)
4. `min(dist(wall, eq_pt))` ≥ MIN_CLR (= 1.0 ft) — wall not flush against equipment block
5. **Score** = `d(wall, room_center) + d(wall, nearest_door) − 0.4 × wall_length`
6. Pick wall with minimum score.

This algorithm produced acceptable placements for all 11 spa rooms in Charlotte including the tricky 103G CRYOLOUNGE (whose actual west wall is a 6.3-ft segment, not the misleading 38-ft wall a block away).

**Bottom-row tanning rooms (103B–103F) specifically**: each is bounded south by the building exterior wall at `Y ≈ -303.71`. The recep on the south wall should sit at `(room_center_X, -303.71)` facing north (rotation_rad = 0). North wall of the same room divides from the corridor.

**Layers to scan for tanning equipment**: `A-N-SPA EQUIPMENT` and `A-X-SPA EQUIP` are reliable for HYBRID/RED WAVE/HYDROMASSAGE/HYBRID equipment. **However**, tanning beds in rooms 103C / 103E / 103K / 103L / SAUNA / CRYOLOUNGE / BLACK CARD SPA are drawn as **plain CAD line/arc geometry, not block inserts** in this project. The block-detector approach can't find them. **For these rooms, use the room name/location as anchor and find equipment via raw geometry detection** (e.g., find a closed polyline + curves that looks bed-shaped within the room boundary).

## Misc cleanup learnings from this iteration

- **PLUMB FIX block discrimination**: A block at `(106.9, -294.2)` on layer `A-N-PLUMB FIX` was a shower head/drain, not a hand dryer. Placing a J-box there is wrong. **Rule: do not place J-boxes on PLUMB FIX blocks that fall inside SHOWER rooms** (room name pattern `*SHOWER*`). Only place J-boxes on PLUMB FIX blocks in dry areas (corridor near drinking fountain, restroom areas adjacent to but not inside showers). Better long-term: probe the block's geometry — hand dryers are small rectangles (~1 ft x 0.5 ft); toilets/sinks are larger.
- **Project layer status**: in this Charlotte takeover, layers `A-N-BEVERAGE COOLER`, `A-N-CHECK-IN`, `A-N-FIXT-CHECK-IN`, `A-N-FAN`, `A-N-FIXT-PHONE`, `A-N-SIGNAGE`, `A-N-FIXTURE-COMPUTERS`, `A-N-FLOOR FIX`, `A-N-FURN` have **zero blocks inside the building bounds**. Likely meaning: existing equipment is being reused (takeover), or it's drawn on other layers. For new construction projects expect these layers to be populated.

## CORRECTED rotation formula (USER-CONFIRMED bug)

My original `facing_rad = math.atan2(inward.x, inward.y)` gave wrong results for east/west walls — receps pointed OUTWARD (into the wall) instead of INTO the room. The family's default orientation in this CED library faces **world +Y (north)**.

**Correct formula** to rotate a default-north family to face arbitrary `inward = (dx, dy)`:
```python
facing_rad = math.atan2(inward[1], inward[0]) - math.pi/2
```

Verification:
- inward (0, 1) [north]: atan2(1, 0) − π/2 = π/2 − π/2 = 0 ✓ (face north)
- inward (0, −1) [south]: atan2(−1, 0) − π/2 = −π/2 − π/2 = −π ≡ π ✓ (face south)
- inward (−1, 0) [west]: atan2(0, −1) − π/2 = π − π/2 = π/2 ✓ (face west)
- inward (1, 0) [east]: atan2(0, 1) − π/2 = 0 − π/2 = −π/2 ✓ (face east)

## INWARD-OFFSET rule (USER-CONFIRMED)

After projecting equipment onto the wall, offset the recep position by `INWARD_OFFSET ≈ 0.8 ft` toward the room interior (along the `inward` direction). Smaller offsets (0.3 ft) put the recep visually INSIDE the wall thickness, where it appears as if it's in the adjacent room. 0.8 ft puts it clearly on the room side of the wall face.

## Direction-specific wall picker (USER-CONFIRMED hard constraint)

Each room has a **direction preference** that hard-constrains which wall to choose. From Charlotte project user feedback:

| Room | Direction (wall side relative to room center) | Why |
|---|---|---|
| Bottom-row tanning (103B/C/D/F) | south | South building wall (217 ft); equipment is centered, door at north corridor |
| 103E TANNING | north | Equipment block is in NE corridor; long N wall ~12 ft |
| 103A HYDROMASSAGE (4 beds) | west | Beds line up along west wall; recep behind each bed |
| 103G CRYOLOUNGE | west | Equipment is to the SW of room center; west wall faces equipment + door |
| 103H SAUNA | east | (per user) |
| 103J SPRAY TANNING | west | (per user) |
| 103K TANNING | east | Door is at NE corner → east wall |
| 103L TANNING | west | Door is at NW corner → west wall |

Implementation: separate `half_ext_axial` (how far in the direction we search, e.g., 6–12 ft) and `half_ext_perpendicular` (lateral tolerance, e.g., 6 ft) so a long shared wall is allowed slightly off-axis while still keeping recepts inside the target room.

## Multi-fixture-per-profile rule (USER-CONFIRMED)

Yaml profiles list multiple Electrical Fixtures per piece of equipment. Each must be placed:

- **Tanning beds** (Upright Tanning, Tanning Bed 42.4, HM62B, HM62A, HM62E, PF Tanning 42-3, etc.): Non-Fused 60A disconnect (240V dedicated, panel L4) + Duplex Wall recep (120V, panel L3 for controls)
- **Hybrid Tanning Booth**: Disconnect 60A + Duplex Wall (same as tanning)
- **Hydromassage / Hydro Lounge**: Specialty Wall 240V/1Ph
- **Cryolounge**: Duplex Wall
- **Sauna Redzone**: Duplex Wall
- **Spray Tanning**: Duplex Wall - GFCI
- **TV (43, 27, 55, 60, 65, 70, 75)**: Duplex Wall - GFCI (or plain Duplex Wall for some) + Data outlet
- **Sloan Trough Sink** (locker rooms): 5× Duplex Wall - GFCI (vanity outlets) + 2× Junction Box Wall (hand dryers)
- **PF_Plan Hand Dryer**: Junction Box - Wall, 42″ AFF
- **PF_Plan Computer / IT Racks**: Quad Wall

When multiple fixtures share a wall, space them ~2 ft apart along the wall direction.

The yaml also includes Data Devices (JB, Data outlet, Speaker) and Generic Annotations (keynotes) — these are NOT in the Electrical Fixtures category and are deferred to a separate placement pass.

## Wall ORIENTATION constraint (USER-CONFIRMED bug)

When the user says "east wall" of a room, they mean a **vertical wall** (running N-S) located east of the room center — NOT just any wall whose projection is east of room center. Without this, my algorithm picked the horizontal south wall (whose left endpoint happened to be east of room center) and dropped the SAUNA recep into the door opening at (248.65, -290.54).

Implementation:
```python
for w in long_walls:
    wdx, wdy = w.p1.x - w.p0.x, w.p1.y - w.p0.y
    if abs(wdx)/w.len > 0.95: w.orient = 'horizontal'   # runs along X
    elif abs(wdy)/w.len > 0.95: w.orient = 'vertical'   # runs along Y
    else: w.orient = 'other'

# When picking by direction:
direction → required orientation:
  'east'/'west'   → 'vertical'
  'north'/'south' → 'horizontal'
```

## Flush offset (USER-CONFIRMED)

INWARD_OFFSET dropped from 0.8 ft to **0.1 ft** — fixtures sit flush against the room-side face of the wall. 0.8 ft pulls them noticeably into the room space ("jut out"), which the user rejected. With 0.1 ft offset the recep is visually on the wall.

## Doorway avoidance (USER-CONFIRMED issue, partial fix)

After picking a wall, project all candidate doors onto that wall. If a door projection is within ~2.5 ft (along-wall direction) of the recep position, slide the recep along the wall AWAY from the door until min_dist met. Otherwise the recep ends up IN the door opening (which is visually a gap in the wall).

```python
def shift_away_from_door(target, door_pts, min_dist=2.5):
    for door in door_pts:
        d, dpx, dpy = dist_pt_seg(door.x, door.y, wall.p0.x, wall.p0.y, wall.p1.x, wall.p1.y)
        if d > 1.5: continue   # door not on this wall
        along = (dpx - rx)*wd.x + (dpy - ry)*wd.y
        if abs(along) < min_dist:
            shift = min_dist - along  (or − min_dist − along for negative side)
            rx += shift * wd.x;  ry += shift * wd.y
    return (rx, ry)
```

## Multi-recep-per-room (USER-CONFIRMED)

Some rooms have multiple equipment items each needing its own recep:

- **HYDROMASSAGE**: 4 hydro beds along west wall — 4× Specialty 240V/1Ph, one per bed
- **CRYOLOUNGE**: 3 cryolounge beds along west wall — 3× Duplex Wall, evenly spaced at Y ≈ −283, −289, −295
- **BLACK CARD SPA**: multiple chairs/equipment items along the south wall (between BCS and the top tanning row) — at least 4 recepts at X ≈ 218, 228, 238, 248

For per-equipment placement, supply `multi_y` (or `multi_x`) array and re-run `find_wall` with each bed's coordinate as the room reference point so the algorithm picks the right wall segment at that specific Y/X.

## Match rooms by TYPE NAME, not number (USER-CONFIRMED hard rule)

Room number + letter combinations vary between PF projects (e.g., 103L TANNING in Charlotte Central might be 103A TANNING in another store). The placement logic should match on the **GENERAL ROOM TYPE NAME** (TANNING, SAUNA, SPRAY TANNING, CRYOLOUNGE, HYDROMASSAGE, RECEPTION, CHECK-IN, etc.) and on the **CAD blocks inside the room**, not on the specific number.

**Implementation**: when iterating rooms, key on the upper-cased ROOM_NAME parameter. Filter sub-types ("TANNING" excludes "SPRAY TANNING" if needed). Then place per the type's profile.

Example: `if "TANNING" in name and "SPRAY" not in name:` → use tanning bed profile. `if "SPRAY" in name:` → spray tanning profile.

## Placement heuristics — consolidated understanding (USER-CONFIRMED through 9 rooms)

A unified rationale for WHY each fixture goes where it does. Apply these in order:

**1. Equipment-tied receps go AT the equipment, not on the room perimeter.**
- Each CAD block (Treadmill, Stepmill, HM5A, SmarteCarte Massage Chair, etc.) defines an equipment position.
- The recep is placed at/near the block, with offset depending on block geometry:
  - Cardio: at block insertion point (= bbox north-edge center for these blocks).
  - Massage chair: at the chair's bbox CENTER (insertion point ≠ center for SmarteCarte; offset +2.14 X, +21.37 Y from insertion).
  - Bay-divided rooms (cryolounge): at the bay's RELATIVE CORNER (near the equipment's plug end), not centered.
  - TV truss: along the truss line, paired ±28" X per anchor.

**2. Convenience / counter receps go on ACTUAL CAD walls** (extracted from `A-N-WALL` / `A-WALL` / `A-X-WALL` / `A-N-WALL LOW` / `A-WALL-New` / `A-X-DEMISING WALL` / `A-X-LOW WALL` / `A96-WALL NEW-INT`), with `+0.1 ft` inward offset for flush placement.

**3. Fixture orientation = perpendicular to wall, facing INTO the room.**
- N wall (room south): rot = π (face south).
- S wall (room north): rot = 0 (face north).
- W wall (room east): rot = −π/2 (face east).
- E wall (room west): rot = +π/2 (face west).

**4. Counter / desk rooms wrap receps around 3 sides of the counter** (not all on one wall):
- Few on long shared room boundary wall (e.g., N wall of v/r/c-i).
- More on the front-of-counter (typically S in v/r/c-i).
- Some on the W or E partial wall (the short interior partition that creates the counter alcove).
- BCS suite standing-space corner: ~6 USB Duplex receps clustered with 1× #15 TYP keynote.

**5. Keynote placement rules**:
- **At least 2 ft from the recep** (USER-CONFIRMED) — far enough that the keynote and recep symbols don't overlap at typical view scale. 1 ft is too close.
- On the SAME SIDE the recep faces (so keynote sits IN the room, visible without overlapping the wall).
- For cardio raceway: 1 per row at midpoint (NOT per machine).
- For dense identical clusters: 1 TYP keynote + fan-out leaders to each (deferred since `HasLeader` not supported on this family).
- For vending machines: "above or below" based on recep orientation:
  - Recep faces N → keynote 2 ft N (above).
  - Recep faces S → keynote 2 ft S (below).

**6. Panel assignment by load type** (PF rule):
- **L1** (600A 208Y/120V) — cardio (Treadmill 1500 VA, Stepmill 1200 VA, HM5A 400 VA), all dedicated 1-pole 120V.
- **L2** (125A 208Y/120V) — TV TRUSS, strength side, vestibule heater.
- **L3** (125A 208Y/120V) — general 120V: TV receps, computers, convenience, vending (1000 VA), USB clusters, IT RACK Quad.
- **L4** (240V Δ) — tanning disconnects (60A 3-pole), hydromassage (Specialty Wall 240V/1Ph), hybrid bed, cryolounge duplex (1440 VA, exception — duplex on L4 not L3), cryolounge disconnect bays (4800 VA NF), beauty angel (1500 VA 3-pole), SmarteCarte Massage Chair Duplex Floor (500 VA), spray tan specialty (5040 VA).

**7. Asheville's keynote legend texts** (applied to `CED-G-NOTE TEXT` param):
- 7 = ISOLATED GROUND FOR CHECK-IN IT RECEP (PF_Plan Computer)
- 8 = PHONE / RJ45 JACK FOR FA/MGR; DATA WIRE TO TANNING (TMAX cluster)
- 9 = DATA RACK / AV VENDOR COORD (IT RACK Quad)
- 11 = TV MOUNTING HEIGHT/CONNECTION (all TVs)
- 13 = POWER & DATA RACEWAY + RECESSED J-BOX FOR FLOOR ROUTING (cardio raceway)
- 15 = MOUNT RECEP/USB PORTS HORIZONTALLY ABOVE COUNTER (USB clusters)
- 16 = BEVERAGE COOLER — DEDICATED RECEP PER MFR (vending)
- 19 = COORD WITH FAN MFR (FAN 1 disconnects — NOT 18 despite yaml; legend overrides)
- 20 = FALSE COLUMN RACEWAY (also used for massage chairs per user direction)
- 21 = POWER FOR ECH-1, SEE L1 SCHEDULE

**8. Bay-divided rooms** (cryolounge, locker stalls): place fixture near the bay's RELATIVE CORNER (corner closest to equipment's plug end), not centered. Charlotte's 4 cryo bays shifted south 2'-4" from bay center to put receps near the bed head.

**9. CAD block raw-geometry on layer X**: when a "block" is drawn as raw PolyLine geometry (not as a named block), search the CAD layer for that geometry and use the bbox center for placement. Example: A-N-BEVERAGE COOLER has 3 vending machine rectangles drawn as PolyLines (no named block) — use their bbox centers as placement anchors.

**10. Unit conversion** — ALWAYS use `UnitUtils.ConvertToInternalUnits(display, unit_type_id)` for Double params. Factor ≈10.76 for VA/V (m²/ft²); factor 1 for Amps.

## Block names ARE extractable (USER-CONFIRMED required for matching)

Prior assumption ("can't get block names, use only layer + symbol-id + bbox") is WRONG. Block names map directly to yaml profile entry names. Without using them, Voronoi room-assignment alone places equipment in the wrong rooms.

**Path**: nested `GeometryInstance` → `GetSymbolGeometryId().SymbolId` → `doc.GetElement(symbol_id)` → `Name`. Result is `"X_BG.dwg.<BLOCK_NAME>"`. Strip the prefix.

**Example in CAD background of this project (X_BG.dwg)**:
- `Treadmill (T5X-5PL-PF)` × 40
- `Stepmill (C5X-5PL-PF)` × 12
- `Elliptical (E5X-5PL-PF)` × 7
- `Ascent Full (A5X-4PL-PF)` × 5
- `Bike Rec. (R5X-5PL-PF)` × 4
- `Bike Up. (U5X-5PL-PF)` × 4
- `Rower (ROWER-5PL-PF)` × 4
- `HM5A` × 4 — Indoor Cycle
- `HM63C` × 4 — Hybrid Light Bed (tanning)
- `Flat Adj. Incline Bench (MG-A82-5PL-PF)` × 12 — strength, unpowered
- `Smith Machine (MG-PL62-5PL-PF)` × 6 — strength, unpowered
- `Dumbbell Rack (G3-FW91-5PL-PF)` × 6 — strength, unpowered
- `E1` × 17 — keynote E1 marker positions (raceway base)
- `E2` × 11, `E5` × 4 — other keynote anchors
- `F1`, `F3`, `F9`, `F13` — light fixtures (count up to 134 for F3)
- `SPEAKER` × 14
- `Room Name` × 38 — CAD room-name labels (better room anchors than placeholder Spaces)

Total distinct block names in the CAD: 167. Each maps either to a yaml profile (place fixtures) or to a passive block (no power needed).

**Room assignment**: do NOT rely on Voronoi to the placeholder Space `loc_xy` (8×6 ft placeholders). It misclassifies (e.g., 8 stepmills near the south end of the gym got pulled into "MENS LOCKER ROOM" by Voronoi because the locker label was nearest). Either use the 38 `Room Name` CAD blocks as room anchors, or build hand-tuned room bboxes from the CAD architecture, or — for an open gym area — assign by block NAME (cardio-class blocks belong to the gym regardless of which sub-zone label they sit nearest).

## No wall, no recep — verify before placing (USER-CONFIRMED hard rule)

When placing any convenience recep via perimeter walk or computed position, **verify the target XY is within ~1 ft of an actual CAD wall segment** before placing. If not, SKIP placement — do not place a recep "floating" in the room with no wall behind it.

**Implementation**:
```python
def nearest_wall_dist(px, py, segs):
    best = 1e9
    for s in segs:
        d = dist_pt_seg(px, py, *s)
        if d < best: best = d
    return best

# Before placing:
if nearest_wall_dist(target_x, target_y, wall_segs) > 1.0:
    skip  # do not place
```

Charlotte 107 FTMR perimeter walk placed 7 receps; 3 had to be deleted because the assumed N and S walls (Y=-237 and Y=-285) only existed in part of the X range — the W and E walls + partial N/S were real, but the remainder of the N/S "walls" were open boundaries (doors / no wall). Final count: 4 receps placed where walls actually exist.

## Convenience-recep / room placement synthesis (USER-CONFIRMED through 25+ rooms)

For BREAKROOM, STORAGE ROOM, LOCKER, JANITOR, MECHANICAL, etc. — convenience receptacles follow a consistent pattern:

### Per-wall placement rule

For square/rectangular rooms, place **one convenience recep at the CENTER of each wall**, facing into the room from each wall. So a 4-wall room gets 4 receps (N, S, E, W wall mids).

For long narrow rooms, more receps may be placed along the long walls.

### Wall identification — important!

The actual N/S/E/W walls of a room may NOT be where the Space's loc_xy suggests. Use the CAD walls from `*WALL*` layers, and **search a WIDE Y/X range** around the room loc. Some breakroom/storage rooms in Charlotte have:
- 101C BREAKROOM: actual extent Y=-267 (building N wall) to Y=-289 (interior partition) — 22 ft tall, NOT 7 ft as initial bbox suggested
- 101D STORAGE: Y=-289 (shared with breakroom S) to Y=-303.4 (building S wall) — 14 ft tall

The H walls I initially found at Y=-275 and Y=-282 were *interior partitions / casework*, not the room's bounding walls.

### Symbol-edge vs anchor-point

Revit's Duplex Wall recep symbol has a height of ~6" in plan view. To make the recep visually flush against a wall:
- Anchor point (center) must be 0.25–0.5 ft from the wall face (half the symbol height), not just 0.1 ft.
- For N wall (face at Y_w), anchor center = Y_w - 0.35 (south of wall face by symbol half-height).
- For wall-mounted disconnects (which have larger symbols ~1 ft tall), use 0.5–0.7 ft inward offset; even more if needed for handle clearance.

### Sloan trough sinks (locker / restroom pattern)

Each trough has:
- **5× Duplex Wall - GFCI** vanity outlets at X-offsets [−50, −30, 0, +30, +50] inches from trough center
- **2× Junction Box - Wall With Stem** hand dryers at X-offsets [±72] inches from trough center, **PLUS an additional 2'3" outward** (where the hand dryer CAD imagery is drawn)
- **Plus**: JBs are also offset south 1'10" + east 1'3" (inner) / west 1'3" (outer) from the trough vanity Y to match hand-dryer imagery in CAD

The trough is centered under the STORAGE room (104D for women's, 105D for men's), on the south wall of locker / north wall of toilet — but shifted 5'4" outward from the locker rooms' dividing wall.

### Storage room flanking the locker-room dividing wall

For 104D STORAGE (W locker side, east of dividing wall) and 105D STORAGE (M locker side, west of dividing wall):
- **Outer recep**: on the N wall of storage, facing south into storage room.
- **Inner recep**: against the dividing wall (X=141 area in Charlotte), facing OUTWARD from the dividing wall (east for 104D, west for 105D), positioned at Y ~4-5 ft north of storage center (typically up by the locker room area).

## Spa room synthesis — TANNING / SPRAY TANNING / SAUNA / HYBRID / RED WAVE (USER-CONFIRMED through 9 rooms)

These rooms share a common pattern. Match by **GENERAL ROOM TYPE NAME** in the Space ROOM_NAME parameter (not room number).

### Fixture set per room type

| Type | Equipment | Disconnect | Receptacle | Keynote |
|---|---|---|---|---|
| TANNING | tanning bed | Non-Fused 60A, L4, 240V, 3-pole, 10392 VA, "TANNING BED - <room>" | Duplex Wall, L3, 180 VA, "RECEPT - BCS" | #5 "TANNING/HYBRID BED 240V/30A NEMA 6-30 + 60A/3 NF DISCONNECT" |
| HYBRID | tanning bed (HM62E etc.) | Non-Fused 60A, L4, 240V, 10392 VA, "HYBRID - <room>" | Duplex Wall, L3, 180 VA, "RECEPT - BCS" | #5 same text |
| RED WAVE | HM62E variant | Non-Fused 60A, L4, 240V, 3-pole, 30A rating, 7067 VA, "RED WAVE - <room>" | Duplex Wall, L3, 180 VA, "RECEPT - BCS", NONGROUPEDBLOCK | #5 same text |
| SAUNA | sauna redzone | Non-Fused 60A, L4, 240V, 2-pole, 5760 VA, "SAUNA REDZONE - <room>" | Duplex Wall, L3, 180 VA, "RECEPT - BCS" | **Plus** JB Wall-With-Stem on OPPOSITE wall mid (60″ mount) + #22 keynote 2 ft from JB |
| SPRAY TANNING | spray tan unit | Specialty Wall 240V/1Ph, L4, 240V, 2-pole, 5040 VA, "SPRAY TANNING - <room>" | Duplex Wall - GFCI, L3, 180 VA, "SPRAY TANNING GFI", NONGROUPEDBLOCK, GFI | #21 "POWER FOR ECH-1, SEE L1 SCHEDULE" 2 ft from specialty |

### Placement rules

1. **Disconnect** — mounted on the room's E or W side wall (NOT the door wall):
   - Side selection: pick the corner closest to the CAD block equipment.
   - X offset: ~0.15 ft from mounting wall (tucked right against, but inside the room).
   - Y offset: 6+ inches from perpendicular (door) wall — for upper-row tanning (door = N) use Y_disc = N_wall - 0.7; for bottom-row use Y_disc = N_wall - 0.7 to -1.0 depending on equipment overlap.
   - Rotation: 90° CW if on E wall (faces west into room); 90° CCW if on W wall (faces east into room).
   - Mount height: 60" AFF.
   - Must NOT overlap or block the CAD equipment block.
   - Must NOT be placed behind door swings.
2. **Receptacle (Duplex Wall or GFCI)** — on the room's DOOR WALL (typically N for upper row, N for bottom row), middle of wall, facing into room (perpendicular to door wall).
   - X offset: at the door-wall snapped position + 0.1 ft inward (flush).
   - Y position: just inside the door wall (~0.1 ft inward).
   - Mount: 18" AFF (general) or 60" (TV recep).
3. **Keynote** — 4 ft into the room from the disconnect, shifted 2 ft toward room center for visibility.
4. **Same wall as equipment** — if equipment is on east side of room, disconnect on east wall, etc.
5. **For SAUNA only**: additional JB on the OPPOSITE side wall (opposite from equipment) at middle, with #22 keynote 2 ft into the room from the JB.
6. **For SPRAY TANNING only**: Specialty on E wall S corner; GFCI on N door wall mid; #21 keynote 2 ft from specialty (into room).

### Charlotte Central reference positions (for verification)

Upper-row tanning row (Y around -285):
- 103L TANNING (W wall mounted disc at SW area, N door wall recep), 103K TANNING, 103J SPRAY TANNING, 103H SAUNA.

Bottom-row tanning row (Y around -297):
- 103B RED WAVE (NE corner), 103C TANNING (NW), 103D HYBRID (NE), 103E TANNING (NW), 103F HYBRID (NE).
- All discs at ~Y=-297.1 to -297.6 (6+ in clearance from N wall at Y=-296.4).
- Disc X positions ~0.15 ft from respective E or W partition walls.

## Disconnect mount + clearance rules (USER-CONFIRMED)

When placing wall-mounted disconnects:

1. **Mount on the side wall (E or W)**, not the door wall (N or S) — even when at a corner. The disconnect sits against the side wall with its body extending into the room.
2. **Rotation**: face into the room from the mounting wall:
   - E wall (NE corner): rotate 90° CW from default → faces west (into room).
   - W wall (NW corner): rotate 90° CCW from default → faces east (into room).
3. **Wall clearance**: at least **6 inches (0.5 ft) clearance** from any perpendicular wall (e.g., the door wall) — the disconnect's body extends ~6 inches in plan view, so the disc CENTER needs to sit ~1.2 ft from a perpendicular wall to give the handle 6+ inch operational clearance.
4. **Inward offset from mounting wall**: ~0.7 ft (the disconnect housing thickness).
5. **Keynote position**: 4 ft from disc, perpendicular to mounting wall (toward room center), and shifted 2 ft toward the room's X-center for visibility.

**Example (Charlotte 103B–F)**: Disconnects on E or W side walls (0.7 ft inward), Y=-297.60 (1.2 ft south of N wall at -296.4 = >6" handle clearance), facing into the room.

## Disconnect placement — corner nearest equipment + door (USER-CONFIRMED hard rule)

For tanning / sauna / hybrid / red wave rooms (any room with a wall-mounted disconnect serving a single piece of equipment): the disconnect goes on the wall that has the DOOR, at the corner closest to the equipment's CAD-block position.

**Procedure**:
1. Identify the door wall (typically the corridor side — for tanning rows, this is the north wall).
2. Determine which side of the room the equipment block is on relative to room center (east or west).
3. Place the disconnect at the corner formed by (door wall) ∩ (equipment-side wall).
4. The keynote (#5 or appropriate) goes 4 ft INTO the room from the disconnect, perpendicular to the door wall.

**Why**: this gives a person walking in the door (from the corridor) immediate visual access to the disconnect, and keeps the disconnect away from the equipment so it can be operated without obstruction.

**Counter-example (wrong)**: Charlotte bottom-row tanning rooms initially had disconnects in the SW corner (far from door); these were moved to NE/NW corners on the N door wall per equipment side.

## Place bay fixtures near the bay's relative corner, not centered (USER-CONFIRMED)

For rooms divided into bays (cryolounge bay partitions, locker stalls, tanning beds), the fixture for each bay should sit **near the corner formed by the back wall + partition wall**, not centered in the bay. The "relative corner" is the corner closest to the equipment's plug location (e.g., head of the bed).

**Example (103G CRYOLOUNGE)**: 4 east-wall fixtures originally centered in their bays (Y = -281.5, -288, -294, -300) were shifted **south by 2'-4"** to Y = -283.8, -290.3, -296.3, -302.3 — placing each near the south partition of its bay (relative corner).

**Rule for future bay-divided rooms**: 
1. Compute the bay's bbox from the partition walls.
2. Determine the equipment's plug-end (head of bed, etc.).
3. Place the fixture at the corner of (back wall ∩ plug-end-side partition wall), with the standard inward offset.

## Keynote with "TYP" + fan-out leaders for cluster of identical fixtures (USER-CONFIRMED)

When a cluster of identical fixtures share a single keynote (e.g., the BCS suite standing-space cluster), use ONE keynote symbol with a "TYP" annotation + **multiple leader lines fanning out to each fixture** — NOT one keynote per fixture.

**Example**: Check-in BCS suite corner standing space:
- 4–5 USB receptacles on west partial wall (mounting heights +30″, +39″, +46″)
- Each labeled with a "U" letter callout (USB type identifier tag)
- Receps share panel/circuit (e.g., L3/1)
- ONE #15 keynote placed centrally with "TYP" annotation + leader to each USB recep

**Implementation rules**:
1. Use `EF-U_Receptacle_CED :: Duplex Wall - USB` for USB receps in counter clusters (not plain Duplex Wall).
2. Place ONE keynote per cluster, position it centrally near the cluster (not adjacent to a single fixture).
3. Add a "TYP" sub-annotation (text note or part of family) below the keynote number.
4. Draw a thin leader from the keynote to EACH fixture in the cluster.
5. The keynote inherits the cluster's shared L3/1 (panel/circuit) when circuited.
6. Add a separate "U" letter tag near each USB recep (USB identifier).

## Reception / Check-in counter recep distribution (USER-CONFIRMED)

For RECEPTION + CHECK-IN areas, do NOT pile receps on the long north wall. The actual PF distribution wraps around the counter island:

- **Few** on the N wall (1-2 in reception, 1-2 in check-in)
- **More** on the S wall (front face of counter, ~4 in check-in plus 2 in reception)
- **Some** on the WEST PARTIAL WALL of check-in (the short X≈230 vertical segment, ~3 receps)
- BCS suite alcove receps cluster at +42″ counter height
- Receps face into the open area of the alcove, not into the wall

The check-in counter is effectively an island; receps surround it on 3 sides. Same logic for RECEPTION's counter cluster.

**Rule for any counter/desk area**: identify the counter alcove walls (N, S, and W or E partial wall) and distribute receps proportional to wall length, with bias toward the front-of-counter and counter-back walls (NOT the long shared room boundary wall).

## Convenience receps require ACTUAL wall geometry (USER-CONFIRMED hard rule)

Convenience receptacles (Duplex Wall, not tied to a CAD-block-driven fixture) must be placed on **actual wall lines from the CAD**, not on hand-defined bbox edges. A bbox edge is a guess at where the wall is, and a guess is unacceptable.

**Rule**: do not place a convenience recep unless you have a verified wall line segment (from the X_BG.dwg wall layers — A-N-WALL / A-WALL / A-X-WALL / A-N-WALL LOW / A-WALL-New / A-X-DEMISING WALL / A-X-LOW WALL / A96-WALL NEW-INT) within tight tolerance of the target position.

**Procedure**:
1. Extract wall line segments from CAD (filtered by `gs_to_layer[gsid.IntegerValue] in WALL_LAYERS`).
2. For each target room, build a closed polygon from the wall segs that bound the room.
3. Distribute N convenience receps along the polygon's perimeter at evenly spaced positions ON THE WALL.
4. Compute facing direction perpendicular to the wall, pointing toward room center (inward).
5. If steps 1–4 fail to find a wall match, do not place — surface the failure to the user instead.

Voronoi-binning by `Space.loc_xy` is unreliable because Charlotte Spaces are placeholder 8×6 ft boxes whose label point is not the room center. Use real wall polygons instead.

## Revit internal-units conversion (USER-CONFIRMED critical bug)

Setting numeric parameters with raw display values produces wrong results. Specifically `Apparent Load Input_CED` and `Voltage_CED` store values internally at ~10.764× their display value (the m²/ft² factor — Revit's internal unit is per-m² for power density-typed quantities). Empirically observed:

- `1500` set on `Apparent Load Input_CED` → displays as **"139 VA"** (wrong)
- Family-default `Voltage_CED` of internal `1291.67` displays as **"120 V"** (correct, read-only)

**Fix**: always convert display→internal via `UnitUtils.ConvertToInternalUnits(value, display_unit_id)` where `display_unit_id` comes from the parameter's spec → `doc.GetUnits().GetFormatOptions(spec).GetUnitTypeId()`. Apply to **every Double-storage parameter** unless you've verified factor=1 (Amps appear to be 1:1).

```python
from Autodesk.Revit.DB import UnitUtils
def to_internal(elem, pname, display_val):
    p = elem.LookupParameter(pname)
    spec = p.Definition.GetDataType()
    du = doc.GetUnits().GetFormatOptions(spec).GetUnitTypeId()
    return UnitUtils.ConvertToInternalUnits(float(display_val), du)
```

Observed conversion factors (Charlotte project, this Revit version):
- Apparent Power (VA): factor ≈ 10.764 (1500 VA → 16145.87 internal)
- Voltage (V): factor ≈ 10.764 (120 V → 1291.67 internal)
- Current (A): factor 1.0 (no conversion)
- Integer / String params: no conversion

**Read-only by family-connector definition (cannot override)**:
- `Voltage_CED` — comes from the family's electrical connector
- `Number of Poles_CED` — same
- `Load Classification_CED` — storage type is ElementId; the `LoadClassification` class is restricted in this Revit MCP API (not collectible via FilteredElementCollector). Skip on placement; rely on circuit assignment to drive it.

## Load-parameter data from yaml profiles (USER-CONFIRMED critical)

Each fixture MUST be created with proper load data so circuiting later is correct. The yaml profiles supply the canonical parameter set per equipment type. Extracted profile parameters used in Charlotte:

| Profile | Family :: Type | Panel | V | Poles | Load (VA) | Circuit | Class |
|---|---|---|---|---|---|---|---|
| Treadmill (T5X-5PL-PF) | Receptacle :: Duplex Wall | L1 | 120 | 1 | 1500 | DEDICATED | NC |
| BCS_Upright Tanning | Disconnect :: Non-Fused 60A | L4 | 240 | 3 | 10392 | DEDICATED | NC |
| BCS_Upright Tanning | Receptacle :: Duplex Wall | L3 | 120 | 1 | 180 | DEDICATED | C |
| BCS_Cryolounge | Receptacle :: Duplex Wall | **L4** | 120 | 1 | 1440 | DEDICATED (GFI) | C |
| BCS_Hydro Lounge | Specialty Wall 240V/1Ph | L4 | 240 | 2 | 4320 | DEDICATED (GFI) | NC |
| BCS_Hydro Lounge | Duplex Wall - GFCI | L4 | 120 | 1 | 180 | DEDICATED | C |
| SAUNA REDZONE | Disconnect :: Non-Fused 60A | L4 | 240 | 2 | 5760 | DEDICATED | NC |
| SAUNA REDZONE | Receptacle :: Duplex Wall | L3 | 120 | 1 | 180 | '' | REC |
| Spray Tanning | Receptacle :: Duplex Wall - GFCI | L3 | 120 | 1 | 180 | NONGROUPEDBLOCK | REC |
| Spray Tanning | Specialty Wall 240V/1Ph (L6-30R) | L4 | 240 | 2 | 5040 | DEDICATED | NC |
| 43 tv | Receptacle :: Duplex Wall - GFCI | L3 | 120 | 1 | 300 | DEDICATED | C |
| IT RACKS | Receptacle :: Quad Wall | L3 | 120 | 1 | 600 | DEDICATED | C |

**Per-instance load-name override**: each placed fixture should have its `CKT_Load Name_CEDT` overridden to include the **room number** (e.g., `TANNING BED - 103L`, `HYDROMASSAGE - 103A bed3`). This lets the panel schedule readers identify each load uniquely.

**Implementation: `apply_params(elem, profile_params, load_name_override)`** sets the parameters on a placed `FamilyInstance`. Skip read-only parameters; respect `StorageType` (String/Integer/Double). Per Charlotte run this set **573 parameter values across 92 fixtures** in one transaction.

**Critical observations:**
- Cryolounge recep is on **L4** (the 240V Δ panel) per yaml — not L3 like other lounges. The duplex is at 1440 VA (not the standard 180 VA) — accounts for the cryotherapy machine.
- HydroMassage 240V Specialty has **4320 VA** at 30A — much higher than a typical duplex (180 VA). Distinct circuit per bed.
- SAUNA Disconnect is **5760 VA / 2-pole** — half the load of an upright tanner.
- Spray Tan Specialty is **5040 VA / 30A NEMA L6-30R**.
- 240V Specialty receptacles cluster on **L4** (the only 240V Δ panel in PF).

**Workflow: combine yaml profile data with previous-project (training E101) data** —
- Yaml gives canonical params per equipment profile (Apparent Load, Voltage, Panel, Circuit type).
- Training-project `hosts.json` shows realised positions, panel-circuit assignments, and which families ended up actually used.
- Together: choose family/type from yaml profile, set params from yaml, and use training-project patterns to inform placement decisions in the new project.

## Iteration audit: comparison of replication-attempt vs reference screenshot

After placing 119 fixtures across all stages, comparison vs the reference (saved final-state) screenshot:

### Errors (clearly wrong, not just shifted-on-correct-wall)

1. **103E TANNING disconnect+duplex landed on NORTH wall** — reference shows it on the SOUTH wall like the other bottom-row tanning rooms. My algorithm used the corridor SPA-EQUIPMENT block at (246.8, -293.3) as 103E's equipment reference, which flipped the same-side rule. **Fix**: when the room is in the bottom row (Y≈-300) and the building's south exterior wall is right there, prefer "south" direction over equipment-based direction, OR don't use corridor blocks as equipment anchors for rooms on the opposite side of a corridor.

2. **Restroom hand dryers in 104A / 105A TOILET — not placed** — `find_wall` returned None for direction='north'. Toilet rooms are small (~5×5 ft) and bounded by partition walls that may not pass the LONG_MIN (4 ft) filter. **Fix**: drop LONG_MIN to 2.5 ft for restroom hand dryers (a typical wall between toilet stalls is shorter than the building exterior walls).

3. **BCS Beauty Angel disconnects (north wall) — not placed** — `find_wall` returned None for direction='north' on BCS. BCS is bounded north by the building/tanning-corridor wall but my half_ext_axial=8 may have rejected it. **Fix**: increase ha to 12-15 for BCS (large room) and accept walls regardless of orientation if they're the only candidate.

4. **Reception counter on east wall (256.29, -250.48) — not placed** — `find_wall` returned None. Reception is bounded east by the VESTIBULE wall — that wall may be on a non-standard layer or shorter than 4 ft. **Fix**: lower LONG_MIN for counter recepts, or use south wall as fallback.

5. **Vestibule fixtures — entirely missing** — I never had a placement plan for the VESTIBULE 100 room. Reference shows recepts there (likely ECH-1 heater per training E101 pattern, possibly entry receptacles).

6. **Drinking fountain (PF_Plan DF) — missing** — yaml has a profile for this with Duplex Wall - GFCI; usually placed between locker rooms. The corridor PLUMB FIX block at (198.1, -272.1) is likely the drinking fountain.

7. **Storage / janitor / mechanical-room J-boxes — missing** — STORAGE ROOM 101D, JANITOR rooms, MECHANICAL room (per yaml `Room Name 101A ROOM MECHANICAL`) all need fixtures.

8. **Employee Breakroom — missing additional recepts** — only got 2 vending machine GFCI; reference shows more recepts there (counter, microwave, coffee maker).

9. **Floor outlets in cardio raceway** — reference shows raceway dots between cardio rows (per training E101 the floor outlets serve the cardio rows). Yaml has Duplex Floor / Quad Floor families. I placed the receps at north-edge of each block but didn't add the false-column raceway floor outlets.

### Gaps (planned for follow-up passes, not errors per se)

- Wire homeruns (THWN arcs from L1/L2 panels to cardio rows)
- Wire tags ("L1/X,Y,Z" slash-separated)
- Fixture tags (Panel/Circuit, Identity, Elevation/Comments)
- Keynote symbols (`Manual Key Note - All Shapes :: Square` with `CED-G-NOTE #`)
- Data devices (Junction Box - Wall, Data - Wall, Speaker)
- Lighting fixtures, exit signs (different system, deferred)
- Circuiting (`Power → Circuits` to actually assign each fixture to a panel circuit)

### Lessons to apply on next iteration

1. **Per-room wall direction should accept fallbacks** if the primary direction fails — try N→E→S→W in order, with looser constraints (smaller wall_length threshold) on later attempts.
2. **Don't use corridor blocks as equipment reference for rooms across the corridor** — match blocks to rooms by proximity only when the block falls inside the room's bbox.
3. **Restroom and other small rooms need shorter LONG_MIN** (2.5 ft instead of 4 ft) — bathroom partition walls are typically shorter than gym/spa walls.
4. **VESTIBULE, EMPLOYEE BREAKROOM, MECHANICAL, STORAGE rooms** need explicit profile mapping — they have specific yaml profiles I haven't been triggering.
5. **Cardio "raceway" floor outlets** are a separate element class — add them along with the wall receps (one floor outlet between every 2-3 cardio recepts per training E101 pattern).
6. **CRYOLOUNGE wall preference is ambiguous** — verify with user what side the cryo beds back up to. Currently chose east; possibly west or south depending on bed orientation.

## Iteration 2 results (USER-INITIATED replay after first audit)

**131 fixtures placed** (vs 119 in iteration 1) — gains from fixes:
- ✅ 103E TANNING moved to south wall (matching other bottom-row rooms)
- ✅ BCS Beauty Angel: 4 disconnects placed via `find_wall_fallback` (north→south→west→east)
- ✅ Reception + Check-in counter recepts placed via fallback (3 each: GFCI/IG/USB)
- ✅ Vestibule ECH-1 heater placed
- ✅ Drinking fountain at PLUMB FIX block (198.1, -272.1)
- ✅ Storage/Janitor/Mechanical J-boxes added
- ✅ Breakroom 3 counter GFCI
- ✅ Vending machines moved to VESTIBULE (member-facing area)

**Remaining errors after iteration 2:**

1. **Restroom hand dryers landed on BUILDING SOUTH WALL (Y=-303.71)** instead of toilet partition wall. Cause: my `find_wall_fallback` with `min_wall_len=2.5` allowed short partition walls, but the score `(d_room + d_door − 0.4 × length)` strongly preferred the 217.8-ft south wall (length-bonus dominated). The fall-back found a valid wall but too long.

   **Fix for next iter**: for restrooms specifically, cap `half_ext_axial` to room half-size (~3 ft) so the long building wall is out of range. Equivalently: clip the score's length bonus when wall_length > 3× room dimension.

2. **Sloan trough hand-dryer JBs were classified as "storage_jb"** in my parameter classifier (because the Y range overlapped). The locker-room JBs have 950 VA / hand-dryer params but got 180 VA storage params. **Fix**: tighten classifier — locker JB X/Y range needs to differentiate locker (Y -283..-270) from storage (Y -302..-292).

3. **Some BCS placeholder receps** got mis-classified or missed in classification. 11 of 131 fixtures didn't match any classifier (so didn't get params).

### Open insights for next iteration

- **For small rooms** (restrooms, janitor closets), the long-wall length-bonus dominates and pulls the algorithm to the building's exterior wall. Either: (a) skip the length-bonus when `room.half_ext < 5 ft`, or (b) constrain wall length to `wall_length ≤ 2 × room.diagonal`.
- **Restroom hand dryers** should anchor to specific A-N-PLUMB FIX block positions (which are at the partition wall) rather than the room name. The block I have at (151.7, -285.7) IS the hand-dryer block per the user's earlier feedback — I should use it as anchor.
- **Classifier ranges need refinement** — too many disjoint X/Y bbox tests; a single dict of `{room_name: (xrange, yrange)}` would be cleaner and less error-prone.
- **Vending machines per user**: my best guess location is VESTIBULE (member-facing). If wrong, the right location might be near reception or breakroom. Awaiting user verification.

## Sloan Trough Sink placement (USER-CONFIRMED)

The Sloan trough sink in locker rooms is **NOT centered on the locker room itself** — it's mounted in the corridor between the locker room and the toilet/shower area, **centered under the corresponding STORAGE room** (105D for men's, 104D for women's). The trough is mounted on a long horizontal wall around Y ≈ -285 (south wall of locker / north wall of toilet corridor).

Rule for future PF projects:
- For MEN's Sloan trough sink: center X = STORAGE room 105D center.X
- For WOMEN's Sloan trough sink: center X = STORAGE room 104D center.X
- Y is found by `find_wall(direction='south')` starting from the STORAGE room center, NOT from the locker room center
- Fixtures arrayed along wall:
  - 5× Duplex Wall - GFCI at offsets [-50, -30, 0, +30, +50] inches from center (vanity outlets)
  - 2× Junction Box - Wall at offsets [-72, +72] inches (hand dryer JBs at outer ends)

The Charlotte project STORAGE room positions:
- 105D STORAGE: (136.12, -283.25) — MEN's sink centerline
- 104D STORAGE: (146.29, -283.35) — WOMEN's sink centerline

The locker-room-center anchor is wrong:
- 105 MENS LOCKER center X=115.45 ❌ (off by ~21 ft from correct sink position)
- 104 WOMENS LOCKER center X=168.00 ❌ (off by ~22 ft)

## Open questions / TBD

- Better detector for the lower-left cardio missed blocks (currently brute-forced by "place at every block in dense row ≥15").
- `A-X-GYM EQUIP` (existing equipment) — 72 blocks; do these need receps too, or are they served by existing infrastructure?
- Tag types in this project are limited (only 3 fixture tag types: `Comments`, `Identity`, `Panel & Circuit Number`). No `Elevation (Inches)` — need to load that family from CED library if elevation tags are required.
- For TVs and other wall-mounted receps: should the recep be **on** the wall (Y = wall face) or **offset** by ~3″ (the family thickness)? Snapping to wall face may visually overlap the wall lineweight.
