# HEB_MCP_AUTOMATION — multi-area power-design replication

## Structure (shared engine, per-area data)
```
HEB_MCP_AUTOMATION\
  skills\             shared, area-agnostic skill scripts (the engine)
  bakery_auto\data\   collected JSON for BAKERY
  pharmacy_auto\data\ collected JSON for PHARMACY
  README.md
```
Skills are parameterized by globals set before exec: `BA_DATA` = area data
dir, `BA_VIEW` = source/target view name (`_lib.py` reads these; defaults =
BAKERY). Same source project `RunUpdateProfilesHEBCarrollton95PercentCLOUDMODEL`,
same target `CED_HEB Template_MEPR_R24_OkmntProfCorr`, different view per area:
- BAKERY  : "Power Callout - BAKERY - L1"
- PHARMACY: "Power Callouts - PHARMACY"

Orchestrators (skills/):
- `collect_all.py`  — runs all 8 collectors for the configured area (read-only).
- `run_pipeline.py` — resets the area id map, runs the 7 placement steps.
- `run_all.py`      — MASTER: full pipeline for BOTH areas in turn.

Collected & verified: BAKERY 78 host/74 wire/23 kn/16 tn/42 ftag/60 wtag/152 ckt;
PHARMACY 73 host/57 wire/16 kn/12 tn/18 ftag/31 wtag/89 ckt.

---

Goal: learn the electrical power design from a source project and reproduce it
in a target project, generalizing via equipment-relative placement so it works
when equipment sits differently.

## Folders
- `skills/` — reusable scripts (run via Revit MCP `execute_revit_code`).
- `data/`   — JSON artifacts produced by the collectors (offline-reproducible).

## Skill scripts (run with SOURCE project active, both projects open)
| Script | Output | Notes |
|---|---|---|
| `_lib.py` | — | shared helpers (exec'd by others) |
| `collect_linked_elements.py` | `linked_elements.json` | equipment anchors in crop (2343) |
| `collect_host_elements.py` | `host_elements.json` | 78 devices/panels: params, circuits, equip-relative vectors |
| `collect_keynotes.py` | `keynotes.json` | 23 GA_Keynote Symbol_CED |
| `collect_wires.py` | `wires.json` | 74 wires: vertex polylines, circuit, connectivity |
| `collect_circuits.py` | `circuits.json` | 152 circuits: panel/ckt/rating/load/members |
| `replicate.py` | (mutates TARGET) | places devices/keynotes/wires; circuiting optional |

## Source→Target coordinate transform (VERIFIED 2026-05-19)
Both BAKERY views are NOT in shared coordinates. The transform is derived from
the common equipment link instance:

    T = T_tgtEquipLink * T_srcEquipLink.Inverse
    p_target = T.OfPoint(p_source)      ; v_target = T.OfVector(v_source)

- src Equip link origin (0,0,0); tgt link `00000_Carrollton_v24_Equip` origin (6.076, 475.964, 0)
- For this run T is a pure translation +(6.076, 475.964, 0); verified against
  559 uniquely-matched equipment families (dominant delta exact, outliers =
  equipment that genuinely moved → the generalization target).

## Run order (both projects open in one Revit; Revit MCP)
COLLECT (source active or readable):
1. `exec(open(.../collect_linked_elements.py).read())`
2. `exec(open(.../collect_host_elements.py).read())`
3. `exec(open(.../collect_keynotes.py).read())`
4. `exec(open(.../collect_wires.py).read())`
5. `exec(open(.../collect_circuits.py).read())`

REPLICATE (writes target cross-doc; set phase before exec):
- `BA_PHASE="A"; exec(open(.../replicate.py).read())`  -> 75 devices
- `BA_PHASE="B"; ...`  -> 23 keynotes
- `BA_PHASE="C"; ...`  -> 74 wires (creates THWN WireType if absent)
- `BA_PHASE="D"; ...`  -> circuits on existing branch panels (BA/RC), auto-numbered,
  with source circuit descriptions copied (Load Name, Schedule Circuit Notes,
  CKT_Load Name_CEDT, CKT_Schedule Notes_CEDT). Self-healing: reconciles against
  live model — recreates undone circuits, only refreshes desc on intact ones.

Idempotent: re-running a phase skips already-placed (tracked in
`data/replication_map.json`, and Comments stamped `[BA:src<id>]`).

## Equipment-relative replication (cross-building) — `place_relative.py`
HEB keeps one master equipment model; each store links a repositioned copy that
keeps identical Revit element IDs + IfcGUID. So equipment is matched ACROSS
projects by element Id (GUID-validated), not family/type or coords. Each device
anchors to the nearest SHARED source equipment, placed at
`tgt_anchor + Rz(dtheta)*offset`. Globals: `BA_DRYRUN`, `BA_TGT_TITLE`,
`BA_SRC_EQUIP`, `BA_TGT_EQUIP`, `BA_TGT_LEVEL`. Dry-run first (read-only report).
Proven on Carrollton -> Oakmont template: 75/75 devices on-equipment, 33
diverged >0.5ft from rigid. Devices tagged `[BA:src<id>][equiprel anchor=<id>]`.
place_relative.py now melds table-anchored + wall-snap corrections into Phase A
(anchor_dist>3ft -> table-anchored if on a table within 4ft, else wall/equip
face snap), all in one assimilated TransactionGroup.
Full equip/device-relative pipeline (each its own skill, dry-run->apply,
TransactionGroup undo, map in place_relative_map.json). ORDER MATTERS:
1. place_relative.py     - devices (equip-relative + table-anchored + wall-snap).
                           Also applies SAFE writable instance params captured
                           by collect_host_elements (`wparams`): load inputs,
                           symbol/visibility config. SKIPS identity/unique
                           (IfcGUID, Mark), worksharing (Workset, Export to IFC),
                           placement geometry (Offset from Host, Elevation from
                           Level, ACA*), Element_Linker (source IDs), Comments
                           (trace tag), and ALL circuit-derived params (CKT_*,
                           ampacity/voltage-drop/conduit-fill, conduit/wire
                           sizing) -> those are recomputed by the circuiting
                           step. collect_host_elements captures wparams with
                           value+storagetype+BuiltInParameter id (ElementId
                           params skipped: not portable across projects).
2. place_keynotes.py     - keynotes -> nearest placed device
3. place_textnotes.py    - text notes, horizontal, leaders device-relative
4. place_circuits.py     - ElectricalSystem per source circuit -> panel
                           (BA/RC/FS), auto numbers, copies Load Name/notes.
                           NOTE: commit regenerates; never doc.Regenerate()
                           after t.Commit() (illegal, no open txn).
5. place_wires.py        - wires device-relative, START bound to the circuited
                           device's electrical connector so wire joins the
                           circuit (Panel/Circuits populate) + solid
                           FilledRegion triangle homerun arrowhead at free end.
6. place_fixture_tags.py - IndependentTag -> placed device.
7. place_wire_tags.py    - IndependentTag -> placed wire. MUST match the source
                           tag TYPE (not just family): source uses
                           "Homerun - Panel & Circuits - Slash"; the "-Dash"
                           type renders blank. Tag only shows once the wire is
                           circuited (step 4+5).
Collectors: collect_textnotes/fixture_tags/wire_tags.py (+ leader/offset geo).
Proven end-to-end on Oakmont: 75 dev / 23 kn / 16 tn / 53 ckt / 74 wire /
42 ftag / 60 wtag, all displaying. Model NOT saved.

## Scaling to a NEW target project
The only project-specific knobs in `replicate.py`:
- `TGT_TITLE`, `SRC_EQUIP_LINK`, `TGT_EQUIP_LINK` (link doc title substrings),
  `TGT_LEVEL`, `OK_PANELS`.
The transform is auto-derived from the two equipment-link instances, so a
differently-laid-out bakery works as long as both projects link the same
equipment model. Equipment that genuinely moved is the generalization frontier:
use `rel_to_equipment` in `host_elements.json` (device offset vs nearest
equipment) to place relative to matched equipment instead of the rigid transform.

## Result (2026-05-19 first run, CED HEB Test Run_MEPR_R24)
75 devices / 23 keynotes / 74 wires all at 0.000 ft residual; 52 BA circuits
auto-numbered. Flags: panel FS absent (skipped); 1 circuit had no usable
connector. Model NOT saved (left for human review).

## Known gaps / decisions
- Target panels: BA, RC, LDP3, LDP4 exist; **panel `FS` does NOT** (32 source FS circuits).
- All 7 required families already loaded in target.
- Circuiting via API is the riskiest phase; recommended last, after visual placement check.
