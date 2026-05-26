## Per-room placement folder

Each room folder contains everything needed to re-execute placement for that one room.

**Files**:
- `place.py` — IronPython snippet, callable from Revit MCP `execute_revit_code`. Wraps the placement logic in one transaction.
- `manifest.json` — last-run placement output (element ids, positions, params actually set).
- `block_map.json` — what CAD block names in this room get which family/type + which yaml profile.
- `notes.md` — anything room-specific (rotation rules, wall constraints, deviations from yaml).

**To re-execute** a room's placement (e.g., after a model rollback):
1. Delete any prior placements in the room (via `manifest.json` ids).
2. Run `place.py` via Revit MCP `execute_revit_code`.
3. The script reads `block_map.json` + the global `charlotte_central_blocks_v3.json` + the yaml profiles, and idempotently places everything.

**Conventions**:
- Position: block insertion point unless a per-equipment rule overrides (e.g., TV TRUSS uses Y=-222.167 truss line).
- Rotation: 0 = family default (faces +Y north). Use π for south-facing.
- Mount height: per yaml profile or per LEARNINGS rules (cardio=46″, TV=60″, hand dryer=42″, etc.).
- Unit conversion: ALWAYS use `UnitUtils.ConvertToInternalUnits(display, doc.GetUnits().GetFormatOptions(spec).GetUnitTypeId())` for Double parameters (factor ≈10.764 for VA/V).
- Read-only params (Voltage_CED, Number of Poles_CED) come from the family connector — do not try to set.
