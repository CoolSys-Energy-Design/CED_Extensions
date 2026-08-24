# Compute cardio placement targets: match TREADMILL/STAIRMASTER/POWERED BIKE
# staged profiles to A-N-RACEWAY power-bar squares near gym-equipment blocks.
import json, os, math
from collections import defaultdict

RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
d = json.load(open(os.path.join(RUN, "extract.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

# ---- collect A-N-RACEWAY ~0.25 ft squares (the outlet points) ----
squares = []
gym_blocks = []
for s in g.get("shapes", []):
    lay = (s.get("lay") or "").upper()
    if "bb" in s:
        cx, cy = s["c"]
        bx = s["bb"][2]-s["bb"][0]; by = s["bb"][3]-s["bb"][1]
        size = max(bx, by)
        if "RACEWAY" in lay and 0.1 <= size <= 0.6:
            squares.append((cx, cy))
        if "GYM EQUIP" in lay and 1.5 <= size <= 12:
            gym_blocks.append((cx, cy, size))

# dedupe squares within 0.4 ft
uniq = []
for (x, y) in squares:
    if not any(math.hypot(x-u[0], y-u[1]) < 0.4 for u in uniq):
        uniq.append((x, y))
squares = uniq
print("A-N-RACEWAY squares:", len(squares), " gym blocks:", len(gym_blocks))

# ---- staged cardio fixtures ----
cardio_loads = {"TREADMILL", "STAIRMASTER", "POWERED BIKE"}
fix = defaultdict(list)
for f in d.get("fixtures", []):
    ln = f.get("params", {}).get("CKT_Load Name_CEDT")
    if ln in cardio_loads and f.get("loc"):
        fix[ln].append({"id": f["id"], "loc": f["loc"][:2], "rot": f.get("rot") or 0})
for ln in fix:
    print("staged %-14s %d" % (ln, len(fix[ln])))

# report square Y-clusters (cardio rows) to sanity check
ys = sorted(set(round(y) for (x, y) in squares))
print("square Y values (rows):", ys[:40])
print("square count by y-row:")
rowc = defaultdict(int)
for (x, y) in squares:
    rowc[round(y)] += 1
for y in sorted(rowc, key=lambda k: -rowc[k])[:12]:
    print("   y=%d : %d squares  x-range [%.0f..%.0f]" % (
        y, rowc[y], min(x for x, yy in squares if round(yy)==y),
        max(x for x, yy in squares if round(yy)==y)))

json.dump({"squares": squares, "gym_blocks": gym_blocks,
           "staged": {k: v for k, v in fix.items()}},
          open(os.path.join(RUN, "cardio_ctx.json"), "w"))
