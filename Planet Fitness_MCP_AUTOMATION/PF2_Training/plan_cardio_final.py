# Build cardio placement plan: assign 56 staged fixtures to 56 squares with
# type zones from staging (SM=left bank one row, BIKE=right bank left edge,
# TM=rest), rotation = face the machine side (data-driven from fingerprint).
import json, os, math
from collections import defaultdict

RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
ctx = json.load(open(os.path.join(RUN, "cardio_ctx.json")))
fps = json.load(open(os.path.join(RUN, "cardio_fps.json")))
squares = ctx["squares"]
staged = ctx["staged"]

# side per square (machine side): from fingerprint
side_of = {}
for fp in fps:
    side_of[tuple(fp["sq"])] = fp.get("side", "S")

# organize squares into rows/banks
rows = defaultdict(list)
for (x, y) in squares:
    rows[round(y)].append((x, y))
row_ys = sorted(rows)   # [-253, -241]
south_y, north_y = row_ys[0], row_ys[1]
for ry in rows:
    rows[ry].sort()

def bank(x):
    return "L" if x < 156 else "R"   # gap between x=151 and x=161

# assignment sets
assign = {}   # (x,y) square -> load
# STAIRMASTER: south row, left bank, first 12 columns
south_L = [s for s in rows[south_y] if bank(s[0]) == "L"]
for s in south_L[:12]:
    assign[s] = "STAIRMASTER"
# BIKE: north row, right bank, first 4 columns
north_R = [s for s in rows[north_y] if bank(s[0]) == "R"]
for s in north_R[:4]:
    assign[s] = "POWERED BIKE"
# TREADMILL: everything else
for (x, y) in squares:
    if (x, y) not in assign:
        assign[(x, y)] = "TREADMILL"

# sanity: counts must match staged
cnt = defaultdict(int)
for v in assign.values():
    cnt[v] += 1
print("assigned counts:", dict(cnt))
print("staged counts:", {k: len(v) for k, v in staged.items()})

# build fixture_id -> target. Greedy: for each load, sort staged by x, sort its
# target squares by (bank,row,x), pair in order.
plan = []
for load in ("TREADMILL", "STAIRMASTER", "POWERED BIKE"):
    tgt = sorted([sq for sq, l in assign.items() if l == load],
                 key=lambda s: (bank(s[0]), round(s[1]), s[0]))
    src = sorted(staged[load], key=lambda f: (f["loc"][0], f["loc"][1]))
    # pair by index; if counts differ, pair min and flag
    n = min(len(tgt), len(src))
    for i in range(n):
        sq = tgt[i]; f = src[i]
        side = side_of.get(sq, "S")
        rot = 0.0 if side == "N" else math.pi   # face the machine
        plan.append({"id": f["id"], "load": load,
                     "x": round(sq[0], 3), "y": round(sq[1], 3), "rot": round(rot, 5)})
    if len(tgt) != len(src):
        print("!! count mismatch for", load, "tgt", len(tgt), "src", len(src))

json.dump(plan, open(os.path.join(RUN, "plan_cardio.json"), "w"), indent=0)
print("plan entries:", len(plan))
# show a few
for p in plan[:3] + plan[-3:]:
    print(p)
