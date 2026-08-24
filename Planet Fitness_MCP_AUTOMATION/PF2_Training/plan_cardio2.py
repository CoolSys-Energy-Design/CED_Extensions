# Fingerprint each A-N-RACEWAY square by the machine it serves, classify
# TREADMILL / STAIRMASTER / POWERED BIKE to match staged counts (40/12/4),
# and emit a fixture_id -> (x,y,rotation) placement plan.
import json, os, math
from collections import defaultdict

RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
ctx = json.load(open(os.path.join(RUN, "cardio_ctx.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))
squares = ctx["squares"]
staged = ctx["staged"]

# gym-equipment shapes (centroid + size)
gym = []
for s in g.get("shapes", []):
    lay = (s.get("lay") or "").upper()
    if "GYM EQUIP" in lay and "bb" in s:
        gym.append((s["c"][0], s["c"][1], max(s["bb"][2]-s["bb"][0], s["bb"][3]-s["bb"][1]),
                    s["bb"]))

# for each square, the machine extends AWAY from the row centerline.
# row at y=-241 -> machine extends north (toward -235/up); row at y=-253 -> south.
# Fingerprint: within a 3.5 ft x-window and 6 ft y-window on the machine side,
# gather gym shapes; measure footprint length (y-extent) + feature count.
def fingerprint(sx, sy):
    # machine side: squares at y=-241 serve machines to the north (y > -241)?
    # Determine by which side has more gym geometry.
    feats_n = [t for t in gym if abs(t[0]-sx) < 2.2 and 0 < (t[1]-sy) < 7]
    feats_s = [t for t in gym if abs(t[0]-sx) < 2.2 and 0 < (sy-t[1]) < 7]
    feats = feats_n if len(feats_n) >= len(feats_s) else feats_s
    side = "N" if feats_n and len(feats_n) >= len(feats_s) else "S"
    if not feats:
        return {"count": 0, "ylen": 0, "maxsz": 0, "side": side}
    ys = [t[1] for t in feats]
    ylen = max(ys) - min(ys)
    maxsz = max(t[2] for t in feats)
    return {"count": len(feats), "ylen": round(ylen, 2), "maxsz": round(maxsz, 2), "side": side}

fps = []
for (sx, sy) in squares:
    fp = fingerprint(sx, sy)
    fp["sq"] = (sx, sy)
    fps.append(fp)

# classify by feature count + footprint. Sort by count desc: bikes densest,
# treadmills longest deck, stairmasters compact.
counts = sorted(fp["count"] for fp in fps)
print("feature-count distribution:", counts)
print("ylen distribution:", sorted(round(fp["ylen"]) for fp in fps))
for fp in sorted(fps, key=lambda f: (f["sq"][1], f["sq"][0]))[:8]:
    print(fp)
json.dump(fps, open(os.path.join(RUN, "cardio_fps.json"), "w"), indent=0)
