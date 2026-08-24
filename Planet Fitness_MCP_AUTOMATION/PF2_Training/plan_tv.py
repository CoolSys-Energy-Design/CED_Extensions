import json, os, math
from collections import defaultdict
RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
d = json.load(open(os.path.join(RUN, "extract.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

# TV blocks: A-N-TELEVISION shapes (dedupe clusters)
tv = []
for s in g.get("shapes", []):
    if "TELEVISION" in (s.get("lay") or "").upper() and "bb" in s:
        tv.append((s["c"][0], s["c"][1], s["bb"]))
# cluster within 2 ft
clusters = []
for (x, y, bb) in sorted(tv):
    hit = None
    for c in clusters:
        if math.hypot(x-c["x"], y-c["y"]) < 2.5:
            hit = c; break
    if hit:
        hit["pts"].append((x, y)); hit["x"]=sum(p[0] for p in hit["pts"])/len(hit["pts"]); hit["y"]=sum(p[1] for p in hit["pts"])/len(hit["pts"])
    else:
        clusters.append({"x": x, "y": y, "pts": [(x, y)]})
print("TV clusters (blocks):", len(clusters))
# show y-distribution to find truss row
ys = defaultdict(int)
for c in clusters:
    ys[round(c["y"]/5)*5] += 1
for yy in sorted(ys, key=lambda k:-ys[k])[:10]:
    xs = [c["x"] for c in clusters if round(c["y"]/5)*5==yy]
    print("  y~%d : %d TVs  x[%.0f..%.0f]" % (yy, ys[yy], min(xs), max(xs)))

# staged TV TRUSS
truss = [{"id": f["id"], "loc": f["loc"][:2]} for f in d.get("fixtures", [])
         if f.get("params", {}).get("CKT_Load Name_CEDT") == "TV TRUSS" and f.get("loc")]
print("staged TV TRUSS:", len(truss), "at y~", set(round(t['loc'][1]) for t in truss))
json.dump({"clusters": clusters, "truss": truss}, open(os.path.join(RUN, "tv_ctx.json"), "w"))
