# Geometric placement-rationale analyzer across all PF2 projects.
# For each fixture (grouped by load name), correlate against CAD underlay features:
#   - nearest wall segment (distance, orientation, layer)
#   - nearest door
#   - nearest equipment-block polyline cluster (which layer)
#   - whether sitting on an A-N-RACEWAY power-bar square (~0.25 ft)
#   - facing direction vs wall normal (faces into room?)
# Aggregate per load name -> dominant pattern + confidence -> PLACEMENT_STATS.md + placement_stats.json
import json, os, math
from collections import defaultdict, Counter

ROOT = r"c:\CED_Extensions\Planet Fitness_MCP_AUTOMATION\PF2_Training"
IU = 10.7639
PROJECTS = sorted([d for d in os.listdir(ROOT)
                   if os.path.exists(os.path.join(ROOT, d, "extract.json"))
                   and os.path.exists(os.path.join(ROOT, d, "cad_geom.json"))])

WALL_KEYS = ("WALL", "DEMISING", "LEASE", "GLAZ", "PARTITION")
DOOR_KEYS = ("DOOR",)
EQUIP_KEYS = ("EQUIP", "FURN", "GYM", "SPA", "PLUMB FIX", "TELEVISION", "MILLWORK", "CASEWORK", "FLOR-OVHD")

def seg_dist(px, py, x0, y0, x1, y1):
    dx = x1-x0; dy = y1-y0
    L2 = dx*dx+dy*dy
    if L2 == 0:
        return math.hypot(px-x0, py-y0), 0.0
    t = max(0.0, min(1.0, ((px-x0)*dx+(py-y0)*dy)/L2))
    cx = x0+t*dx; cy = y0+t*dy
    return math.hypot(px-cx, py-cy), math.atan2(dy, dx)

class Grid:
    def __init__(self, cell=6.0):
        self.cell = cell
        self.lines = defaultdict(list)   # walls/doors as (x0,y0,x1,y1,layer)
        self.pts = defaultdict(list)     # polyline centroids (cx,cy,layer,size)
    def _key(self, x, y):
        return (int(x//self.cell), int(y//self.cell))
    def add_line(self, x0, y0, x1, y1, layer):
        for (x, y) in ((x0, y0), (x1, y1), ((x0+x1)/2, (y0+y1)/2)):
            self.lines[self._key(x, y)].append((x0, y0, x1, y1, layer))
    def add_pt(self, cx, cy, layer, size):
        self.pts[self._key(cx, cy)].append((cx, cy, layer, size))
    def near_lines(self, x, y, r=2):
        kx, ky = self._key(x, y)
        out = []
        for i in range(-r, r+1):
            for j in range(-r, r+1):
                out += self.lines.get((kx+i, ky+j), [])
        return out
    def near_pts(self, x, y, r=2):
        kx, ky = self._key(x, y)
        out = []
        for i in range(-r, r+1):
            for j in range(-r, r+1):
                out += self.pts.get((kx+i, ky+j), [])
        return out

def load(p, name):
    with open(os.path.join(ROOT, p, name)) as f:
        return json.load(f)

# accumulate per-load-name stats
stats = defaultdict(lambda: {
    "n": 0, "projects": set(), "d_wall": [], "wall_layers": Counter(), "wall_orient": Counter(),
    "faces_into_room": 0, "d_door": [], "d_equip": [], "equip_layers": Counter(),
    "on_raceway_sq": 0, "rot": Counter(), "family_type": Counter(),
})

for p in PROJECTS:
    try:
        d = load(p, "extract.json"); g = load(p, "cad_geom.json")
    except Exception:
        continue
    gw = Grid(); gd = Grid(); ge = Grid(); graceway = Grid()
    for s in g.get("shapes", []):
        lay = (s.get("lay") or "").upper()
        if "l" in s:
            x0, y0, x1, y1 = s["l"]
            if any(k in lay for k in WALL_KEYS) and "DEMO" not in lay:
                gw.add_line(x0, y0, x1, y1, lay)
            if any(k in lay for k in DOOR_KEYS):
                gd.add_line(x0, y0, x1, y1, lay)
        if "bb" in s:
            cx, cy = s["c"]
            bx = s["bb"][2]-s["bb"][0]; by = s["bb"][3]-s["bb"][1]
            size = max(bx, by)
            if any(k in lay for k in EQUIP_KEYS):
                ge.add_pt(cx, cy, lay, size)
            if "RACEWAY" in lay and 0.1 <= size <= 0.5:
                graceway.add_pt(cx, cy, lay, size)

    for f in d.get("fixtures", []):
        if not f.get("loc"):
            continue
        ln = f.get("params", {}).get("CKT_Load Name_CEDT")
        if not ln:
            continue
        x, y = f["loc"][0], f["loc"][1]
        st = stats[ln]
        st["n"] += 1; st["projects"].add(p)
        st["family_type"][f["family"] + " : " + f["type"]] += 1
        rotdeg = round((f.get("rot") or 0) * 180/math.pi) % 360
        st["rot"][rotdeg] += 1
        # nearest wall
        best_d = 9e9; best_ang = None; best_lay = None
        for (x0, y0, x1, y1, lay) in gw.near_lines(x, y, r=2):
            dd, ang = seg_dist(x, y, x0, y0, x1, y1)
            if dd < best_d:
                best_d = dd; best_ang = ang; best_lay = lay
        if best_d < 50:
            st["d_wall"].append(round(best_d, 2))
            st["wall_layers"][best_lay] += 1
            # orientation: horizontal wall (ang ~0/180) vs vertical (~90)
            if best_ang is not None:
                a = abs(math.degrees(best_ang)) % 180
                st["wall_orient"]["H" if (a < 45 or a > 135) else "V"] += 1
                # faces into room? fixture facing default +Y rotated by rot
                fdx = -math.sin(math.radians(rotdeg)); fdy = math.cos(math.radians(rotdeg))
                # wall normal (perp to wall dir)
                nx = -math.sin(best_ang); ny = math.cos(best_ang)
                # face should be opposite to direction from fixture to wall; approximate: |dot(face, normal)| high
                if abs(fdx*nx + fdy*ny) > 0.5:
                    st["faces_into_room"] += 1
        # nearest door
        bd = 9e9
        for (x0, y0, x1, y1, lay) in gd.near_lines(x, y, r=1):
            dd, _ = seg_dist(x, y, x0, y0, x1, y1)
            if dd < bd:
                bd = dd
        if bd < 50:
            st["d_door"].append(round(bd, 2))
        # nearest equipment cluster
        be = 9e9; be_lay = None
        for (cx, cy, lay, size) in ge.near_pts(x, y, r=1):
            dd = math.hypot(x-cx, y-cy)
            if dd < be:
                be = dd; be_lay = lay
        if be < 30:
            st["d_equip"].append(round(be, 2))
            st["equip_layers"][be_lay] += 1
        # on raceway square?
        onsq = False
        for (cx, cy, lay, size) in graceway.near_pts(x, y, r=1):
            if math.hypot(x-cx, y-cy) < 1.2:
                onsq = True; break
        if onsq:
            st["on_raceway_sq"] += 1

def med(xs):
    return round(sorted(xs)[len(xs)//2], 2) if xs else None

# emit
lines = ["# PF Placement Rationale — geometric stats across %d projects\n" % len(PROJECTS)]
out_json = {}
# sort by count desc
for ln in sorted(stats, key=lambda k: -stats[k]["n"]):
    st = stats[ln]
    if st["n"] < 2:
        continue
    dw = st["d_wall"]; de = st["d_equip"]; dd = st["d_door"]
    rec = {
        "count": st["n"], "n_projects": len(st["projects"]),
        "family_type": st["family_type"].most_common(2),
        "d_wall_med": med(dw), "pct_on_wall": round(100*sum(1 for v in dw if v < 1.0)/max(1, len(dw))),
        "wall_layers": st["wall_layers"].most_common(3),
        "wall_orient": dict(st["wall_orient"]),
        "pct_faces_into_room": round(100*st["faces_into_room"]/max(1, len(dw))),
        "d_door_med": med(dd), "pct_near_door_3ft": round(100*sum(1 for v in dd if v < 3.0)/max(1, len(dd))),
        "d_equip_med": med(de), "pct_on_equip_1ft": round(100*sum(1 for v in de if v < 1.0)/max(1, len(de))),
        "equip_layers": st["equip_layers"].most_common(3),
        "pct_on_raceway_square": round(100*st["on_raceway_sq"]/st["n"]),
        "rot_hist": dict(st["rot"].most_common(4)),
    }
    out_json[ln] = rec
    lines.append("## %s  (n=%d, %d projects)" % (ln, st["n"], len(st["projects"])))
    lines.append("- type: %s" % st["family_type"].most_common(1)[0][0])
    lines.append("- WALL: median dist %s ft, %d%% on-wall (<1ft) | layers %s | orient %s | %d%% face into room" % (
        rec["d_wall_med"], rec["pct_on_wall"], rec["wall_layers"], rec["wall_orient"], rec["pct_faces_into_room"]))
    lines.append("- DOOR: median %s ft, %d%% within 3 ft" % (rec["d_door_med"], rec["pct_near_door_3ft"]))
    lines.append("- EQUIP: median %s ft to nearest block, %d%% on-equip (<1ft) | layers %s" % (
        rec["d_equip_med"], rec["pct_on_equip_1ft"], rec["equip_layers"]))
    lines.append("- RACEWAY power-bar square: %d%% sit on one" % rec["pct_on_raceway_square"])
    lines.append("- rotations: %s" % rec["rot_hist"])
    lines.append("")

with open(os.path.join(ROOT, "PLACEMENT_STATS.md"), "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
with open(os.path.join(ROOT, "placement_stats.json"), "w") as f:
    json.dump(out_json, f, indent=1)
print("wrote PLACEMENT_STATS.md (%d load types) over %d projects" % (len(out_json), len(PROJECTS)))
