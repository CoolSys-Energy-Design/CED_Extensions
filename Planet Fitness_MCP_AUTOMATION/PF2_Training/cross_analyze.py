# Cross-project synthesis over all PF2_Training extracts (CPython).
# Emits CROSS_REPORT.md + machine-readable knowledge JSONs for the PF-Power-Plan-2 skill.
import json, os, math, re
from collections import defaultdict, Counter

ROOT = r"c:\CED_Extensions\Planet Fitness_MCP_AUTOMATION\PF2_Training"
# auto-discover every project folder that has an extract.json
PROJECTS = sorted([d for d in os.listdir(ROOT)
                   if os.path.exists(os.path.join(ROOT, d, "extract.json"))])
IU = 10.7639  # internal VA/V -> real

def load(p, name):
    fp = os.path.join(ROOT, p, name)
    if not os.path.exists(fp):
        return None
    with open(fp) as f:
        return json.load(f)

def d2(a, b):
    return math.hypot(a[0]-b[0], a[1]-b[1])

def median(xs):
    xs = sorted(xs)
    return xs[len(xs)//2] if xs else None

out = []
w = out.append
K = {}  # knowledge dump

# ---------- 1. project dates ----------
w("# PF2 Cross-Project Report")
w("")
w("## Project dates (for trust ranking)")
dates = {}
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    proj = d.get("project", {})
    issues = [s.get("issue") for s in d.get("all_sheets", []) if s.get("issue")]
    cnt = Counter(issues)
    date_keys = {k: v for k, v in proj.items() if "date" in k.lower() or "issue" in k.lower()}
    dates[p] = {"sheet_issue_dates": cnt.most_common(5), "proj_date_params": date_keys}
    w("- **%s**: sheet issues=%s | proj params=%s" % (p, cnt.most_common(4), date_keys))
w("")
K["dates"] = dates

# ---------- 2. keynote legend text per project ----------
w("## Keynote legends (sheet key-note drafting views)")
legends = {}
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    txts = [t["text"] for t in d.get("text_sheet", [])]
    # keynote legend text = numbered lines
    leg = {}
    for t in txts:
        lines = re.split(r"[\r\n\x0b]+", t)
        num = 0
        looks = sum(1 for L in lines if re.match(r"^\s*\d+[\.\)]?\s+\S", L))
        if looks >= 3:
            for L in lines:
                m = re.match(r"^\s*(\d+)[\.\)]?\s+(.*\S)", L)
                if m:
                    leg[m.group(1)] = m.group(2)[:150]
        elif looks == 0 and len(lines) > 3:
            # legend may be sequential lines w/o numbers; keep raw
            pass
    legends[p] = leg
    w("### %s (%d entries)" % (p, len(leg)))
    for k in sorted(leg, key=lambda x: int(x)):
        w("- #%s: %s" % (k, leg[k]))
    w("")
K["keynote_legends"] = legends

# ---------- 3. load table by project ----------
w("## Load values by load-name (VA real units, V real, panel, rating, poles)")
loadtab = defaultdict(dict)
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    groups = defaultdict(list)
    for f in d.get("fixtures", []):
        prm = f.get("params", {})
        ln = prm.get("CKT_Load Name_CEDT")
        if not ln: continue
        groups[ln].append(f)
    for ln, g in groups.items():
        p0 = g[0]["params"]
        va = p0.get("CKT_Apparent Load_CED") or p0.get("Apparent Load Input_CED") or 0
        vv = p0.get("CKT_Voltage_CED") or p0.get("Voltage_CED") or 0
        loadtab[ln][p] = {
            "count": len(g),
            "family": g[0]["family"], "type": g[0]["type"],
            "va": round(float(va)/IU, 0), "v": round(float(vv)/IU, 0),
            "panel": p0.get("CKT_Panel_CEDT"),
            "rating": p0.get("CKT_Rating_CED"),
            "poles": p0.get("Number of Poles_CED") or p0.get("CKT_Number of Poles_CED"),
        }
for ln in sorted(loadtab):
    entries = loadtab[ln]
    w("### %s" % ln)
    for p in PROJECTS:
        if p in entries:
            e = entries[p]
            w("- %s: [%d] %s:%s | %sVA %sV %s %sA %sP" % (
                p, e["count"], e["family"], e["type"], e["va"], e["v"],
                e["panel"], e["rating"], e["poles"]))
    w("")
K["load_table"] = {k: v for k, v in loadtab.items()}

# ---------- 4. per-room inventory ----------
w("## Per-room fixture inventory (space bbox containment)")
room_inv = {}
def room_key(name):
    # normalize room names across projects
    n = (name or "").upper()
    n = re.sub(r"\s*\d+[A-Z]?\s*$", "", n).strip()
    n = n.replace("EX'G ", "").replace("EXG ", "")
    return n
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    spaces = [s for s in d.get("spaces", []) if s.get("bb")]
    inv = defaultdict(Counter)
    for f in d.get("fixtures", []):
        if not f.get("loc"): continue
        ln = f.get("params", {}).get("CKT_Load Name_CEDT")
        label = "%s | %s" % (f["type"], ln or "-")
        best = None; best_area = None
        for s in spaces:
            (x0,y0,_), (x1,y1,_) = s["bb"][0], s["bb"][1]
            if x0-0.5 <= f["loc"][0] <= x1+0.5 and y0-0.5 <= f["loc"][1] <= y1+0.5:
                area = (x1-x0)*(y1-y0)
                if best is None or area < best_area:
                    best = s; best_area = area
        rk = room_key(best["name"]) if best else "(outside spaces)"
        inv[rk][label] += 1
    room_inv[p] = {k: dict(v) for k, v in inv.items()}
# print merged by room key
allrooms = sorted(set(k for p in room_inv for k in room_inv[p]))
for rk in allrooms:
    w("### %s" % rk)
    for p in PROJECTS:
        if p in room_inv and rk in room_inv[p]:
            items = sorted(room_inv[p][rk].items(), key=lambda kv: -kv[1])
            w("- %s: %s" % (p, "; ".join("%dx %s" % (c, l) for l, c in items)))
    w("")
K["room_inventory"] = room_inv

# ---------- 5. wires ----------
w("## Wire conventions (E101 view only)")
wire_stats = {}
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    pv = d["plan_view"]["id"]
    wires = [x for x in d.get("wires", []) if x.get("ownerview") in (None, pv)]
    fixtures = [f for f in d.get("fixtures", []) if f.get("loc")]
    fxpts = [(f["loc"][0], f["loc"][1]) for f in fixtures]
    def near_fx(pt2):
        best = 9e9
        for q in fxpts:
            dd = math.hypot(pt2[0]-q[0], pt2[1]-q[1])
            if dd < best: best = dd
        return best
    nv = Counter(len(x.get("verts") or []) for x in wires)
    seglens = []
    end_near = Counter()  # both/one/none endpoints near a fixture
    for x in wires:
        vs = x.get("verts") or []
        if len(vs) >= 2:
            L = sum(math.hypot(vs[i+1][0]-vs[i][0], vs[i+1][1]-vs[i][1]) for i in range(len(vs)-1))
            seglens.append(round(L,1))
            n0 = near_fx(vs[0][:2]); n1 = near_fx(vs[-1][:2])
            near = (n0 < 1.0) + (n1 < 1.0)
            end_near[near] += 1
    wtypes = Counter(x.get("type") for x in wires)
    wire_stats[p] = {"count": len(wires), "types": dict(wtypes), "verts": dict(nv),
                     "median_len_ft": median(seglens), "ends_near_fixture": dict(end_near)}
    w("- **%s**: %d wires | types=%s | verts=%s | median len=%s ft | endpoints near fixture (2=both,1=one,0=none): %s" % (
        p, len(wires), dict(wtypes), dict(nv), median(seglens), dict(end_near)))
w("")
K["wire_stats"] = wire_stats

# ---------- 6. wire tags ----------
w("## Wire tags (E101)")
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    pv = d["plan_view"]["id"]
    wtags = [x for x in d.get("wire_tags", []) if x.get("ownerview") in (None, pv)]
    texts = Counter(x.get("text") or "" for x in wtags)
    w("- **%s** (%d): %s" % (p, len(wtags), [t for t, c in texts.most_common(12)]))
w("")

# ---------- 7. fixture tags ----------
w("## Fixture tag conventions (E101)")
tag_stats = {}
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    pv = d["plan_view"]["id"]
    fx_by_id = {f["id"]: f for f in d.get("fixtures", [])}
    ftags = [x for x in d.get("fixture_tags", []) if x.get("ownerview") in (None, pv)]
    per_type = defaultdict(lambda: {"n": 0, "dx": [], "dy": [], "texts": Counter()})
    for t in ftags:
        key = t.get("type")
        e = per_type[key]
        e["n"] += 1
        e["texts"][t.get("text") or ""] += 1
        hosts = t.get("host") or []
        if t.get("head") and hosts and hosts[0] in fx_by_id and fx_by_id[hosts[0]].get("loc"):
            h = fx_by_id[hosts[0]]["loc"]
            e["dx"].append(t["head"][0]-h[0]); e["dy"].append(t["head"][1]-h[1])
    tag_stats[p] = {}
    for k, e in per_type.items():
        tag_stats[p][k] = {"n": e["n"],
                           "med_off": [round(median(e["dx"]) or 0, 2), round(median(e["dy"]) or 0, 2)]}
        w("- **%s** %s: n=%d med_offset=(%.2f, %.2f) samples=%s" % (
            p, k, e["n"], median(e["dx"]) or 0, median(e["dy"]) or 0,
            [t for t, c in e["texts"].most_common(3)]))
w("")
K["fixture_tag_stats"] = tag_stats

# ---------- 8. keynote-to-fixture association ----------
w("## Keynote number -> associated fixture/load (nearest fixture within 8 ft)")
kn_assoc = {}
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    fixtures = [f for f in d.get("fixtures", []) if f.get("loc")]
    assoc = defaultdict(Counter)
    dists = defaultdict(list)
    for n in d.get("keynotes", []):
        if not n.get("loc"): continue
        num = str(n.get("params", {}).get("CED-G-NOTE #") or "?")
        best = None; bd = 9e9
        for f in fixtures:
            dd = d2(n["loc"], f["loc"])
            if dd < bd: bd = dd; best = f
        if best and bd <= 8.0:
            ln = best.get("params", {}).get("CKT_Load Name_CEDT") or best["type"]
            assoc[num][ln] += 1
            dists[num].append(round(bd,1))
    kn_assoc[p] = {k: {"loads": dict(v), "med_dist": median(dists[k])} for k, v in assoc.items()}
    w("### %s" % p)
    for k in sorted(assoc, key=lambda x: (len(x), x)):
        w("- #%s -> %s (med dist %s ft)" % (k, dict(assoc[k].most_common(3)), median(dists[k])))
    w("")
K["keynote_assoc"] = kn_assoc

# ---------- 9. wall-relative placement ----------
w("## Wall-relative placement check (wall-type fixtures vs CAD wall lines)")
for p in PROJECTS:
    d = load(p, "extract.json")
    g = load(p, "cad_geom.json")
    if not d or not g: continue
    walls = []
    for s in g.get("shapes", []):
        lay = (s.get("lay") or "").upper()
        if "WALL" in lay or "DEMISING" in lay or "LEASE" in lay or "GLAZ" in lay:
            if "l" in s:
                walls.append(s["l"])
    def dist_to_wall(x, y):
        best = 9e9
        for (x0,y0,x1,y1) in walls:
            dx = x1-x0; dy = y1-y0
            L2 = dx*dx+dy*dy
            if L2 == 0: continue
            t = max(0, min(1, ((x-x0)*dx+(y-y0)*dy)/L2))
            px = x0+t*dx; py = y0+t*dy
            dd = math.hypot(x-px, y-py)
            if dd < best: best = dd
        return best
    offs = []
    n_far = 0
    import random
    random.seed(1)
    wall_fx = [f for f in d.get("fixtures", []) if f.get("loc") and "Wall" in (f.get("type") or "")]
    sample = wall_fx if len(wall_fx) <= 80 else random.sample(wall_fx, 80)
    for f in sample:
        dd = dist_to_wall(f["loc"][0], f["loc"][1])
        offs.append(round(dd,2))
        if dd > 1.0: n_far += 1
    w("- **%s**: %d wall fixtures sampled=%d | median dist to CAD wall=%.2f ft | >1ft: %d (%.0f%%) | n_wall_segs=%d" % (
        p, len(wall_fx), len(sample), median(offs) or -1, n_far,
        100.0*n_far/max(1,len(sample)), len(walls)))
w("")

# ---------- 10. panels ----------
w("## Panels per project")
for p in PROJECTS:
    d = load(p, "extract.json")
    if not d: continue
    rows = []
    for e in d.get("equipment", []):
        prm = e.get("params", {})
        pn = prm.get("Panel Name")
        if pn:
            rows.append("%s(%s)" % (pn, e["type"].split(" - ")[0]))
    w("- **%s**: %s" % (p, ", ".join(sorted(set(rows)))))
w("")

with open(os.path.join(ROOT, "CROSS_REPORT.md"), "w", encoding="utf-8") as f:
    f.write("\n".join(out))
with open(os.path.join(ROOT, "knowledge.json"), "w", encoding="utf-8") as f:
    json.dump(K, f, indent=1)
print("wrote CROSS_REPORT.md (%d lines) + knowledge.json" % len(out))
