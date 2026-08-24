# Consolidate synthesis facts across all 26 projects into SYNTHESIS_FACTS.json:
#  - recency ranking (latest parseable sheet date per project)
#  - canonical load table: dominant VA/V/panel/type per load name + variants
#  - keynote number -> load map (voted across projects)
#  - fixture-tag offset convention
import json, os, re
from collections import defaultdict, Counter
from datetime import datetime

ROOT = r"c:\CED_Extensions\Planet Fitness_MCP_AUTOMATION\PF2_Training"
IU = 10.7639
PROJECTS = sorted([d for d in os.listdir(ROOT)
                   if os.path.exists(os.path.join(ROOT, d, "extract.json"))])

def load(p):
    with open(os.path.join(ROOT, p, "extract.json")) as f:
        return json.load(f)

def parse_date(s):
    if not s: return None
    s = s.strip()
    for fmt in ("%m/%d/%y", "%m/%d/%Y", "%Y-%m-%d", "%m-%d-%y", "%m-%d-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

# ---- recency ----
recency = {}
for p in PROJECTS:
    d = load(p)
    dates = [parse_date(s.get("issue")) for s in d.get("all_sheets", [])]
    dates = [x for x in dates if x and x.year >= 2020]
    latest = max(dates) if dates else None
    recency[p] = latest.strftime("%Y-%m-%d") if latest else None
rank = sorted(PROJECTS, key=lambda p: (recency[p] or "0000"), reverse=True)

# ---- canonical load table ----
load_rows = defaultdict(list)
for p in PROJECTS:
    d = load(p)
    seen = {}
    for f in d.get("fixtures", []):
        prm = f.get("params", {})
        ln = prm.get("CKT_Load Name_CEDT")
        if not ln: continue
        va = prm.get("CKT_Apparent Load_CED") or prm.get("Apparent Load Input_CED") or 0
        vv = prm.get("CKT_Voltage_CED") or prm.get("Voltage_CED") or 0
        key = ln
        rec = {"proj": p, "type": f["family"] + " : " + f["type"],
               "va": round(float(va)/IU), "v": round(float(vv)/IU),
               "panel": prm.get("CKT_Panel_CEDT"),
               "rating": prm.get("CKT_Rating_CED"),
               "poles": prm.get("Number of Poles_CED") or prm.get("CKT_Number of Poles_CED")}
        # one representative per project per load
        if key not in seen:
            seen[key] = rec
    for k, rec in seen.items():
        load_rows[k].append(rec)

canonical = {}
for ln, rows in load_rows.items():
    types = Counter(r["type"] for r in rows)
    vas = Counter(r["va"] for r in rows)
    vs = Counter(r["v"] for r in rows)
    panels = Counter(r["panel"] for r in rows if r["panel"])
    ratings = Counter(r["rating"] for r in rows if r["rating"])
    poles = Counter(r["poles"] for r in rows if r["poles"])
    # recency-weighted VA: prefer the value in the most-recent project that has it
    va_recent = None
    for p in rank:
        hit = [r for r in rows if r["proj"] == p]
        if hit:
            va_recent = hit[0]["va"]; break
    canonical[ln] = {
        "projects": len(rows),
        "type": types.most_common(1)[0][0], "type_variants": types.most_common(4),
        "va_mode": vas.most_common(1)[0][0], "va_recent": va_recent, "va_variants": vas.most_common(4),
        "v_mode": vs.most_common(1)[0][0],
        "panel_mode": panels.most_common(1)[0][0] if panels else None, "panel_variants": panels.most_common(4),
        "rating_mode": ratings.most_common(1)[0][0] if ratings else None,
        "poles_mode": poles.most_common(1)[0][0] if poles else None,
    }

# ---- keynote number -> load (voted) ----
kn_votes = defaultdict(Counter)
for p in PROJECTS:
    d = load(p)
    fixtures = [f for f in d.get("fixtures", []) if f.get("loc")]
    import math
    for n in d.get("keynotes", []):
        if not n.get("loc"): continue
        num = str(n.get("params", {}).get("CED-G-NOTE #") or "?")
        if num == "?": continue
        best = None; bd = 9e9
        for f in fixtures:
            dd = math.hypot(n["loc"][0]-f["loc"][0], n["loc"][1]-f["loc"][1])
            if dd < bd: bd = dd; best = f
        if best and bd <= 6.0:
            ln = best.get("params", {}).get("CKT_Load Name_CEDT") or best["type"]
            kn_votes[num][ln] += 1
keynote_map = {k: v.most_common(5) for k, v in sorted(kn_votes.items(), key=lambda kv: (len(kv[0]), kv[0]))}

# ---- fixture tag offset ----
tag_off = []
for p in PROJECTS:
    d = load(p)
    pv = d["plan_view"]["id"]
    fx = {f["id"]: f for f in d.get("fixtures", [])}
    for t in d.get("fixture_tags", []):
        if t.get("ownerview") not in (None, pv): continue
        if t.get("type") != "Panel & Circuit Number": continue
        h = t.get("host") or []
        if t.get("head") and h and h[0] in fx and fx[h[0]].get("loc"):
            f = fx[h[0]]
            tag_off.append((round(t["head"][0]-f["loc"][0],2), round(t["head"][1]-f["loc"][1],2)))
import statistics
tag_dx = statistics.median([o[0] for o in tag_off]) if tag_off else None
tag_dy = statistics.median([o[1] for o in tag_off]) if tag_off else None

out = {
    "n_projects": len(PROJECTS),
    "recency_rank": [(p, recency[p]) for p in rank],
    "canonical_loads": canonical,
    "keynote_map": keynote_map,
    "fixture_tag_offset_ft": [round(tag_dx,2) if tag_dx else None, round(tag_dy,2) if tag_dy else None],
    "tag_offset_note": "Panel & Circuit Number tag median offset from fixture (internal ft, +y is north)",
}
with open(os.path.join(ROOT, "SYNTHESIS_FACTS.json"), "w") as f:
    json.dump(out, f, indent=1)

print("RECENCY (newest first):")
for p in rank[:8]:
    print("  %-24s %s" % (p, recency[p]))
print("\nKEYNOTE MAP:")
for k, v in keynote_map.items():
    print("  #%s -> %s" % (k, v[:3]))
print("\nTAG OFFSET (ft):", [round(tag_dx,2) if tag_dx else None, round(tag_dy,2) if tag_dy else None])
print("\nCANONICAL LOADS:", len(canonical))
