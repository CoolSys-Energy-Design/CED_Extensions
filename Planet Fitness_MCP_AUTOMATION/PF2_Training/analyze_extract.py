# Digest a PF2_Training extract.json into a compact summary.md (CPython)
import json, sys, os, math, re
from collections import defaultdict, Counter

def dist2(a, b):
    return math.hypot(a[0]-b[0], a[1]-b[1])

def main(folder):
    with open(os.path.join(folder, "extract.json")) as f:
        d = json.load(f)
    out = []
    w = out.append
    w("# %s  (project %s / %s)" % (d["doc_title"], d["project"].get("number"), d["project"].get("name")))
    w("Sheet: %s  |  Plan view: %s (scale 1:%s)" % (
        d.get("sheet", {}).get("number") if d.get("sheet") else None,
        d["plan_view"]["name"], d["plan_view"]["scale"]))
    w("")

    # spaces
    spaces = d.get("spaces", [])
    w("## Spaces (%d)" % len(spaces))
    for s in sorted(spaces, key=lambda s: s.get("number") or ""):
        w("- %s %s  ctr=%s" % (s.get("number"), s.get("name"),
          [round(x,1) for x in s["loc"][:2]] if s.get("loc") else s.get("bb")))
    w("")

    # panels
    w("## Panels (%d)" % len(d.get("equipment", [])))
    for e in d.get("equipment", []):
        p = e.get("params", {})
        w("- %s | %s:%s | loc=%s | mains=%s volts=%s" % (
            p.get("Panel Name"), e["family"], e["type"],
            [round(x,1) for x in e["loc"][:2]] if e.get("loc") else None,
            p.get("Mains"), p.get("Distribution System")))
    w("")

    # fixture inventory by family:type + load
    fixtures = d.get("fixtures", [])
    pv_id = d["plan_view"]["id"]
    fx_by_id = {f["id"]: f for f in fixtures}
    w("## Fixtures in plan view (%d)" % len(fixtures))
    groups = defaultdict(list)
    for f in fixtures:
        p = f.get("params", {})
        key = (f["family"], f["type"], p.get("CKT_Load Name_CEDT") or p.get("Load Name") or "?")
        groups[key].append(f)
    for key in sorted(groups, key=lambda k: (-len(groups[k]), k[0])):
        fam, typ, load = key
        g = groups[key]
        p0 = g[0].get("params", {})
        rots = Counter(round((f.get("rot") or 0) * 180 / math.pi) % 360 for f in g)
        w("- [%d] %s : %s | load=%s | VA=%s V=%s panel=%s rating=%s poles=%s | rot=%s" % (
            len(g), fam, typ, load,
            p0.get("CKT_Apparent Load_CED") or p0.get("Apparent Load Input_CED"),
            p0.get("CKT_Voltage_CED") or p0.get("Voltage_CED"),
            p0.get("CKT_Panel_CEDT"), p0.get("CKT_Rating_CED"),
            p0.get("Number of Poles_CED") or p0.get("CKT_Number of Poles_CED"),
            dict(rots)))
    w("")

    # circuits by panel
    systems = d.get("systems", [])
    w("## Circuits (%d)" % len(systems))
    bypanel = defaultdict(list)
    for s in systems:
        bypanel[s.get("panel") or "?"].append(s)
    for pn in sorted(bypanel):
        w("### Panel %s (%d ckts)" % (pn, len(bypanel[pn])))
        for s in sorted(bypanel[pn], key=lambda s: (len(s.get("circuit") or ""), s.get("circuit") or "")):
            w("- ckt %s | %s | %sA %sP %sV | %s VA | %d fixtures" % (
                s.get("circuit"), s.get("load_name"), s.get("rating"),
                s.get("poles"), s.get("volts"), s.get("app_load_va"),
                len(s.get("members", []))))
    w("")

    # keynotes
    notes = d.get("keynotes", [])
    w("## Keynotes / generic annotations in view (%d)" % len(notes))
    kn_groups = defaultdict(list)
    for n in notes:
        num = n.get("params", {}).get("CED-G-NOTE #") or n.get("params", {}).get("Key Note Number") or "?"
        kn_groups[(n.get("family"), n.get("type"), str(num))].append(n)
    for key in sorted(kn_groups, key=lambda k: (k[0] or "", len(k[2]), k[2])):
        fam, typ, num = key
        g = kn_groups[key]
        # nearest fixture distance for each keynote
        dists = []
        for n in g:
            if not n.get("loc"): continue
            best = None
            for f in fixtures:
                if not f.get("loc"): continue
                dd = dist2(n["loc"], f["loc"])
                if best is None or dd < best: best = dd
            if best is not None: dists.append(best)
        w("- [%d] %s:%s #%s | nearest-fixture dist min/med/max = %s" % (
            len(g), fam, typ, num,
            [round(min(dists),1), round(sorted(dists)[len(dists)//2],1), round(max(dists),1)] if dists else None))
    w("")

    # fixture tags: offsets from host
    ftags = [x for x in d.get("fixture_tags", []) if x.get("ownerview") in (None, pv_id)]
    w("## Fixture tags (%d)" % len(ftags))
    t_groups = defaultdict(list)
    for t in ftags:
        t_groups[(t.get("family"), t.get("type"))].append(t)
    for key, g in sorted(t_groups.items(), key=lambda kv: -len(kv[1])):
        offs = []
        texts = Counter()
        for t in g:
            texts[t.get("text") or ""] += 1
            hosts = t.get("host") or []
            if t.get("head") and hosts and hosts[0] in fx_by_id and fx_by_id[hosts[0]].get("loc"):
                h = fx_by_id[hosts[0]]["loc"]
                offs.append((round(t["head"][0]-h[0], 2), round(t["head"][1]-h[1], 2)))
        med = None
        if offs:
            xs = sorted(o[0] for o in offs); ys = sorted(o[1] for o in offs)
            med = (xs[len(xs)//2], ys[len(ys)//2])
        w("- [%d] %s:%s | median offset from fixture=%s | sample texts=%s" % (
            len(g), key[0], key[1], med, list(texts.most_common(5))))
    w("")

    # wires
    wires = [x for x in d.get("wires", []) if x.get("ownerview") in (None, pv_id)]
    w("## Wires (%d)" % len(wires))
    wtypes = Counter((wi.get("type"), wi.get("wiring_type"), len(wi.get("verts") or [])) for wi in wires)
    for (t, wt, nv), c in wtypes.most_common(20):
        w("- [%d] type=%s wiring=%s verts=%s" % (c, t, wt, nv))
    w("")

    # wire tags
    wtags = [x for x in d.get("wire_tags", []) if x.get("ownerview") in (None, pv_id)]
    w("## Wire tags (%d)" % len(wtags))
    texts = Counter(t.get("text") or "" for t in wtags)
    for txt, c in texts.most_common(40):
        w("- [%d] %r" % (c, txt))
    w("")

    # sheet text (keynote legend)
    w("## Sheet text (legend candidates)")
    for t in d.get("text_sheet", [])[:80]:
        txt = t["text"].replace("\r", " / ").replace("\n", " / ")
        w("- %r" % txt[:220])
    w("")
    w("## Plan-view text (first 40)")
    for t in d.get("text_plan", [])[:40]:
        txt = t["text"].replace("\r", " / ").replace("\n", " / ")
        w("- %r" % txt[:160])
    w("")

    # links
    w("## Links")
    for l in d.get("links", []):
        w("- %s %s origin=%s basisX=%s" % (l["kind"], l.get("name"),
          [round(x,2) for x in (l.get("origin") or [])][:2],
          [round(x,2) for x in (l.get("basisX") or [])][:2]))
    w("")
    # leaders count
    w("## Leaders/detail lines in view: %d" % len(d.get("leaders", [])))

    with open(os.path.join(folder, "summary.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print("wrote", os.path.join(folder, "summary.md"), "(%d lines)" % len(out))

if __name__ == "__main__":
    main(sys.argv[1])
