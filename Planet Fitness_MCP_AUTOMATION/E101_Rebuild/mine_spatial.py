import json, os, re, math, collections, statistics
TRAIN = r"C:/CED_Extensions/Planet Fitness_MCP_AUTOMATION/PF2_Training"
U = 10.7639104
projects = sorted(d for d in os.listdir(TRAIN) if os.path.isdir(os.path.join(TRAIN,d)))

def norm_load(nm):
    if not nm: return None
    n = re.sub(r'\s+',' ',nm.upper().replace('\r',' ').replace('\n',' ')).strip()
    n = re.sub(r'#?\d+$','',n).strip(' -#')
    n = re.sub(r'- \d+[A-Z]?$','',n).strip(' -')
    return n

CARDIO = {'TREADMILL','STAIRMASTER','STEPMILL','POWERED BIKE','ELLIPTICAL','AMT','BIKE'}

def cluster_pts(pts, cell=0.5, min_area=1.0):
    cells = {}
    for x,y in pts:
        cells[(int(math.floor(x/cell)), int(math.floor(y/cell)))] = 0
    parent = {}
    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]; a = parent[a]
        return a
    def union(a,b):
        ra,rb = find(a),find(b)
        if ra != rb: parent[ra] = rb
    for c in cells: parent[c] = c
    for (i,j) in cells:
        for di in (-1,0,1):
            for dj in (-1,0,1):
                n = (i+di,j+dj)
                if n in cells: union((i,j),n)
    groups = collections.defaultdict(list)
    for c in cells: groups[find(c)].append(c)
    out = []
    for root, cs in groups.items():
        xs = [c[0] for c in cs]; ys = [c[1] for c in cs]
        x0,x1 = min(xs)*cell, (max(xs)+1)*cell
        y0,y1 = min(ys)*cell, (max(ys)+1)*cell
        if (x1-x0)*(y1-y0) < min_area: continue
        out.append((x0,y0,x1,y1))
    return out

def pt_seg_dist(px,py,x0,y0,x1,y1):
    dx,dy = x1-x0, y1-y0
    L2 = dx*dx+dy*dy
    if L2 == 0: return math.hypot(px-x0,py-y0), 0.0
    t = max(0, min(1, ((px-x0)*dx+(py-y0)*dy)/L2))
    return math.hypot(px-(x0+t*dx), py-(y0+t*dy)), math.degrees(math.atan2(dy,dx))%180

# accumulators
cardio_rel = collections.defaultdict(list)   # load -> list of (u,v,rot_rel, blockW,blockH, dist)
wall_rel  = collections.defaultdict(list)    # famtype-short -> (wall_d, rot_minus_wallang mod 360)
row_gap   = collections.defaultdict(list)    # load -> consecutive gaps within multi-fix circuits
tv_rel = []
proj_seen = 0

for p in projects:
    fe = os.path.join(TRAIN,p,'extract.json')
    fc = os.path.join(TRAIN,p,'cad_geom.json')
    if not (os.path.exists(fe) and os.path.exists(fc)): continue
    try:
        E = json.load(open(fe)); C = json.load(open(fc))
    except Exception as ex:
        print("skip",p,ex); continue
    proj_seen += 1
    fixById = {f['id']:f for f in E.get('fixtures',[])}
    eq_pts = []; wall_segs = []; tv_pts = []
    for s in C.get('shapes',[]):
        lay = (s.get('lay') or '').upper()
        if 'l' in s:
            x0,y0,x1,y1 = s['l']
            if ('GYM' in lay and 'EQUIP' in lay) or 'SPA EQUIP' in lay:
                eq_pts.append((x0,y0)); eq_pts.append((x1,y1))
            if 'WALL' in lay or 'DEMISING' in lay:
                wall_segs.append((x0,y0,x1,y1))
            if 'TELEVISION' in lay:
                tv_pts.append(((x0+x1)/2,(y0+y1)/2))
        elif 'bb' in s:
            x0,y0,x1,y1 = s['bb']
            if ('GYM' in lay and 'EQUIP' in lay) or 'SPA EQUIP' in lay:
                eq_pts += [(x0,y0),(x1,y1),(x0,y1),(x1,y0),tuple(s['c'])]
            if 'TELEVISION' in lay:
                tv_pts.append(tuple(s['c']))
        elif 'a' in s:
            x0,y0,x1,y1 = s['a']
            if ('GYM' in lay and 'EQUIP' in lay) or 'SPA EQUIP' in lay:
                eq_pts.append((x0,y0)); eq_pts.append((x1,y1))
    if not eq_pts: continue
    clusters = cluster_pts(eq_pts)
    WCELL = 6.0
    wgrid = collections.defaultdict(list)
    for si,(x0,y0,x1,y1) in enumerate(wall_segs):
        for i in range(int(min(x0,x1)//WCELL), int(max(x0,x1)//WCELL)+1):
            for j in range(int(min(y0,y1)//WCELL), int(max(y0,y1)//WCELL)+1):
                wgrid[(i,j)].append(si)
    def near_wall(px,py):
        best = (1e9, None)
        ci,cj = int(px//WCELL), int(py//WCELL)
        for di in (-1,0,1):
            for dj in (-1,0,1):
                for si in wgrid.get((ci+di,cj+dj),[]):
                    dd,ang = pt_seg_dist(px,py,*wall_segs[si])
                    if dd < best[0]: best = (dd,ang)
        return best
    def near_cluster(px,py):
        best = (1e9, None)
        for c in clusters:
            x0,y0,x1,y1 = c
            dx = max(x0-px, 0, px-x1); dy = max(y0-py, 0, py-y1)
            dd = math.hypot(dx,dy)
            if dd < best[0]: best = (dd, c)
        return best

    for s in E.get('systems',[]):
        ln = norm_load(s.get('load_name'))
        if not ln: continue
        members = [fixById.get(m) for m in s.get('members',[]) if m in fixById]
        members = [f for f in members if f and f.get('loc') and f['loc'][2] > -50]
        # cardio machine-frame relation
        if ln in CARDIO:
            for f in members:
                px,py = f['loc'][0], f['loc'][1]
                dd, c = near_cluster(px,py)
                if c is None or dd > 4: continue
                x0,y0,x1,y1 = c
                w,h = x1-x0, y1-y0
                cx,cy = (x0+x1)/2,(y0+y1)/2
                if w >= h:
                    u = (px-cx)/max(w,0.01); v = (py-cy)/max(h,0.01); axis = 0
                else:
                    u = (py-cy)/max(h,0.01); v = (px-cx)/max(w,0.01); axis = 90
                rot = round(math.degrees(f.get('rot') or 0)) % 360
                cardio_rel[ln].append((round(u,2), round(v,2), (rot-axis)%360, round(max(w,h),1), round(min(w,h),1), round(dd,2)))
        # wall relation for wall families
        for f in members:
            typ = f.get('type') or ''
            if 'Wall' not in typ: continue
            px,py = f['loc'][0], f['loc'][1]
            wd, wang = near_wall(px,py)
            if wang is None or wd > 3: continue
            rot = round(math.degrees(f.get('rot') or 0)) % 360
            rel = (rot - round(wang)) % 180
            wall_rel[ln if ln else typ].append((round(wd,2), rel))
        # spacing within multi-fixture circuits
        if len(members) >= 3:
            pts = sorted((f['loc'][0], f['loc'][1]) for f in members)
            gaps = []
            for i in range(len(pts)-1):
                gaps.append(math.hypot(pts[i+1][0]-pts[i][0], pts[i+1][1]-pts[i][1]))
            if gaps:
                row_gap[ln].append(round(statistics.median(gaps),1))
        # TV truss relation
        if ln and 'TV' in ln:
            for f in members:
                px,py = f['loc'][0], f['loc'][1]
                if tv_pts:
                    dmin = min(math.hypot(px-tx, py-ty) for tx,ty in tv_pts)
                    tv_rel.append(round(dmin,2))

print("projects mined:", proj_seen)
print()
print("=== CARDIO machine-frame placement (u=along axis -0.5..0.5, v=across) ===")
for ln, rows in sorted(cardio_rel.items(), key=lambda kv:-len(kv[1])):
    us = [r[0] for r in rows]; vs = [r[1] for r in rows]
    rr = collections.Counter(r[2] for r in rows)
    ws = [r[3] for r in rows]; hs = [r[4] for r in rows]
    ds = [r[5] for r in rows]
    print("  %-14s n=%4d u med=%.2f  v med=%.2f  rot-rel top=%s  block=%.1fx%.1f  d med=%.2f" % (
        ln, len(rows), statistics.median(us), statistics.median(vs), rr.most_common(3), statistics.median(ws), statistics.median(hs), statistics.median(ds)))
print()
print("=== WALL-mounted: dist + rot-wallang (mod 180) by load ===")
for ln, rows in sorted(wall_rel.items(), key=lambda kv:-len(kv[1]))[:24]:
    ds = [r[0] for r in rows]
    rels = collections.Counter(r[1] for r in rows)
    print("  %-30s n=%4d d med=%.2f rot-rel top=%s" % (str(ln)[:30], len(rows), statistics.median(ds), rels.most_common(3)))
print()
print("=== Row spacing (median gap ft) for multi-fixture circuits ===")
for ln, g in sorted(row_gap.items(), key=lambda kv:-len(kv[1]))[:15]:
    print("  %-30s n=%3d gap med=%.1f" % (ln[:30], len(g), statistics.median(g)))
print()
if tv_rel:
    print("=== TV fixtures dist to TELEVISION cad geometry: med=%.2f (n=%d) ===" % (statistics.median(tv_rel), len(tv_rel)))
json.dump({
 'cardio_rel': {k: v for k,v in cardio_rel.items()},
 'wall_rel': {str(k): v for k,v in wall_rel.items()},
 'row_gap': {k: v for k,v in row_gap.items()},
}, open('analysis/spatial_rules.json','w'))
