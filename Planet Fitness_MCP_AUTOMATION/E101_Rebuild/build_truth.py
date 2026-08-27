import json, math, re, collections, statistics
U = 10.7639104
E = json.load(open('current_model/extract.json'))
E2 = json.load(open('current_model/extract2.json'))
VIEW = E['plan_view']['id']

def deg(r): return (round(math.degrees(r)) % 360) if r is not None else None

# ---------- fixtures ----------
def linker(p):
    s = p.get('Element_Linker') or ''
    out = {}
    m = re.search(r'Set Definition ID: (\S+?),', s)
    if m: out['set'] = m.group(1)
    m = re.search(r'Host Name: (.*?), Parent_location', s)
    if m: out['host'] = m.group(1)
    m = re.search(r'Parent_location: ([\-\d\.]+),([\-\d\.]+)', s)
    if m: out['ploc'] = [float(m.group(1)), float(m.group(2))]
    m = re.search(r'Rotation \(deg\): ([\-\d\.]+)', s)
    if m: out['prot'] = float(m.group(1))
    return out

dd_ids = set(x['id'] for x in E2['data_devices'])
# map member id -> (panel, circuit, load, system id) from electrical systems (authoritative)
memsys = {}
for s in E['systems']:
    for m in s.get('members', []):
        memsys[m] = (s.get('panel'), s.get('circuit'), s.get('load_name'), s['id'])

viewset = set(v['id'] for v in E['view_instances'])
fixtures = []
for f in E['fixtures'] + E2['data_devices']:
    if not f.get('loc') or f['loc'][2] < -50:  continue
    if f['id'] not in viewset: continue
    p = f.get('params', {})
    va = p.get('Apparent Load_CED')
    ms = memsys.get(f['id'], (None, None, None, None))
    fixtures.append({
        'id': f['id'],
        'cat': 'data' if f['id'] in dd_ids else 'elec',
        'fam': f['family'], 'typ': f['type'],
        'x': round(f['loc'][0],3), 'y': round(f['loc'][1],3),
        'rot': deg(f.get('rot')),
        'elev': p.get('Elevation from Level'),
        'load': ms[2] or p.get('CKT_Load Name_CEDT'),
        'panel': ms[0] or p.get('CKT_Panel_CEDT') or p.get('Panel'),
        'ckt': ms[1] or p.get('CKT_Circuit Number_CEDT') or p.get('Circuit Number'),
        'sys': ms[3],
        'va': round(va/U,0) if va else None,
        'group': p.get('dev-Group ID'),
        'link': linker(p),
        'mark': p.get('Mark'),
    })
print("plan fixtures+data:", len(fixtures))

# ---------- circuits ----------
circuits = []
for s in E['systems']:
    ln = s.get('load_name')
    if not ln or ln.upper() in ('SPARE','SPACE','HIGH LEG SPACE'): continue
    circuits.append({'id': s['id'], 'panel': s['panel'], 'ckt': s['circuit'],
        'load': ln, 'rating': s['rating'], 'poles': s['poles'],
        'volts': round(s['volts']/U) if s['volts'] else None,
        'va': round(s['app_load_va']/U) if s['app_load_va'] else None,
        'members': s['members']})
print("real circuits:", len(circuits))

# ---------- annotations in view ----------
ftags = [t for t in E['fixture_tags'] if t['ownerview'] == VIEW]
etags = [t for t in E['equip_tags'] if t['ownerview'] == VIEW]
wtags = [t for t in E['wire_tags'] if t['ownerview'] == VIEW]
wires = [w for w in E['wires'] if w['ownerview'] == VIEW]
print("view tags: fix", len(ftags), "equip", len(etags), "wire", len(wtags), "| wires", len(wires))

# legend text -> number (order of entries in the E101 keynote legend)
LEGEND = [
 (1, 'GFCI FACEPLATE RECEPTACLE ABOVE COUNTER'),
 (2, 'HAND DRYER MOUNTED AT 42'),
 (3, 'PROVIDE A GFCI DUPLEX RECEPTACLE MOUNTED BELOW COUNTERTOP'),
 (4, 'HYDROMASSAGE UNIT. 208 VOLT'),
 (5, 'TANNING AND HYBRID BED/BOOTH'),
 (6, 'NOT USED'),
 (7, 'PROVIDE ISOLATED GROUND FOR CHECK-IN'),
 (8, 'PHONE JACK AND RECEPTACLE FOR T-MAX'),
 (9, 'LOCATION OF DATA RACK AND TELEVISION DISTRIBUTION'),
 (10, 'COORDINATE LOCATION OF BACKWRAP AND CHECK-IN'),
 (11, 'COORDINATE TV MOUNTING HEIGHT'),
 (12, 'MOUNT RECEPTACLE SERVING DRINKING FOUNTAIN'),
 (13, 'TRENCH POWER AND DATA FROM END OF GRATEFUL'),
 (14, 'COORDINATE CLUB COMM CONNECTIONS'),
 (15, 'MOUNT RECEPTACLES/USB PORTS HORIZONTALLY'),
 (16, 'BEVERAGE COOLER'),
 (17, 'POWER CONNECTION FOR SIGNAGE'),
 (18, 'PLANET FITNESS CLOCK'),
 (19, 'COORDINATE ELECTRICAL REQUIREMENTS WITH FAN'),
 (20, 'TRENCH POWER AND DATA FROM RECEPTACLE TO NEAREST'),
 (21, 'SPRAY TANNING BOOTH'),
 (22, 'REDZONE SAUNA UNIT'),
]
def resolve_kn(val):
    if val is None: return None
    s = str(val).strip()
    if re.match(r'^\d+$', s): return int(s)
    su = s.upper()
    for num, pref in LEGEND:
        if su.startswith(pref[:28]): return num
    return None

keynotes = []
for k in E['keynotes']:
    p = k.get('params', {})
    raw = None
    if p.get('CED-G-NOTE #') not in (None, '', 'XXXX'):
        raw = p.get('CED-G-NOTE #')
    else:
        for key in p:
            kl = key.lower()
            if (('key' in kl or 'ced-g' in kl) and 'note' in kl) or kl in ('number','note number','label'):
                if p[key] not in (None, '', 'XXXX'):
                    raw = p[key]; break
    keynotes.append({'id': k['id'], 'fam': k['family'], 'typ': k['type'],
        'loc': k['loc'][:2] if k.get('loc') else None, 'rot': deg(k.get('rot')),
        'num': resolve_kn(raw), 'rawlen': len(str(raw)) if raw else 0})
kn_missing = [k for k in keynotes if k['num'] is None]
print("keynotes:", len(keynotes), "missing num:", len(kn_missing))
if kn_missing:
    for k in kn_missing[:3]:
        kk = [x for x in E['keynotes'] if x['id']==k['id']][0]
        print("  unresolved:", {a:b for a,b in kk.get('params',{}).items() if 'NOTE' in a.upper() or a in ('Type','Family')})

# ---------- equipment clustering from CAD ----------
d = json.load(open('current_model/cad_full.json'))
eq_pts = []
EQL = ('A-N-GYM EQUIPMENT','A-X-GYM EQUIP','A-N-SPA EQUIPMENT','A-X-SPA EQUIP')
segs_by_layer = collections.defaultdict(list)
for s in d['shapes']:
    lay = s[0]
    if len(s) >= 2 and s[1] is not None:
        x0,y0,x1,y1 = s[1]
        segs_by_layer[lay].append((x0,y0,x1,y1))
        if lay in EQL:
            eq_pts.append((x0,y0)); eq_pts.append((x1,y1))
    elif len(s) >= 3 and s[2] is not None:
        pts = list(zip(s[2][0::2], s[2][1::2]))
        for i in range(len(pts)-1):
            segs_by_layer[lay].append((pts[i][0],pts[i][1],pts[i+1][0],pts[i+1][1]))
        if lay in EQL:
            eq_pts.extend(pts[::2])
    elif len(s) >= 4 and s[3] is not None:
        cx,cy,r = s[3][0], s[3][1], s[3][2]
        if lay in EQL:
            eq_pts.append((s[3][3],s[3][4])); eq_pts.append((s[3][5],s[3][6]))

print("equip points:", len(eq_pts))
CELL = 0.5
cells = {}
for x,y in eq_pts:
    cells[(int(math.floor(x/CELL)), int(math.floor(y/CELL)))] = 0
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
clusters = []
for root, cs in groups.items():
    xs = [c[0] for c in cs]; ys = [c[1] for c in cs]
    x0,x1 = min(xs)*CELL, (max(xs)+1)*CELL
    y0,y1 = min(ys)*CELL, (max(ys)+1)*CELL
    w,h = x1-x0, y1-y0
    if w*h < 1.0: continue
    clusters.append({'bb':[round(x0,2),round(y0,2),round(x1,2),round(y1,2)],
                     'c':[round((x0+x1)/2,2),round((y0+y1)/2,2)], 'w':round(w,2),'h':round(h,2),
                     'ncell': len(cs)})
clusters.sort(key=lambda c:-c['ncell'])
print("equipment clusters:", len(clusters))
hist = collections.Counter((round(c['w']), round(c['h'])) for c in clusters)
print("size histogram (w x h):")
for k,v in hist.most_common(15): print("   ", k, v)

# ---------- walls ----------
wall_layers = [l for l in segs_by_layer if 'WALL' in l.upper() or 'DEMISING' in l.upper()]
wall_segs = []
for l in wall_layers: wall_segs += segs_by_layer[l]
print("wall segs:", len(wall_segs), "from", wall_layers)

def pt_seg_dist(px,py,x0,y0,x1,y1):
    dx,dy = x1-x0, y1-y0
    L2 = dx*dx+dy*dy
    if L2 == 0: return math.hypot(px-x0,py-y0), 0
    t = max(0, min(1, ((px-x0)*dx+(py-y0)*dy)/L2))
    return math.hypot(px-(x0+t*dx), py-(y0+t*dy)), math.degrees(math.atan2(dy,dx))%180

WCELL = 5.0
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
    for i,c in enumerate(clusters):
        x0,y0,x1,y1 = c['bb']
        dx = max(x0-px, 0, px-x1); dy = max(y0-py, 0, py-y1)
        dd = math.hypot(dx,dy)
        if dd < best[0]: best = (dd, i)
    return best

spaces = [{'number':s['number'],'name':s['name'],'x':s['loc'][0],'y':s['loc'][1]} for s in E2['spaces'] if s.get('loc')]
def near_space(px,py):
    best=(1e18,None)
    for s in spaces:
        dd=(px-s['x'])**2+(py-s['y'])**2
        if dd<best[0]: best=(dd,s['name'])
    return best[1]

for f in fixtures:
    wd, wang = near_wall(f['x'], f['y'])
    cd, ci = near_cluster(f['x'], f['y'])
    f['wall_d'] = round(wd,2) if wd < 10 else None
    f['wall_ang'] = round(wang) if (wang is not None and wd < 10) else None
    f['clus_d'] = round(cd,2) if cd < 25 else None
    f['clus_i'] = ci if cd < 25 else None
    f['space'] = near_space(f['x'], f['y'])

json.dump({'fixtures':fixtures,'circuits':circuits,'clusters':clusters,
           'spaces':spaces,'keynotes':keynotes,
           'ftags':ftags,'etags':etags,'wtags':wtags,'wires':wires,
           'text':E['text_plan'],'leaders':E['leaders'],
           'equipment':[e for e in E['equipment'] if e.get('loc') and e['loc'][2]>-50]},
          open('analysis/truth.json','w'))
print()
print("=== wall-dist by type (top) ===")
bt = collections.defaultdict(list)
for f in fixtures:
    if f['wall_d'] is not None: bt[f['typ']].append(f['wall_d'])
for k,v in sorted(bt.items(), key=lambda kv:-len(kv[1]))[:14]:
    print("  %-34s n=%3d wall_d med=%.2f" % (k[:34], len(v), statistics.median(v)))
print()
print("=== cluster-dist for equipment-group fixtures ===")
bg = collections.defaultdict(list)
for f in fixtures:
    if f['group'] and f['clus_d'] is not None: bg[f['group']].append(f['clus_d'])
for k,v in sorted(bg.items(), key=lambda kv:-len(kv[1]))[:12]:
    print("  %-36s n=%3d clus_d med=%.2f" % (k[:36], len(v), statistics.median(v)))
