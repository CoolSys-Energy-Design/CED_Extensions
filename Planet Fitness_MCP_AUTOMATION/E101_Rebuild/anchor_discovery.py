# Anchor discovery: for each truth device class, test candidate CAD feature sets and report
# which feature type explains positions to sub-inch, with the fitted constant offset.
import json, math, collections, statistics, re

C = json.load(open('current_model/cad_full.json'))
T = json.load(open('analysis/truth.json'))

segs = collections.defaultdict(list)
polys = collections.defaultdict(list)
arcs = collections.defaultdict(list)
for s in C['shapes']:
    lay = s[0]
    if len(s) >= 2 and s[1] is not None:
        segs[lay].append(tuple(s[1]))
    elif len(s) >= 3 and s[2] is not None:
        pts = list(zip(s[2][0::2], s[2][1::2]))
        polys[lay].append(pts)
        for i in range(len(pts)-1):
            segs[lay].append((pts[i][0],pts[i][1],pts[i+1][0],pts[i+1][1]))
    elif len(s) >= 4 and s[3] is not None:
        arcs[lay].append((s[3][0], s[3][1], s[3][2]))

def bbox(pts):
    xs=[p[0] for p in pts]; ys=[p[1] for p in pts]
    return min(xs),min(ys),max(xs),max(ys)

# ---------- candidate feature sets (each: list of (x,y,tagstr)) ----------
feat = collections.defaultdict(list)

# raceway: section south/north edges at box-center x
boxes=[]; sections=[]
for pts in polys['A-N-RACEWAY']:
    b=bbox(pts)
    if (b[2]-b[0])<0.6 and (b[3]-b[1])<0.6: boxes.append(b)
    elif (b[2]-b[0])>2: sections.append(b)
for b in boxes:
    cx=(b[0]+b[2])/2
    sec=[s for s in sections if s[0]-0.2<cx<s[2]+0.2 and abs((s[1]+s[3])/2-(b[1]+b[3])/2)<1]
    if sec:
        s0=sec[0]
        feat['race_s_edge'].append((cx, s0[1], 'sec s'))
        feat['race_n_edge'].append((cx, s0[3], 'sec n'))
        feat['race_box_c'].append((cx, (b[1]+b[3])/2, 'box c'))
        feat['race_box_s'].append((cx, b[1], 'box s'))

# TV brackets: 0.8-wide boxes on the truss row -> pair midpoints
tvsm=[]
for pts in polys['A-N-TELEVISION']:
    b=bbox(pts)
    if 0.3<(b[2]-b[0])<1.2 and -223.0<b[1]<-222.0:
        tvsm.append(b)
tvsm.sort(key=lambda b:b[0])
for i in range(len(tvsm)-1):
    a,b2 = tvsm[i], tvsm[i+1]
    gap = b2[0]-a[2]
    if 1.0<gap<2.2:
        mx=(a[0]+a[2]+b2[0]+b2[2])/4
        feat['tv_pairmid'].append((mx, (a[1]+a[3])/2, 'brkt pair'))
        feat['tv_pairmid_top'].append((mx, max(a[3],b2[3]), 'brkt top'))
# all TELEVISION cluster boxes (for wall TVs)
for pts in polys['A-N-TELEVISION']:
    b=bbox(pts)
    feat['tv_any'].append(((b[0]+b[2])/2,(b[1]+b[3])/2,'tv box'))

# wall vertices + faces
WALLL=[l for l in segs if 'WALL' in l.upper() or 'DEMISING' in l.upper()]
wall_segs=[]
for l in WALLL: wall_segs+=segs[l]
for (x0,y0,x1,y1) in wall_segs:
    feat['wall_vert'].append((x0,y0,'wv')); feat['wall_vert'].append((x1,y1,'wv'))

# columns
wp=[]
for (x0,y0,x1,y1) in segs['A-WALL-PATT']:
    wp.append(((x0+x1)/2,(y0+y1)/2))
def cluster(pts,cell=0.6,min_area=0.0):
    cells={}
    for x,y in pts: cells[(math.floor(x/cell),math.floor(y/cell))]=0
    parent={}
    def find(a):
        while parent[a]!=a: parent[a]=parent[parent[a]]; a=parent[a]
        return a
    def union(a,b):
        ra,rb=find(a),find(b)
        if ra!=rb: parent[ra]=rb
    for c2 in cells: parent[c2]=c2
    for (i,j) in cells:
        for di in(-1,0,1):
            for dj in(-1,0,1):
                if (i+di,j+dj) in cells: union((i,j),(i+di,j+dj))
    gr=collections.defaultdict(list)
    for c2 in cells: gr[find(c2)].append(c2)
    out=[]
    for cs in gr.values():
        xs=[c2[0] for c2 in cs]; ys=[c2[1] for c2 in cs]
        x0,x1=min(xs)*cell,(max(xs)+1)*cell; y0,y1=min(ys)*cell,(max(ys)+1)*cell
        if (x1-x0)*(y1-y0)<min_area: continue
        out.append({'bb':[x0,y0,x1,y1],'c':[(x0+x1)/2,(y0+y1)/2],'w':x1-x0,'h':y1-y0})
    return out
for c2 in cluster(wp):
    if c2['w']<3 and c2['h']<3:
        feat['column'].append((c2['c'][0],c2['c'][1],'col'))

# doors: cluster A-N-DOOR + frames
dp=[]
for l in ('A-N-DOOR','A-N-Door Frame','A-N-DOOR-HEAD'):
    for (x0,y0,x1,y1) in segs[l]:
        dp.append(((x0+x1)/2,(y0+y1)/2))
    for (cx,cy,r) in arcs.get(l,[]):
        dp.append((cx,cy))
for c2 in cluster(dp, cell=0.8):
    feat['door'].append((c2['c'][0],c2['c'][1],'door'))
    feat['door_edge'].append((c2['bb'][0],c2['c'][1],'door w'))
    feat['door_edge'].append((c2['bb'][2],c2['c'][1],'door e'))
    feat['door_edge'].append((c2['c'][0],c2['bb'][1],'door s'))
    feat['door_edge'].append((c2['c'][0],c2['bb'][3],'door n'))

# plumbing fine clusters (basins/faucets)
pp=[]
for (x0,y0,x1,y1) in segs['A-N-PLUMB FIX']:
    pp.append(((x0+x1)/2,(y0+y1)/2))
for c2 in cluster(pp, cell=0.4):
    feat['plumb_fine'].append((c2['c'][0],c2['c'][1],'pl %0.1fx%0.1f'%(c2['w'],c2['h'])))
for (cx,cy,r) in arcs['A-N-PLUMB FIX']:
    feat['plumb_arc'].append((cx,cy,'arc r%.1f'%r))

# spa fine clusters
sp=[]
for l in ('A-N-SPA EQUIPMENT','A-X-SPA EQUIP'):
    for (x0,y0,x1,y1) in segs[l]:
        sp.append(((x0+x1)/2,(y0+y1)/2))
    for (cx,cy,r) in arcs[l]:
        sp.append((cx,cy))
for c2 in cluster(sp, cell=0.5):
    feat['spa_fine'].append((c2['c'][0],c2['c'][1],'spa %0.1fx%0.1f'%(c2['w'],c2['h'])))
    feat['spa_edges'].append((c2['bb'][0],c2['c'][1],'spa wedge'))
    feat['spa_edges'].append((c2['bb'][2],c2['c'][1],'spa eedge'))
    feat['spa_edges'].append((c2['c'][0],c2['bb'][1],'spa sedge'))
    feat['spa_edges'].append((c2['c'][0],c2['bb'][3],'spa nedge'))

# millwork fine
wd=[]
for (x0,y0,x1,y1) in segs['A-FLOR-WDWK']:
    wd.append(((x0+x1)/2,(y0+y1)/2))
for c2 in cluster(wd, cell=0.5):
    feat['wdwk_fine'].append((c2['c'][0],c2['c'][1],'wd %0.1fx%0.1f'%(c2['w'],c2['h'])))

# gym equip fine (for vending etc.)
ge=[]
for l in ('A-N-GYM EQUIPMENT',):
    for (x0,y0,x1,y1) in segs[l]:
        m=((x0+x1)/2,(y0+y1)/2)
        if -280<m[1]<-240 and 180<m[0]<275:
            ge.append(m)
for c2 in cluster(ge, cell=0.5):
    if c2['w']*c2['h']>1:
        feat['equip_lobby'].append((c2['c'][0],c2['c'][1],'eq %0.1fx%0.1f'%(c2['w'],c2['h'])))

# electrical marks
el=[]
for l in ('A-N-ELECTRICAL','A-X-ELECTRICAL'):
    for (x0,y0,x1,y1) in segs[l]:
        el.append(((x0+x1)/2,(y0+y1)/2))
for c2 in cluster(el, cell=0.8):
    feat['elec_mark'].append((c2['c'][0],c2['c'][1],'em'))

# low wall verts
for (x0,y0,x1,y1) in segs['A-N-WALL LOW']+segs.get('A-X-LOW WALL',[]):
    feat['lowwall_v'].append((x0,y0,'lw')); feat['lowwall_v'].append((x1,y1,'lw'))

# ---------- classify truth (same as residuals.py) ----------
def normload(l):
    if not l: return ''
    n = re.sub(r'\s+',' ', l.upper()).strip()
    n = re.sub(r'\s*-\s*1?\d\d[A-Z]?$','',n)
    return n.replace('RCTPS','RCPTS').replace('SPARY','SPRAY')
def klass(load, typ):
    n = normload(load)
    if n in ('TREADMILL','STAIRMASTER','POWERED BIKE','STEPMILL'): return 'cardio'
    if 'TV TRUSS' in n: return 'tv-truss'
    if n.startswith('TV') or 'LOCKRM TV' in n or 'RADIANCE' in n: return 'tv-wall'
    if 'HYDROMASSAGE' in n and 'RCPT' not in n: return 'hydro240'
    if 'HYDROMASSAGE' in n: return 'hydro120'
    if any(k in n for k in ('TANNING BED','HYBRID TANNER','STAND-UP','RED WAVE','TLT')): return 'bed'
    if 'SPRAY TAN' in n: return 'spray'
    if 'SAUNA' in n: return 'sauna'
    if 'POLARWAVE' in n or 'CRYO' in n: return 'cryo'
    if 'MASSAGE CHAIR' in n: return 'chairs'
    if 'HYPERICE' in n: return 'hyperice'
    if 'BCS MAINT' in n: return 'bcsmaint'
    if 'HAND DRYER' in n: return 'dryer'
    if 'SINK' in n: return 'sink'
    if 'VANITY' in n: return 'vanity'
    if 'EWC' in n: return 'ewc'
    if 'FANS' in n: return 'fans'
    if n=='F.A.C.U': return 'facu'
    if 'ROOFTOP' in n: return 'roofmaint'
    if n.startswith('RTU') or n.startswith('(E)RTU'): return 'rtudisc'
    if 'MEZZ' in n: return 'mezz'
    if 'LOCKER RM MAINT' in n: return 'lockmaint'
    if 'MAINT' in n and ('MENS' in n or 'WOMENS' in n): return 'sidemaint'
    if 'RECEPTION MAINT' in n: return 'recmaint'
    if 'MAINT' in n: return 'gymmaint'
    if 'TABLET' in n: return 'tablets'
    if 'DESK' in n: return 'desk'
    if 'BACKWRAP' in n: return 'backwrap'
    if 'TMAX' in n: return 'tmax'
    if 'I.T.' in n or 'IT RACK' in n: return 'itrack'
    if 'VENDING' in n: return 'vending'
    if 'BREAK' in n: return 'break'
    if 'STORAGE' in n: return 'storage'
    if 'OFFICE' in n: return 'office'
    if 'BOH' in n: return 'boh'
    if 'BIG FAN' in n: return 'bigfan'
    if 'ECH' in n or 'HWH' in n or 'P-1' in n: return 'mechpt'
    if 'SIGN' in n or 'CLOCK' in n: return 'sign'
    return 'other'

tf=[f for f in T['fixtures'] if f['cat']=='elec']
byk=collections.defaultdict(list)
for f in tf: byk[klass(f.get('load'),f['typ'])].append(f)

FOCUS = {
 'cardio': ['race_s_edge','race_n_edge','race_box_c','race_box_s'],
 'tv-truss': ['tv_pairmid','tv_pairmid_top'],
 'tv-wall': ['tv_any'],
 'sink': ['plumb_fine','plumb_arc'],
 'vanity': ['plumb_fine','wdwk_fine','plumb_arc'],
 'dryer': ['plumb_fine','wall_vert','door_edge'],
 'bed': ['wall_vert','spa_edges','door_edge','spa_fine'],
 'sauna': ['wall_vert','spa_edges','door_edge'],
 'spray': ['spa_edges','wall_vert'],
 'cryo': ['spa_fine','spa_edges','wall_vert'],
 'hydro240': ['spa_fine','spa_edges','wall_vert'],
 'hydro120': ['spa_fine','spa_edges'],
 'chairs': ['spa_fine','spa_edges'],
 'hyperice': ['spa_fine','spa_edges','wall_vert'],
 'bcsmaint': ['door','door_edge','wall_vert','column'],
 'gymmaint': ['column','wall_vert','door_edge'],
 'lockmaint': ['wall_vert','door_edge','column','plumb_fine'],
 'sidemaint': ['wall_vert','door_edge','plumb_fine','wdwk_fine'],
 'recmaint': ['wall_vert','door_edge','column','lowwall_v'],
 'tablets': ['lowwall_v','wdwk_fine'],
 'vending': ['equip_lobby','spa_fine','wall_vert'],
 'itrack': ['elec_mark','wall_vert'],
 'boh': ['wall_vert','door_edge','column'],
 'office': ['wall_vert','door_edge'],
 'break': ['wall_vert','door_edge','plumb_fine'],
 'storage': ['wall_vert','door_edge'],
 'bigfan': ['column','equip_lobby'],
 'other': ['column','wall_vert','elec_mark'],
}

print("Per class: candidate feature -> median |resid| in, p90, and the modal (dx,dy) offset")
for k in sorted(byk):
    fl=byk[k]
    cands = FOCUS.get(k, ['wall_vert','door_edge','column'])
    print("%s  (n=%d)" % (k, len(fl)))
    for cn in cands:
        pts=feat.get(cn,[])
        if not pts: continue
        res=[]; offs=[]
        for f in fl:
            best=(1e9,None)
            for (x,y,tg) in pts:
                dd=math.hypot(f['x']-x, f['y']-y)
                if dd<best[0]: best=(dd,(x,y,tg))
            if best[1] is None: continue
            res.append(best[0]*12)
            offs.append((round((f['x']-best[1][0])*12,1), round((f['y']-best[1][1])*12,1)))
        if not res: continue
        res_s=sorted(res)
        med=statistics.median(res_s); p90=res_s[int(len(res_s)*0.9)-1] if len(res_s)>1 else res_s[0]
        offc=collections.Counter((round(o[0]/2)*2, round(o[1]/2)*2) for o in offs)
        print("   %-14s med=%6.1f in  p90=%6.1f in  modal_off=%s" % (cn, med, p90, offc.most_common(2)))
