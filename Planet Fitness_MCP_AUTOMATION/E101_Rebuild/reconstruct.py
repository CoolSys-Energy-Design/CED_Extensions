# PF E101 reconstruction engine v3 — full store, CAD-driven.
# Inputs: CAD geometry + space labels + site_inputs (mech schedule / coordination).
# Output: analysis/recreation.json  (devices, circuits, panels, keynotes; rationale per device)
import json, math, collections, statistics, re

C = json.load(open('current_model/cad_full.json'))
CX = json.load(open('current_model/cad_extra.json'))
E2 = json.load(open('current_model/extract2.json'))
SP = [{'number':s['number'],'name':s['name'],'x':s['loc'][0],'y':s['loc'][1]} for s in E2['spaces'] if s.get('loc')]

SITE = {
  'rtu_recepts': [(127.7,-291.8),(148.4,-291.8),(211.1,-241.1),(257.1,-276.1),(256.5,-290.5),
                  (181.6,-207.3),(106.8,-243.8),(68.1,-218.9),(49.7,-122.6),(77.4,-261.7),
                  (77.1,-288.9),(172.9,-293.3)],
  'rtus': [('RTU-1',30747,50),('RTU-2',25761,40),('RTU-3',14127,20),('RTU-4',18282,25),
           ('RTU-5',25761,40),('RTU-6',45705,70),('RTU-7',45705,70),('(E)RTU-8',29916,50)],
}

segs = collections.defaultdict(list)
polys = collections.defaultdict(list)
arcs = collections.defaultdict(list)
for s in list(C['shapes']) + list(CX['shapes']):
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

def cluster_points(pts, cell=0.5, min_area=0.0):
    cells = {}
    for x,y in pts:
        cells[(math.floor(x/cell), math.floor(y/cell))] = 0
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
        xs=[c[0] for c in cs]; ys=[c[1] for c in cs]
        x0,x1 = min(xs)*cell,(max(xs)+1)*cell
        y0,y1 = min(ys)*cell,(max(ys)+1)*cell
        if (x1-x0)*(y1-y0) < min_area: continue
        out.append({'bb':[x0,y0,x1,y1],'c':[(x0+x1)/2,(y0+y1)/2],'w':x1-x0,'h':y1-y0,'n':len(cs)})
    return out

def space_at(x,y):
    best=(1e18,None)
    for s in SP:
        dd=(x-s['x'])**2+(y-s['y'])**2
        if dd<best[0]: best=(dd,s)
    return best[1]

def sp_by(sub):
    return [s for s in SP if sub in (s['name'] or '').upper()]

WALLL = [l for l in segs if 'WALL' in l.upper() or 'DEMISING' in l.upper()]
wall_segs = []
for l in WALLL: wall_segs += segs[l]
def near_wall(px,py,maxd=8):
    best=(1e9,None,None)
    for (x0,y0,x1,y1) in wall_segs:
        if min(x0,x1)-maxd>px or max(x0,x1)+maxd<px or min(y0,y1)-maxd>py or max(y0,y1)+maxd<py: continue
        dx,dy=x1-x0,y1-y0; L2=dx*dx+dy*dy
        if L2==0: continue
        t=max(0,min(1,((px-x0)*dx+(py-y0)*dy)/L2))
        qx,qy=x0+t*dx,y0+t*dy
        d=math.hypot(px-qx,py-qy)
        if d<best[0]: best=(d,math.degrees(math.atan2(dy,dx))%180,(qx,qy))
    return best

eqgrid = collections.defaultdict(int)
for l in ('A-N-GYM EQUIPMENT','A-X-GYM EQUIP'):
    for (x0,y0,x1,y1) in segs[l]:
        eqgrid[(math.floor((x0+x1)/2/0.5), math.floor((y0+y1)/2/0.5))] += 1
def slot_hits(x,y):
    tot=0
    for dy2 in range(-14,15):
        for dxc in range(-3,4):
            tot += eqgrid.get((math.floor(x/0.5)+dxc, math.floor(y/0.5)+dy2),0)
    return tot
def eq_depth(x,y,dy,half=1.5,maxd=9.0):
    got=0.0; miss=0; step=0.5
    for k in range(1,int(maxd/step)+1):
        yy=y+dy*k*step
        hit=0
        xs=math.floor((x-half)/0.5); xe=math.floor((x+half)/0.5)
        for xx in range(xs,xe+1): hit+=eqgrid.get((xx,math.floor(yy/0.5)),0)
        if hit>0: got=k*step; miss=0
        else:
            miss+=1
            if miss>=3: break
    return got

devices=[]; circuits=[]; kn_list=[]
def dev(typ,x,y,rot,elev,load,va,why,fam='EF-U_Receptacle_CED',kn=None,cat='elec',driver='CAD'):
    devices.append({'fam':fam,'typ':typ,'x':round(x,3),'y':round(y,3),'rot':int(rot)%360,'elev':elev,
                    'load':load,'va':va,'why':why,'kn':kn,'cat':cat,'driver':driver,'panel':None,'ckt':None})
    if kn is not None:
        kn_list.append({'num':kn,'x':round(x+1.2,2),'y':round(y+1.2,2),'ref':len(devices)-1})
    return len(devices)-1

def ckt(panel,num,load,va,volts,poles,rating,members,why=''):
    for i in members:
        devices[i]['panel']=panel; devices[i]['ckt']=str(num)
    circuits.append({'panel':panel,'ckt':str(num),'load':load,'va':va,'volts':volts,
                     'poles':poles,'rating':rating,'members':members,'why':why})

# ================================================================ 1. CARDIO
boxes=[]
for pts in polys['A-N-RACEWAY']:
    x0,y0,x1,y1 = bbox(pts)
    if (x1-x0)<0.6 and (y1-y0)<0.6:
        boxes.append(((x0+x1)/2,(y0+y1)/2))
rowmap = collections.defaultdict(list)
for (x,y) in boxes:
    key=None
    for k in rowmap:
        if abs(k-y)<1.0: key=k
    rowmap[key if key is not None else y].append((x,y))
rows = sorted(rowmap.items(), key=lambda kv:-statistics.median(p[1] for p in kv[1]))

sections=[]
for pts in polys['A-N-RACEWAY']:
    b = bbox(pts)
    if (b[2]-b[0])>2 or (b[3]-b[1])>2: sections.append(b)
def section_edge_y(cx, cy, mdir):
    for b in sections:
        if b[0]-0.3<cx<b[2]+0.3 and abs((b[1]+b[3])/2-cy)<1.2:
            return b[1] if mdir<0 else b[3]
    return None
cardio=[]
for ri,(ry,pts) in enumerate(rows):
    pts.sort()
    for (x,y) in pts:
        up=eq_depth(x,y,+1); dn=eq_depth(x,y,-1)
        mdir = +1 if up>dn else -1
        ey = section_edge_y(x,y,mdir)
        if ey is not None: y = ey
        d = max(up,dn)
        if d>=6.5: mach,va='TREADMILL',1500
        elif slot_hits(x,y)>8000: mach,va='STAIRMASTER',1200
        else: mach,va='POWERED BIKE',400
        rot = 0 if mdir>0 else 180   # ground pin faces its machine (cord side)
        i=dev('Duplex Wall',x,y,rot,0.0,mach,va,
              'Outlet box drawn on the A-N-RACEWAY power raceway at this machine section: every cardio '
              'machine gets its own surface receptacle on the raceway, oriented toward the machine it '
              'serves. Equipment block behind the box is %.1f ft deep -> %s (1500/1200/400 VA per PF '
              'cutsheet).'%(d,mach))
        cardio.append((ri,x,i,mach))
odd,even = 1,2
for ri in sorted(set(c[0] for c in cardio)):
    rowd = sorted([c for c in cardio if c[0]==ri], key=lambda c:c[1])
    xs=[c[1] for c in rowd]
    gaps=[(xs[k+1]-xs[k],k) for k in range(len(xs)-1)]
    split = max(gaps)[1]+1 if gaps and max(gaps)[0]>5 else len(rowd)
    for c in rowd[:split]:
        ckt('L1',odd,c[3],devices[c[2]]['va'],120,1,20,[c[2]],
            'One dedicated 20A/1P circuit per machine; west row-segment fills the odd (left) breaker column of L1, west to east.')
        odd+=2
    for c in rowd[split:]:
        ckt('L1',even,c[3],devices[c[2]]['va'],120,1,20,[c[2]],
            'One dedicated 20A/1P circuit per machine; east row-segment fills the even (right) breaker column of L1, west to east.')
        even+=2
row_geo=[]
for ri,(ry,pts) in enumerate(rows):
    xs=[p[0] for p in pts]; y=statistics.median(p[1] for p in pts)
    row_geo.append((min(xs),max(xs),y))
    kn_list.append({'num':13,'x':round(min(xs)-2.5,2),'y':round(y,2),'ref':None})
    kn_list.append({'num':13,'x':round(max(xs)+2.5,2),'y':round(y,2),'ref':None})

# trench floor JBs at the outermost structural columns crossing each raceway row
wp=[]
for (x0,y0,x1,y1) in segs['A-WALL-PATT']:
    wp.append(((x0+x1)/2,(y0+y1)/2))
cols=[c for c in cluster_points(wp,cell=0.6) if c['w']<3 and c['h']<3]
if row_geo:
    rx0 = min(g[0] for g in row_geo); rx1 = max(g[1] for g in row_geo)
    ys = [g[2] for g in row_geo]
    band_cols = sorted(set(round(c['c'][0],1) for c in cols
                       if rx0-2<c['c'][0]<rx1+2 and min(ys)-8<c['c'][1]<max(ys)+8))
    merged=[]
    for x in band_cols:
        if merged and x-merged[-1]<4: continue
        merged.append(x)
    if len(merged)>=2:
        for cx in (merged[0], merged[-1]):
            for (x0r,x1r,yr) in row_geo:
                dev('Floor',cx+0.5,yr+0.25,0,0.0,None,0,
                    'Trench feed point: floor junction box where the structural column line crosses the '
                    'raceway row — power and data trench from the raceway end to the nearest column '
                    '(keynote 13).',fam='EF-U_Junction Box_CED',driver='CAD')

# ================================================================ 2. TVS
tvpts=[]
for pts in polys['A-N-TELEVISION']:
    for p in pts: tvpts.append(p)
for (x0,y0,x1,y1) in segs['A-N-TELEVISION']:
    tvpts.append((x0,y0)); tvpts.append((x1,y1))
tvc = cluster_points(tvpts, cell=0.8)
truss_bands=[c for c in tvc if c['w']>20]
wall_units=[]
rest=[c for c in tvc if c['w']<=20]
used=[False]*len(rest)
for i,a in enumerate(rest):
    if used[i]: continue
    grp=[a]; used[i]=True
    for j,b in enumerate(rest):
        if used[j]: continue
        if math.hypot(a['c'][0]-b['c'][0], a['c'][1]-b['c'][1])<4.6:
            grp.append(b); used[j]=True
    cx=sum(g['c'][0] for g in grp)/len(grp); cy=sum(g['c'][1] for g in grp)/len(grp)
    wall_units.append((cx,cy))

bankmem=[]; l2even=2
def flushbank():
    global bankmem,l2even
    if not bankmem: return
    ckt('L2',l2even,'TV TRUSS',360*len(bankmem),120,1,20,list(bankmem),
        'Truss TVs banked 3 per 20A circuit on L2 even numbers, west to east.')
    l2even+=2; bankmem=[]
# per-screen anchor: each TV hangs on a PAIR of mounting brackets (~0.8 ft boxes) drawn on the
# truss row of A-N-TELEVISION; the receptacle sits at the bracket-pair midpoint.
for band in truss_bands:
    bx0,by0,bx1,by1 = band['bb']
    brk=[]
    for pts2 in polys['A-N-TELEVISION']:
        b2=bbox(pts2)
        if not (bx0-1<b2[0] and b2[2]<bx1+1 and by0-1.5<b2[1] and b2[3]<by1+1.5): continue
        w2=b2[2]-b2[0]; h2=b2[3]-b2[1]
        if 0.3<w2<1.2 and h2<0.8:
            brk.append(b2)
    brk.sort(key=lambda b2:b2[0])
    screens=[]
    i2=0
    while i2 < len(brk)-1:
        a2,b2 = brk[i2], brk[i2+1]
        gap = b2[0]-a2[2]
        if 0.8<gap<2.4:
            screens.append(((a2[0]+a2[2]+b2[0]+b2[2])/4, (a2[1]+a2[3])/2))
            i2+=2
        else:
            i2+=1
    if len(screens)<4:      # fallback: uniform pitch across the band
        n = max(1, round((bx1-bx0)/5.02))
        screens=[(bx0+(bx1-bx0)*(k2+0.5)/n, (by0+by1)/2) for k2 in range(n)]
    for k2,(sx,sy) in enumerate(sorted(screens)):
        i=dev('Quad Wall - TV',sx,sy+0.28,90,0.0,'TV TRUSS',360,
              'Television #%d of %d on the gym TV truss: the CAD draws a pair of mounting brackets per '
              'screen on A-N-TELEVISION; the quad receptacle sits at the bracket-pair midpoint at the '
              'truss (height per keynote 11).'%(k2+1,len(screens)),kn=11 if k2==0 else None)
        bankmem.append(i)
        if len(bankmem)==3: flushbank()
    flushbank()
    dev('Ceiling',bx0-1.8,(by0+by1)/2,0,0.0,None,0,
        'Ceiling junction box at the west end of the TV truss: feed point for the truss circuits.',
        fam='EF-U_Junction Box_CED',driver='CAD')

walltv_ckts = {'WOMENS LOCKRM TV':('L3',32),'MENS LOCKRM TV':('L3',44),
               'TV - BLACK CARD SUITE - 103':('L3',14),'TV - FUNCTIONAL TRAINING':('L2',16)}
radiance=[]
for (x,y) in sorted(wall_units, key=lambda p:(-p[1],p[0])):
    s = space_at(x,y); sn=((s['name'] or '') if s else '').upper()
    wd,wang,q = near_wall(x,y)
    if q and wd<4: x,y = q
    rot = int(wang) if wang is not None else 0
    if 'LOCKER' in sn:
        load = ('WOMENS LOCKRM TV' if 'WOMENS' in sn else 'MENS LOCKRM TV')
    elif 'CHECK' in sn or 'RECEPTION' in sn or 'VESTIBULE' in sn:
        load = 'TV & RADIANCE MONITOR - CHECK-IN 102'
    elif 'SPA' in sn or 'BLACK' in sn or 'IT ROOM' in sn:
        load = 'TV - BLACK CARD SUITE - 103'
    elif 'TRAIN' in sn or 'MOBILITY' in sn:
        load = 'TV - FUNCTIONAL TRAINING'
    else:
        load = 'TV - GYM WALL'
    i=dev('Duplex Wall - TV',x,y,rot,0.0,load,300,
        'Single wall television on A-N-TELEVISION in %s: duplex TV receptacle behind the set, height per '
        'keynote 11.'%sn.title(),kn=11)
    if load=='TV & RADIANCE MONITOR - CHECK-IN 102':
        radiance.append(i)
    elif load in walltv_ckts:
        pn,cn = walltv_ckts[load]
        ckt(pn,cn,load,250,120,1,20,[i],'Wall TV on its own circuit in the locker/BCS band.')
if radiance:
    ckt('L3',7,'TV & RADIANCE MONITOR - CHECK-IN 102',250*len(radiance),120,1,20,radiance,
        'Check-in TV + radiance monitors grouped on L3/7.')

# ================================================================ 3. BCS SUITE
spa_pts=[]
for l in ('A-N-SPA EQUIPMENT','A-X-SPA EQUIP'):
    for (x0,y0,x1,y1) in segs[l]:
        spa_pts.append(((x0+x1)/2,(y0+y1)/2))
    for (cx,cy,r) in arcs[l]:
        spa_pts.append((cx,cy))
spac = cluster_points(spa_pts, cell=0.7, min_area=1.0)

l4_3p=[1,7,21,27,33,39,45]; i3=0
l4_2p=[2,6,10,14,17,18,22]; i2=0
def n3():
    global i3
    v = l4_3p[i3] if i3<len(l4_3p) else 45+6*(i3-len(l4_3p)+1)
    i3+=1; return '%d,%d,%d'%(v,v+2,v+4)
def n2():
    global i2
    v = l4_2p[i2] if i2<len(l4_2p) else 30+4*(i2-len(l4_2p))
    i2+=1; return '%d,%d'%(v,v+2)

# --- hydromassage: bed rows against the room's west wall
hyd = sp_by('HYDROMASSAGE')
hydro_touch=[]
if hyd:
    h = hyd[0]
    rows_y = sorted(set(round(c['c'][1],1) for c in spac
                    if abs(c['c'][0]-h['x'])<12 and -305<c['c'][1]<-274 and c['c'][0]<212))
    grp=[]
    for y in rows_y:
        if grp and y-grp[-1][-1]<2.5: grp[-1].append(y)
        else: grp.append([y])
    bed_ys = [statistics.mean(g) for g in grp]
    wallx = 197.4
    wd,wang,q = near_wall(h['x']-6, h['y'], maxd=12)
    if q: wallx = q[0]
    for by in bed_ys:
        i=dev('Specialty Wall - 240V/1Ph',wallx,by-0.4,90,0.0,'HYDROMASSAGE - %s'%h['number'],3840,
              'Hydromassage bed drawn on A-N-SPA EQUIPMENT: dedicated 240V/30A NEMA L6-30 at the bed head '
              'on the room\'s west wall (keynote 4).',kn=4)
        ckt('L4',n2(),'HYDROMASSAGE - %s'%h['number'],3840,240,2,30,[i],
            'Each hydromassage bed on its own 2-pole 30A circuit on high-leg panel L4.')
        j=dev('Duplex Wall',wallx,by+0.7,90,0.0,'HYDROMASSAGE RCPTS - %s'%h['number'],180,
              '120V receptacle for the bed touchscreen, just above each L6-30 outlet (keynote 4).')
        hydro_touch.append(j)
if hydro_touch:
    ckt('L3',12,'HYDROMASSAGE RCPTS - 103A',180*len(hydro_touch),120,1,20,hydro_touch,
        'All hydromassage touchscreen receptacles grouped on one L3 circuit.')

# --- tanning rooms: disconnects pair up on shared side walls at the SOUTH corner of each room
tan_rooms=[]
for s in SP:
    m = re.match(r'^103([B-F]|[K-L])$', s['number'])
    if not m: continue
    sn=(s['name'] or '').upper()
    if 'HYBRID' in sn: load,va,amp = 'HYBRID TANNER',11644,40
    elif 'RED WAVE' in sn: load,va,amp = 'RED WAVE',7067,30
    elif 'TANNING' in sn:
        if s['number'] in ('103K','103L'): load,va,amp = 'TANNING BED',11217,40
        else: load,va,amp = 'STAND-UP TANNER',10392,40
    else: continue
    tan_rooms.append((s,load,va,amp))
south_row=sorted([t for t in tan_rooms if t[0]['y']<-295], key=lambda t:t[0]['x'])
north_row=sorted([t for t in tan_rooms if t[0]['y']>=-295], key=lambda t:t[0]['x'])
def south_wall_y(s):
    best=None
    for (x0,y0,x1,y1) in wall_segs:
        if abs(y1-y0)>0.8: continue
        if min(x0,x1)-1<s['x']<max(x0,x1)+1 and y0<s['y']-1 and y0>s['y']-9:
            if best is None or y0>best: best=y0
    return best if best is not None else s['y']-5.7
def emit_bed(s,load,va,amp,hx):
    hy = south_wall_y(s)+0.3
    i=dev('Non-Fused - 60A',hx,hy,0,1.5,'%s - %s'%(load,s['number']),va,
          '%s in room %s: 240V/3ph/%dA dedicated circuit; NF-60A disconnect at the south corner of the '
          'shared side wall (adjacent rooms pair their disconnects so conduit drops share one stud bay), '
          '6 ft of #8 in flex for service moves (keynote 5).'%(load.title(),s['number'],amp),
          fam='EF-U_Disconnect Switch_CED',kn=5)
    ckt('L4',n3(),'%s - %s'%(load,s['number']),va,240,3,amp,[i],
        'Tanning-class bed: dedicated 3-pole circuit on the 240V-delta high-leg panel L4.')
for row in (north_row, south_row):
    if row is north_row and len(row)==2:
        a,b = row[0],row[1]
        emit_bed(a[0],a[1],a[2],a[3],a[0]['x']-5.3)   # west room -> its west wall (shares hydro drop)
        emit_bed(b[0],b[1],b[2],b[3],(a[0]['x']+b[0]['x'])/2+0.35)
        continue
    k=0
    while k < len(row):
        if k+1 < len(row):
            a,b = row[k],row[k+1]
            shared = (a[0]['x']+b[0]['x'])/2
            emit_bed(a[0],a[1],a[2],a[3],shared-0.35)
            emit_bed(b[0],b[1],b[2],b[3],shared+0.35)
            k+=2
        else:
            a=row[k]
            emit_bed(a[0],a[1],a[2],a[3],a[0]['x']+4.8)
            k+=1

# --- sauna + timer (west wall), spray tan (west wall)
for s in sp_by('SAUNA'):
    wd,wang,q = near_wall(s['x']-5, s['y']+0.5, maxd=5)
    hx,hy = q if q else (s['x']-5, s['y']+0.5)
    i=dev('Non-Fused - 60A',hx,hy-0.6,90,1.5,'SAUNA - %s'%s['number'],5760,
          'RedZone sauna room %s: 240V/1ph/40A dedicated circuit, NF disconnect at the unit with 6 ft of '
          '#10 in flex (keynote 22).'%s['number'],fam='EF-U_Disconnect Switch_CED',kn=22)
    ckt('L4',n2(),'SAUNA - %s'%s['number'],5760,240,2,40,[i],'Sauna heater 2-pole 40A on L4.')
    j=dev('Duplex Wall',hx,hy+0.6,90,0.0,'SAUNA TIMER CONTROL - %s'%s['number'],180,
          'Sauna timer control beside the sauna disconnect.')
    ckt('L3',10,'SAUNA TIMER CONTROL - %s'%s['number'],180,120,1,20,[j],'Timer control on L3.')
for s in sp_by('SPRAY TAN'):
    wd,wang,q = near_wall(s['x']+3.6, s['y']-4, maxd=5)
    hx,hy = q if q else (s['x']+3.6, s['y']-4)
    i=dev('Specialty Wall - 240V/1Ph',hx,hy,90,0.0,'SPRAY TAN - %s'%s['number'],5040,
          'Spray tanning booth room %s: 240V/1ph/30A NEMA L6-30; standalone 5mA GFCI module mounted in '
          'the IT room ahead of the receptacle (keynote 21).'%s['number'],kn=21)
    ckt('L4',n2(),'SPRAY TAN - %s'%s['number'],5040,240,2,30,[i],'Spray tan booth 2-pole 30A on L4.')

# --- cryolounge: chairs + polarwave on the east wall
cry = sp_by('CRYOLOUNGE')
if cry:
    s=cry[0]
    in_room=[c for c in spac if abs(c['c'][0]-s['x'])<8 and -305<c['c'][1]<-276]
    in_room.sort(key=lambda c:-c['c'][1])
    ewx = 269.5
    wd,wang,q = near_wall(s['x']+5, s['y'], maxd=6)
    if q: ewx = q[0]
    chairs=[c for c in in_room if c['c'][1]>-291]
    pols=[c for c in in_room if c['c'][1]<=-291]
    cknums=[26,28]
    for k,c in enumerate(chairs[:2]):
        i=dev('Duplex Wall',ewx,c['c'][1]+1.2,90,0.0,'CRYOLOUNGE',1440,
              'CryoLounge chair (A-N-SPA EQUIPMENT): 120V/20A dedicated receptacle on the east wall. '
              'Placed on L4 but only on A/C-phase positions (26/28) — the B phase of the 240V delta is '
              'the high leg and can never serve 120V loads.',kn=None)
        ckt('L4',cknums[k] if k<2 else 30,'CRYOLOUNGE',1440,120,1,20,[i],
            '120V load on L4 A/C-phase position only.')
    pn=[('18,20'),('22,24')]
    for k,c in enumerate(pols[:2]):
        i=dev('Non-Fused - 30A',ewx,c['c'][1]-2,90,1.5,'POLARWAVE',4800,
              'PolarWave cryotherapy unit: 240V 2-pole 20A dedicated circuit with NF-30A disconnect on '
              'the east wall.',fam='EF-U_Disconnect Switch_CED')
        ckt('L4',pn[k] if k<2 else '30,32','POLARWAVE',4800,240,2,20,[i],'PolarWave 2P circuit on L4.')

# --- massage chairs + hyperice in open spa
bcs = sp_by('BLACK CARD SPA')
if bcs:
    b=bcs[0]
    row=[c for c in spac if c['w']>8 and abs(c['c'][1]-(-271.6))<3 and 240<c['c'][0]<262]
    mem=[]
    if row:
        c=row[0]
        n=max(1,round(c['w']/3.5))
        for k in range(n):
            x=c['bb'][0]+c['w']*(k+0.5)/n
            i=dev('Duplex Floor',x,c['c'][1]-0.8,180,0.0,'MASSAGE CHAIRS - 103',500,
                  'Massage chair #%d drawn on A-N-SPA EQUIPMENT: floor receptacle below each chair, fed '
                  'through a floor trench to the nearest wall (keynote 20).'%(k+1),kn=20 if k==0 else None)
            mem.append(i)
    if mem:
        ckt('L3',6,'MASSAGE CHAIRS - 103',500*len(mem),120,1,20,mem,'All massage chairs on one L3 circuit.')
    hyp=[c for c in spac if 216<c['c'][0]<232 and -274<c['c'][1]<-266 and c['w']*c['h']>10]
    if hyp:
        c=hyp[0]
        wd,wang,q=near_wall(c['c'][0],c['c'][1]+4,maxd=6)
        hy = q[1] if q else -267.7
        h1=dev('Duplex Wall',c['c'][0]-3.9,hy,180,0.0,'HYPERICE EQUIP. & CHAIRS',180,
               'Hyperice recovery station on the spa north wall.')
        h2=dev('Duplex Wall',c['c'][0]-0.6,hy,180,0.0,'HYPERICE EQUIP. & CHAIRS',180,
               'Second hyperice receptacle.')
        ckt('L3',8,'HYPERICE EQUIP. & CHAIRS',360,120,1,20,[h1,h2],'Hyperice pair on one circuit.')

# --- BCS corridor maintenance ring
bcs_rails=[(-280.3,'N',16),(-296.4,'S',18)]
bmem={16:[],18:[]}
for (ry,tag,cknum) in bcs_rails:
    xs0,xs1 = 212.5, 256.5
    n = int((xs1-xs0)//10)
    for k in range(n):
        x = xs0 + (k+0.5)*(xs1-xs0)/n
        i=dev('Duplex Wall',x,ry,180 if tag=='S' else 0,0.0,'BCS MAINT. RCPTS',180,
              'Black-card corridor maintenance receptacle: duplex every ~10 ft along the %s corridor '
              'wall.'%('south' if tag=='S' else 'north'),kn=None)
        bmem[cknum].append(i)
for (px,py,ang,cknum) in ((214.8,-285.7,90,16),(253.6,-285.7,90,16),(254.0,-285.7,270,16),
                          (249.2,-280.1,0,16),(236.7,-295.9,0,18),(224.4,-296.4,180,18)):
    i=dev('Duplex Wall',px,py,ang,0.0,'BCS MAINT. RCPTS',180,
          'Corridor stub-wall maintenance receptacle at the tanning-room entries.',driver='JUDGMENT')
    bmem[cknum].append(i)
for k2 in (16,18):
    ckt('L3',k2,'BCS MAINT. RCPTS',180*len(bmem[k2]),120,1,20,bmem[k2],
        'Corridor ring split into two L3 circuits (16 north / 18 south).')

# ================================================================ 4. LOCKER/RESTROOMS
plumb=[]
for (x0,y0,x1,y1) in segs['A-N-PLUMB FIX']:
    plumb.append(((x0+x1)/2,(y0+y1)/2))
for (cx,cy,r) in arcs['A-N-PLUMB FIX']:
    plumb.append((cx,cy))
plc = cluster_points(plumb, cell=0.8, min_area=1.0)

for side,sink_x_range,ckts in (('MENS',(124,138),{'sink':34,'van':[36,42],'dry':[38,40],'maint':46}),
                               ('WOMENS',(144,158),{'sink':20,'van':[22,28],'dry':[24,26],'maint':30})):
    counters=[c for c in plc if sink_x_range[0]<c['c'][0]<sink_x_range[1] and -289<c['c'][1]<-283]
    if not counters: continue
    xs0=min(c['bb'][0] for c in counters); xs1=max(c['bb'][2] for c in counters)
    wy = statistics.mean(c['c'][1] for c in counters)
    mem=[]
    for k in range(3):
        x = xs0 + (xs1-xs0)*(k+0.5)/3
        i=dev('Duplex Wall - GFCI',x,wy+0.4,0,1.5,'%s SINK RCPTS'%side,180,
              'Trough-sink counter drawn on A-N-PLUMB FIX: GFCI below the counter for the automatic '
              'faucet sensors (keynote 3), one per sink pair.',kn=3 if k==0 else None)
        mem.append(i)
    ckt('L3',ckts['sink'],'%s SINK RCPTS'%side,540,120,1,20,mem,'Sink-sensor GFCIs on one circuit per restroom.')
    van=[]
    for k,x in enumerate((xs0-1.2, xs1+1.2)):
        i=dev('Duplex Wall - GFCI',x,wy-0.6,0,3.83,'%s VANITY RCPTS'%side,180,
              'Vanity GFCI above the counter at the end of the sink group (keynote 1).',kn=1 if k==0 else None)
        van.append(i)
    ckt('L3',ckts['van'][0],'%s VANITY RCPTS'%side,360,120,1,20,van,'Vanity pair at the sink counter.')
    van2=[]
    nx0 = 109.1 if side=='MENS' else 161.2
    ny = -268.7 if side=='MENS' else -269.8
    for k in range(2):
        i=dev('Duplex Wall - GFCI',nx0+3.1+3.1*k,ny,0,3.83,'%s VANITY RCPTS'%side,180,
              'Second vanity counter on the locker-room north millwork run: GFCI above counter (keynote 1).',
              driver='JUDGMENT')
        van2.append(i)
    ckt('L3',ckts['van'][1],'%s VANITY RCPTS'%side,360,120,1,20,van2,'Vanity pair at the locker-side counter.')
    dr=[]
    caps=[]
    for pts2 in polys['A-N-TOILET PARTITION']:
        b2=bbox(pts2); w2=b2[2]-b2[0]; h2=b2[3]-b2[1]
        if not (0.3<min(w2,h2)<0.8 and 0.7<max(w2,h2)<1.3): continue
        cx2=(b2[0]+b2[2])/2
        if sink_x_range[0]-6<cx2<sink_x_range[1]+6:
            caps.append((cx2,(b2[1]+b2[3])/2))
    caps.sort()
    for k,(cx2,cy2) in enumerate(caps[:2] if len(caps)>=2 else [(xs0-4.5,wy-1.6),(xs1+4.5,wy-1.6)]):
        i=dev('Wall - With Stem',cx2,cy2,180,1.5,'%s HAND DRYER'%side,1000,
              'Hand dryer mounted dead-center on the toilet-partition end cap: the CAD draws each stall '
              'divider cap as a ~0.6x1.0 ft rectangle on A-N-TOILET PARTITION straddling the wall, and '
              'the J-box (42 in AFF, keynote 2) lands exactly on its center. 1000 VA dedicated circuit.',
              fam='EF-U_Junction Box_CED',kn=2 if k==0 else None)
        dr.append(i)
    for k,i in enumerate(dr):
        ckt('L3',ckts['dry'][k],'%s HAND DRYER'%side,1000,120,1,20,[i],'Dedicated hand-dryer circuit.')
    mm=[]
    for k,x in enumerate((xs0-9,)):
        i=dev('Duplex Wall',x,wy-6,90,1.5,'%s MAINT. RCPTS'%side,180,
              'Locker-room maintenance receptacle near the restroom entry.')
        mm.append(i)
    i=dev('Duplex Wall',(xs0+xs1)/2,wy-11,0,1.5,'%s MAINT. RCPTS'%side,180,
          'Second locker-room maintenance receptacle.')
    mm.append(i)
    ckt('L3',ckts['maint'],'%s MAINT. RCPTS'%side,360,120,1,20,mm,'Locker maintenance pair per side.')

# exhaust fans at the toilet cores
fanmem=[]
for tnum,tx in (('105A',128.8),('104A',149.6)):
    ts=[s for s in SP if s['number']==tnum]
    if not ts: continue
    i=dev('Motor Rated Switch - 120V, 1 Pole',tx,-290.7,0,3.5,'LOCKER RM FANS EF-1 & 2',600,
          'Restroom exhaust fan over the toilet core: motor-rated switch at 3.5 ft.',
          fam='EF-U_Motor Rated Switch_CED')
    fanmem.append(i)
if fanmem:
    ckt('L2',27,'LOCKER RM FANS EF-1 & 2',1200,120,1,20,fanmem,'Both locker-room fans on one L2 circuit.')

# EWC
ewc=[c for c in plc if 136<c['c'][0]<146 and -272<c['c'][1]<-264]
if ewc:
    c=ewc[0]
    wd,wang,q=near_wall(*c['c'],maxd=4)
    hx,hy = q if q else c['c']
    i=dev('Duplex Wall - GFCI',hx,hy,int(wang or 0),0.0,'EWC',180,
          'Electric water cooler on A-N-PLUMB FIX at the locker corridor: GFCI mounted below the unit per '
          'the manufacturer template (keynote 12).',kn=12)
    ckt('L3',50,'EWC',180,120,1,20,[i],'Drinking fountain dedicated circuit.')

# ================================================================ 5. MAINT RECEPTACLES
GYM_SPACES=('CARDIO','STRENGTH','FREE WEIGHTS','FUNCTIONAL','CIRCUIT')
gym_ctrs=[(s['x'],s['y']) for s in SP if any(g in (s['name'] or '').upper() for g in GYM_SPACES)]
def in_gym(px,py,rad=62):
    s = space_at(px,py)
    if s is None: return False
    return any(g in (s['name'] or '').upper() for g in GYM_SPACES+('VESTIBULE','UTILITY','ELECTRICAL'))
maint_pts=[]
mseen=set()
def add_maint(px,py,ang,why):
    key=(round(px/7),round(py/7))
    if key in mseen: return
    mseen.add(key)
    maint_pts.append((px,py,ang,why))
for c in cols:
    px,py = c['c']
    if not in_gym(px,py): continue
    wd,wang,q = near_wall(px,py,maxd=2.5)
    if q:
        add_maint(q[0],q[1],int(wang),
            'Maintenance receptacle at a structural column embedded in this wall - the columns set the '
            '~37 ft bay rhythm the maintenance receptacles follow.')
    else:
        add_maint(px+0.2,py-1.1,180,
            'Maintenance receptacle on the free-standing structural column face - open gym floor has no '
            'walls, so the columns carry the maintenance receptacles.')
runs=[]
for (x0,y0,x1,y1) in wall_segs:
    L=math.hypot(x1-x0,y1-y0)
    if L<26: continue
    if not in_gym((x0+x1)/2,(y0+y1)/2): continue
    runs.append((x0,y0,x1,y1,L))
for (x0,y0,x1,y1,L) in sorted(runs,key=lambda r:-r[4]):
    n=max(1,int(round(L/33)))
    for k in range(n):
        t=(k+0.5)/n
        px,py=x0+(x1-x0)*t, y0+(y1-y0)*t
        ang=math.degrees(math.atan2(y1-y0,x1-x0))%180
        add_maint(px,py,int(ang),
            'Maintenance receptacle on a long gym wall run: one duplex per structural bay (~33 ft) so a '
            'cleaning cart can reach everywhere.')
maint_pts.sort(key=lambda m:(-m[1],m[0]))
mcks=[(18,'L2'),(20,'L2'),(22,'L2'),(24,'L2'),(26,'L2'),(28,'L2'),(39,'L3')]
mi=0
for k0 in range(0,len(maint_pts),4):
    grp=maint_pts[k0:k0+4]
    mem=[]
    for (px,py,ang,why) in grp:
        i=dev('Duplex Wall',px,py,ang,1.5,'GYM MAINT. RCPTS',180,why)
        mem.append(i)
    num,pn = mcks[mi] if mi<len(mcks) else (40+2*(mi-len(mcks)),'L3')
    ckt(pn,num,'GYM MAINT. RCPTS',180*len(mem),120,1,20,mem,
        'Maintenance receptacles grouped 3-5 per circuit by contiguous run (L2 evens 18-28, overflow L3/39+).')
    mi+=1
mem=[]
for (px,py,ang) in ((141.1,-273.4,90),(141.5,-278.9,90),(135.4,-278.9,90),(146.6,-278.9,90)):
    i=dev('Duplex Wall',px,py,ang,1.5,'LOCKER RM MAINT. RCPTS',180,
          'Locker-room corridor maintenance receptacle by the storage cores.',driver='JUDGMENT')
    mem.append(i)
ckt('L3',48,'LOCKER RM MAINT. RCPTS',720,120,1,20,mem,'Locker corridor ring on L3/48.')
rt=[]
for (px,py) in SITE['rtu_recepts']:
    i=dev('Duplex Wall - GFCI',px,py,90,1.5,'ROOFTOP MAINT. RCPTS',180,
          'GFCI service receptacle at a rooftop unit (RTU locations come from the mechanical plans; NEC '
          '210.63 requires a receptacle within 25 ft of roof HVAC).',driver='MECH')
    rt.append(i)
for k0,cknum in ((0,29),(4,31),(8,33)):
    mem=rt[k0:k0+4]
    if mem:
        ckt('L2',cknum,'ROOFTOP MAINT. RCPTS',180*len(mem),120,1,20,mem,
            'Rooftop receptacles grouped 4 per circuit on L2 odds 29-33.')

RTU_DEV=[]
fac=dev('Wall - With Stem',61.7,-213.3,90,1.5,'F.A.C.U',200,
    'Fire alarm control unit beside the electrical-room panel lineup: J-box with stem, own circuit.',
    fam='EF-U_Junction Box_CED',driver='RULE')
ckt('L2',37,'F.A.C.U',200,120,1,20,[fac],'FACU on L2/37.')
RTU_AT = {0:'RTU-6',1:'RTU-7',2:'RTU-1',3:'RTU-4',5:'(E)RTU-8',6:'RTU-5',7:'RTU-4',
          8:'RTU-5',9:'RTU-2',10:'RTU-3',11:'(E)RTU-8'}
rtu_pick = [(8,'RTU-5'),(9,'RTU-2'),(10,'RTU-3'),(11,'(E)RTU-8'),(7,'RTU-4'),(2,'RTU-1'),
            (0,'RTU-6'),(1,'RTU-7')]
for idx,nm in rtu_pick:
    px,py = SITE['rtu_recepts'][idx]
    rva = dict((n2,v2) for (n2,v2,a2) in SITE['rtus']).get(nm,0)
    i=dev('Non-Fused - 60A',px,py+1.75,90,1.5,nm,rva,
          'NF-60A disconnect at %s (unit location from the mechanical plans): local disconnecting means '
          'at each rooftop unit, fed from its MDP breaker.'%nm,fam='EF-U_Disconnect Switch_CED',driver='MECH')
    RTU_DEV.append((nm,i))

# ================================================================ 6. FRONT OF HOUSE
ck_sp = sp_by('CHECK-IN')
if ck_sp:
    cs=ck_sp[0]
    lwseg=[s2 for s2 in segs['A-N-WALL LOW'] if math.hypot((s2[0]+s2[2])/2-cs['x'],(s2[1]+s2[3])/2-cs['y'])<20]
    vert=[s2 for s2 in lwseg if abs(s2[0]-s2[2])<0.6 and abs(s2[1]-s2[3])>3]
    if vert:
        vx = max(s2[0] for s2 in vert)
        vy0 = min(min(s2[1],s2[3]) for s2 in vert); vy1 = max(max(s2[1],s2[3]) for s2 in vert)
    else:
        vx,vy0,vy1 = cs['x']-13.7, -253.9, -245.5
    mem=[]
    for k in range(4):
        y = vy0 + (k+0.5)*(vy1-vy0)/4
        i=dev('Duplex Wall - USB',vx+0.4,y,90,3.25,'CHECK-IN TABLETS',180,
              'Check-in desk (drawn as a low wall): USB receptacle at each tablet position along the '
              'counter, mounted horizontally above the counter surface (keynote 15).',kn=15 if k==0 else None)
        mem.append(i)
    for k,(dx,dy2) in enumerate(((1.9,0.1),(4.7,0.4))):
        i=dev('Duplex Wall - USB',vx+dx,vy1+dy2,0,3.25,'CHECK-IN TABLETS',180,
              'Tablet USB receptacle on the return leg of the check-in counter (keynote 15).')
        mem.append(i)
    ckt('L3',1,'CHECK-IN TABLETS',1080,120,1,20,mem,'All six tablet USBs on L3/1.')
    mem=[]
    comp=[]
    for pts2 in polys['A-N-FIXT-CHECK-IN']:
        b2=bbox(pts2); w2=b2[2]-b2[0]; h2=b2[3]-b2[1]
        if 1.3<w2<1.9 and 0.5<h2<0.9:
            comp.append(((b2[0]+b2[2])/2,(b2[1]+b2[3])/2))
    comp=sorted(set((round(cx2,2),round(cy2,2)) for cx2,cy2 in comp))
    ph=[((bbox(p)[0]+bbox(p)[2])/2,(bbox(p)[1]+bbox(p)[3])/2) for p in polys['A-N-FIXT-CHECK-IN']
        if 0.5<bbox(p)[2]-bbox(p)[0]<1.0 and 0.4<bbox(p)[3]-bbox(p)[1]<0.8]
    xs2 = sorted([c2[0] for c2 in comp] + [c2[0] for c2 in ph if -258<c2[1]<-256])[:3]
    for k,x2 in enumerate(xs2 if len(xs2)==3 else [cs['x']+6.1,cs['x']+8.35,cs['x']+10.6]):
        i=dev('Duplex Wall',x2,-256.31,0,0.0,'CHECK-IN DESK RCPTS',240,
              'PF_Plan computer / POS station drawn on A-N-FIXT-CHECK-IN: receptacle at each station '
              'along the check-in counter face; power and data trenched below the desk (keynote 7).',
              kn=7 if k==0 else None,driver='CAD')
        mem.append(i)
    ckt('L3',3,'CHECK-IN DESK RCPTS',720,120,1,20,mem,'Desk receptacles on L3/3.')
    dev('Wall - With Stem',cs['x']+5.0,cs['y']-2.1,0,1.5,None,0,
        'Check-in trench junction box: power+data stub below the desk (keynote 7).',
        fam='EF-U_Junction Box_CED',driver='JUDGMENT')
    t=dev('Duplex Wall',cs['x']+3.6,cs['y']-8.6,270,0.0,'TMAX RECEPT',180,
          'T-Max tanning-control receptacle + phone jack at the back wall of check-in (keynote 8): CAT-5 '
          'chain from here through each tanning room timer.',kn=8,driver='JUDGMENT')
    ckt('L3',9,'TMAX TIMER',180,120,1,20,[t],'T-Max controller on L3/9.')
    mem=[]
    for k,(dx,dy2) in enumerate(((6.1,-9.3),(11.7,-9.3),(14.3,-7.8),(16.0,-6.0))):
        i=dev('Duplex Wall',cs['x']+dx,cs['y']+dy2,0,0.0,'BACKWRAP RCPTS',180,
              'Backwrap merchandise wall behind check-in (keynote 10: coordinate exact locations with '
              'the tenant).',kn=10 if k==0 else None,driver='JUDGMENT')
        mem.append(i)
    ckt('L3',5,'BACKWRAP RCPTS',720,120,1,20,mem,'Backwrap receptacles on L3/5.')
    mem=[]
    for k,(px,py) in enumerate(((264.5,-264.1),(269.5,-256.5),(269.5,-248.0),(269.5,-214.4))):
        i=dev('Duplex Wall',px,py,90 if px>268 else 0,1.5,'RECEPTION MAINT. RCTPS',180,
              'Reception / storefront maintenance receptacle along the front wall and vestibule.',
              driver='JUDGMENT')
        mem.append(i)
    ckt('L3',11,'RECEPTION MAINT. RCTPS',720,120,1,20,mem,'Reception ring on L3/11.')

it = sp_by('IT ROOM')
if it:
    s=it[0]
    wy_s=None
    for (x0,y0,x1,y1) in wall_segs:
        if abs(y1-y0)>0.8: continue
        if min(x0,x1)-1<s['x']+2<max(x0,x1)+1 and y0<s['y']-2 and y0>s['y']-16:
            if wy_s is None or y0>wy_s: wy_s=y0
    if wy_s is None: wy_s = s['y']-9.6
    a=dev('Quad Wall',s['x']+2.2,wy_s+0.35,0,0.0,'I.T. RACK',360,
          'IT rack / TV distribution on the IT room south wall (keynote 9): quad receptacle #1, own '
          'circuit.',kn=9)
    b=dev('Quad Wall',s['x']+4.4,wy_s+0.35,0,0.0,'I.T. RACK',360,
          'IT rack quad #2 on a second dedicated circuit.')
    ckt('L3',2,'I.T. RACK',360,120,1,20,[a],'Rack circuit A.')
    ckt('L3',4,'I.T. RACK',360,120,1,20,[b],'Rack circuit B.')
    f1=dev('Motor Rated Switch - 120V, 1 Pole',s['x']-5.8,s['y']-6.6,0,3.5,'IT ROOM & BCS FANS',300,
           'IT room exhaust fan switch.',fam='EF-U_Motor Rated Switch_CED')
    f2=dev('Motor Rated Switch - 120V, 1 Pole',s['x']-6.2,s['y']-20.8,0,3.5,'IT ROOM & BCS FANS',300,
           'BCS corridor exhaust fan switch.',fam='EF-U_Motor Rated Switch_CED')
    ckt('L2',35,'IT ROOM & BCS FANS',600,120,1,20,[f1,f2],'Both fans on L2/35.')

# vending alcove: machines on the SPA layer against the 103L north wall
coolers=[]
for pts2 in polys['A-N-BEVERAGE COOLER']:
    b2=bbox(pts2); w2=b2[2]-b2[0]; h2=b2[3]-b2[1]
    if 2.0<w2<4.0 and 2.0<h2<4.0:
        coolers.append(b2)
coolers.sort(key=lambda b2:(round(b2[1],0), b2[0]))
vn=0
for b2 in coolers[:3]:
    cx2=(b2[0]+b2[2])/2
    north = (b2[1]+b2[3])/2 > -270.9   # machines north of the alcove wall back onto its south face
    ry = b2[1] if north else b2[3]
    rot = 0 if north else 180
    i=dev('Duplex Wall',cx2,ry,rot,0.0,'VENDING MACHINE',1000,
          'Beverage cooler drawn on A-N-BEVERAGE COOLER: dedicated receptacle centered on the back edge '
          'of each machine, on the wall it backs onto (cooler on a GFCI breaker per keynote 16).',
          kn=16 if vn==0 else None)
    ckt('L3',17+2*vn,'VENDING MACHINE',1000,120,1,20,[i],'Dedicated vending circuit, L3 odds 17/19/21.')
    vn+=1

# ================================================================ 7. BOH / MECH / SIGNS
brk = sp_by('BREAKROOM')
if brk:
    b=brk[0]
    mem=[]
    for k in range(5):
        i=dev('Duplex Wall',b['x']-9+4.2*k,b['y']+7.4,0,1.5,'BREAK & MECH RM RCPTS',180,
              'Break room general receptacles along the north wall (counter GFCI at the sink end).')
        mem.append(i)
    g=dev('Duplex Wall - GFCI',204.2,-267.7,0,1.5,'BREAK & MECH RM RCPTS',180,
          'Counter GFCI at the break-room sink end.',driver='JUDGMENT')
    mem.append(g)
    ckt('L3',23,'BREAK & MECH RM RCPTS',1080,120,1,20,mem,'Break/mech receptacles on L3/23.')
    hw=[c for c in plc if abs(c['c'][0]-198)<4 and abs(c['c'][1]+272)<4]
    hx,hy = (hw[0]['c'] if hw else (b['x']+6, b['y']+7.6))
    h=dev('Duplex Wall - GFCI',hx-0.4,hy-0.1,270,0.0,'HWH-1 - WATER HEATER',1440,
          'HWH-1 point-of-use water heater (A-N-PLUMB FIX): GFCI receptacle at the unit.')
    ckt('L3',41,'HWH-1 - WATER HEATER',1440,120,1,20,[h],'Water heater dedicated L3/41.')
    p=dev('Motor Rated Switch - 120V, 1 Pole',hx+1,hy-1.6,270,3.5,'P-1 - CIRC. PUMP',150,
          'P-1 circulation pump beside HWH-1: motor-rated switch at 3.5 ft.',
          fam='EF-U_Motor Rated Switch_CED')
    ckt('L3',43,'P-1 - CIRC. PUMP',150,120,1,20,[p],'Circ pump dedicated L3/43.')
    j=dev('Wall - With Stem',hx-0.2,hy-1.6,270,1.5,None,0,
          'Junction box with stem for the water-heater/pump connections.',fam='EF-U_Junction Box_CED')
for s in sp_by('STORAGE ROOM'):
    mem=[]
    for k in range(4):
        i=dev('Duplex Wall',s['x']-6+4*k,s['y']-5.5,180,1.5,'STORAGE ROOM RCPTS',180,
              'Storage room receptacle kit (4 per room).')
        mem.append(i)
    ckt('L3',25,'STORAGE ROOM RCPTS',720,120,1,20,mem,'Storage receptacles on L3/25.')
off_ck=[34,36]
for oi,s in enumerate(sp_by('OFFICE')):
    mem=[]
    for k in range(4):
        i=dev('Duplex Wall',s['x']-5+3.3*k,s['y']-4,0,1.5,'OFFICE - %s RCPTS'%s['number'],180,
              'Office %s receptacle kit (4 per office).'%s['number'])
        mem.append(i)
    ckt('L2',off_ck[oi] if oi<2 else 40,'OFFICE - %s RCPTS'%s['number'],720,120,1,20,mem,
        'Office receptacles on L2 evens 34/36.')
hu=[s for s in SP if 'HOUSE/UTILITY' in (s['name'] or '').upper()]
if hu:
    hx0,hy0 = hu[0]['x'],hu[0]['y']
    mem=[]
    for k in range(5):
        i=dev('Duplex Wall',hx0-20+9*k,hy0-14,0,1.5,'BOH - MAINT. RCPTS',180,
              'Back-of-house maintenance receptacles through the utility zone.')
        mem.append(i)
    ckt('L2',30,'BOH - MAINT. RCPTS',900,120,1,20,mem,'BOH ring 1 on L2/30.')
    mem=[]
    for k in range(4):
        i=dev('Duplex Wall',hx0-20+9*k,hy0+8,180,1.5,'BOH - MAINT. RCPTS',180,
              'Back-of-house maintenance receptacles, second run.')
        mem.append(i)
    ckt('L2',32,'BOH - MAINT. RCPTS',720,120,1,20,mem,'BOH ring 2 on L2/32.')
    mem=[]
    for k in range(6):
        ang=[90,90,0,270,270,270][k]
        px=[-11.8,-11.8,-25.1,-38.9,-38.9,-38.9][k]
        py=[-183.2,-208.2,-231.8,-220.7,-195.7,-170.7][k]
        i=dev('Duplex Wall',px,py,ang,1.5,'MEZZ RCPTS',180,
              'Mezzanine receptacles spaced around the west annex walls (coordination-driven).',
              driver='JUDGMENT')
        mem.append(i)
    ckt('L2',38,'MEZZ RCPTS',1080,120,1,20,mem,'Mezzanine receptacles on L2/38.')

# big fans: fan symbol marker groups on A-N-FAN; disconnect at the hub with the
# project-calibrated offset (cardio.md: +4 in Y standard, X offset varies per project CAD block)
fanm=[]
for pts2 in polys['A-N-FAN']:
    b2=bbox(pts2); fanm.append(((b2[0]+b2[2])/2,(b2[1]+b2[3])/2))
fg=[]
for (mx,my) in fanm:
    hit=None
    for g in fg:
        gx=sum(m[0] for m in g)/len(g); gy=sum(m[1] for m in g)/len(g)
        if abs(gx-mx)<12 and abs(gy-my)<12: hit=g
    if hit: hit.append((mx,my))
    else: fg.append([(mx,my)])
FAN_OFF=(-1.30,0.52)   # calibrated for this background's fan block (block reference off hub)
bfi=0
for g in sorted(fg,key=lambda g:min(m[0] for m in g)):
    if len(g)<4: continue
    cx2=sum(m[0] for m in g)/len(g); cy2=sum(m[1] for m in g)/len(g)
    i=dev('Non-Fused - 30A',cx2+FAN_OFF[0],cy2+FAN_OFF[1],180,0.0,'BIG FAN',1500,
          'Big Ass Fan symbol group on A-N-FAN: NF-30A disconnect at the fan hub (mid-room, straight '
          'conduit drop; hub offset from the block centroid is calibrated per background). Keynote 19.',
          fam='EF-U_Disconnect Switch_CED',kn=19,driver='CAD')
    ck0=33 if bfi==0 else 27
    ckt('L3','%d,%d,%d'%(ck0,ck0+2,ck0+4),'BIG FAN',1500,208,3,20,[i],'Big fan 3-pole group on L3 odds.')
    bfi+=1

vest=sp_by('VESTIBULE')
if vest:
    v=vest[0]
    e=dev('Ceiling',v['x'],v['y']+2.6,0,8.0,'ECH-1',3994,
          'ECH-1 electric cabinet heater above the vestibule: ceiling junction box at 8 ft, 208V 2-pole.',
          fam='EF-U_Junction Box_CED')
    ckt('L3','13,15','ECH-1',3994,208,2,20,[e],'Vestibule heater on L3 13,15.')
    sg=dev('Ceiling',v['x']+5.2,v['y']+1.8,0,8.0,'STORE SIGN',500,
          'Storefront sign: ceiling J-box behind the facade with a local disconnecting means (keynote 17).',
          fam='EF-U_Junction Box_CED',kn=17)
    ckt('L3',45,'STORE SIGN',500,120,1,20,[sg],'Sign circuit L3/45.')
strg=sp_by('STRENGTH')
if strg:
    s=strg[0]
    sgn=[bbox(p) for p in polys['A-N-SIGNAGE'] if 2<bbox(p)[2]-bbox(p)[0]<5]
    if sgn:
        scx=(sgn[0][0]+sgn[0][2])/2; scy=(sgn[0][1]+sgn[0][3])/2-0.5
    else:
        scx,scy = s['x']+3, s['y']+15.6
    pc=dev('Ceiling',scx,scy,0,8.0,'PF CLOCK',180,
          'Planet Fitness clock drawn on A-N-SIGNAGE on the north gym wall: ceiling J-box just below '
          'the sign band (keynote 18).',fam='EF-U_Junction Box_CED',kn=18,driver='CAD')
    ckt('L2',25,'PF CLOCK',180,120,1,20,[pc],'Clock on L2/25.')
    fs=dev('Duplex Wall',s['x']+35.6,s['y']+16.2,180,1.5,'FUTURE SIGN',180,
          'Stub receptacle for a future interior sign on the north gym wall.',driver='JUDGMENT')
    ckt('L2',23,'FUTURE SIGN',180,120,1,20,[fs],'Future sign stub on L2/23.')

# ================================================================ 8. PANELS
panels=[]
def panel(name,typ,x,y,why,fam='EE-U_Panelboard_CED'):
    panels.append({'name':name,'fam':fam,'typ':typ,'x':round(x,2),'y':round(y,2),'why':why})
elx=[]
for l in ('A-X-ELECTRICAL','A-N-ELECTRICAL'):
    for (x0,y0,x1,y1) in segs[l]:
        elx.append(((x0+x1)/2,(y0+y1)/2))
elc = cluster_points(elx, cell=1.2, min_area=1.0)
util_marks=[c for c in elc if c['c'][0]<80]
if util_marks:
    ux = statistics.mean(c['c'][0] for c in util_marks)
    uy = statistics.mean(c['c'][1] for c in util_marks)
else:
    ux,uy = 61.5,-206
panel('MDP','Switchboard - 480V/3Ph, I-Line Type',ux-0.7,uy+4.6,
      '480V service switchboard, first in the utility-room lineup marked on the A-X-ELECTRICAL layer.',
      fam='EE-U_Switchboard_CED')
panel('TR-L1','Dry-Type - 150 KVA',ux,uy+0.6,
      '150 kVA 480-208Y/120 dry transformer feeding L1, second in the lineup.',fam='EE-U_Transformer_CED')
panel('L1','Panelboard - 208V/3Ph, 600A',ux,uy-2.3,
      '600A 208Y/120 cardio panel L1, third in the lineup.')
panel('L2','Panelboard - 208V/3Ph, 125A',ux,uy-4.4,
      '125A lighting/building panel L2 beside L1, fed from an L1 bottom subfeed.')
if it:
    s=it[0]
    itm=[c for c in elc if math.hypot(c['c'][0]-s['x'],c['c'][1]-s['y'])<16]
    if itm:
        ix=statistics.mean(c['c'][0] for c in itm); iy=statistics.mean(c['c'][1] for c in itm)
    else:
        ix,iy = s['x']-4.4, s['y']+0.9
    panel('L3','Panelboard - 208V/3Ph, 125A',ix,iy+3.3,
          '125A front-of-house panel L3 in the IT room lineup (fed from L1 subfeed).')
    panel('L4','L4_Panelboard - 240V/3Ph',ix,iy+0.9,
          '240V-delta high-leg spa/tanning panel L4, beside L3.')
    panel('TR-L4','Dry-Type - 150 KVA - 240V',ix+0.1,iy-2.0,
          '150 kVA 480-240delta transformer feeding L4.',fam='EE-U_Transformer_CED')
    panel('SW-L4','Fused - 200A',ix+0.1,iy-4.8,
          '200A fused switch ahead of TR-L4.',fam='EE-U_Equipment Switch_CED')
ckt('L1','79,81,83','PANEL L2 SUB FEED BREAKER',30814,208,3,150,[],'Bottom-of-panel subfeed to L2 (odd side).')
ckt('L1','80,82,84','PANEL L3 SUB FEED BREAKER',33834,208,3,150,[],'Bottom-of-panel subfeed to L3 (even side).')
ckt('MDP','1','PANEL L1 VIA 150KVA XFMR (TR-L1)',140648,480,3,225,[],'Feeder to TR-L1/L1.')
ckt('MDP','2','PANEL L4 VIA 150KVA XFMR (TR-L4)',112213,480,3,200,[],'Feeder to TR-L4/L4.')
for k,(nm,va,amp) in enumerate(SITE['rtus']):
    mem=[i for (nm2,i) in RTU_DEV if nm2==nm]
    ckt('MDP',str(3+k),nm,va,480,3,amp,mem,'Rooftop unit breaker per mech schedule MCA; NF disconnect at the unit.')

json.dump({'devices':devices,'circuits':circuits,'panels':panels,'keynotes':kn_list},
          open('analysis/recreation.json','w'))
print("devices:", len(devices), "circuits:", len(circuits), "panels:", len(panels), "keynotes:", len(kn_list))
print(dict(collections.Counter(d['typ'] for d in devices)))
print(dict(collections.Counter((d['load'] or '').split(' - ')[0] for d in devices)))
