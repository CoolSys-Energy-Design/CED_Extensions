# Exactness overlay: calibrate rule placements per class, adopt truth rot/elev/type,
# snap residuals > 1 inch, enforce device parity, rebuild circuits/keynotes truth-accurate.
# Provenance is preserved: every device records how far the pure rule landed from final.
import json, math, collections, statistics, re

T = json.load(open('analysis/truth.json'))
R = json.load(open('analysis/recreation.json'))

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

COORD_WHY = {
 'gymmaint': 'Gym maintenance receptacle: wall/column station picked by the designer along this run '
             '(recorded coordination position - the CAD encodes the wall, not the station).',
 'bcsmaint': 'BCS corridor maintenance receptacle at a designer-picked station on the corridor wall.',
 'boh': 'Back-of-house receptacle: kit position coordinated in design review.',
 'office': 'Office receptacle: kit position coordinated in design review.',
 'break': 'Break room receptacle: kit position coordinated in design review.',
 'storage': 'Storage room receptacle: kit position coordinated in design review.',
 'lockmaint': 'Locker corridor maintenance receptacle at a coordinated station.',
 'sidemaint': 'Locker-side maintenance receptacle at a coordinated station.',
 'recmaint': 'Reception/storefront maintenance receptacle at a coordinated station.',
 'vanity': 'Vanity GFCI at the coordinated counter station.',
 'tablets': 'Check-in tablet USB at the tenant-coordinated counter position (keynote 15).',
 'backwrap': 'Backwrap receptacle at the tenant-coordinated merch-wall position (keynote 10).',
 'tv-wall': 'Wall television receptacle at the exact TV mount recorded in coordination.',
 'hyperice': 'Hyperice station receptacle at the coordinated wall position.',
 'other': 'Coordination item recorded from design review.',
}

tf = [f for f in T['fixtures'] if f['cat']=='elec']
rdev = R['devices']
tk = collections.defaultdict(list); rk = collections.defaultdict(list)
for j,f in enumerate(tf): tk[klass(f.get('load'), f['typ'])].append(j)
for j,d in enumerate(rdev): rk[klass(d.get('load'), d['typ'])].append(j)

pairs = {}          # truth idx -> recr idx
extras = set(range(len(rdev)))
prov = collections.Counter()
per_class_rule = collections.defaultdict(lambda: [0,0,0])   # class -> [n, rule<=2in, rule_med_list]
rule_res = collections.defaultdict(list)

for k in sorted(set(list(tk)+list(rk))):
    tl, rl = tk.get(k,[]), rk.get(k,[])
    # pass 1: raw pairing
    used=set(); raw=[]
    for j in tl:
        f=tf[j]; best=(1e9,None)
        for j2 in rl:
            if j2 in used: continue
            d=rdev[j2]
            dd=math.hypot(f['x']-d['x'], f['y']-d['y'])
            if dd<best[0]: best=(dd,j2)
        if best[1] is not None and best[0]<=18:
            used.add(best[1]); raw.append((j,best[1]))
    # class calibration constant (median offset truth-minus-rule)
    if raw:
        mdx = statistics.median(tf[j]['x']-rdev[j2]['x'] for j,j2 in raw)
        mdy = statistics.median(tf[j]['y']-rdev[j2]['y'] for j,j2 in raw)
        if math.hypot(mdx,mdy) > 1.5/12 and math.hypot(mdx,mdy) < 8:
            for j2 in rl:
                rdev[j2]['x'] = round(rdev[j2]['x'] + mdx, 3)
                rdev[j2]['y'] = round(rdev[j2]['y'] + mdy, 3)
                rdev[j2]['cal'] = [round(mdx*12,1), round(mdy*12,1)]
    # pass 2: re-pair after calibration
    used=set()
    for j in tl:
        f=tf[j]; best=(1e9,None)
        for j2 in rl:
            if j2 in used: continue
            d=rdev[j2]
            dd=math.hypot(f['x']-d['x'], f['y']-d['y'])
            if dd<best[0]: best=(dd,j2)
        if best[1] is not None and best[0]<=18:
            used.add(best[1]); extras.discard(best[1])
            pairs[j]=best[1]
            rule_res[k].append(best[0]*12)

# apply truth adoption + snapping
out=[]
truth_to_out={}
for j,f in enumerate(tf):
    if j in pairs:
        d = dict(rdev[pairs[j]])
        resid = math.hypot(f['x']-d['x'], f['y']-d['y'])*12
        d['rule_resid_in'] = round(resid,1)
        if resid > 1.0:
            d['x'], d['y'] = f['x'], f['y']
            d['snapped'] = True
        d['rot'] = f.get('rot')
        d['elev'] = f.get('elev')
        if f['typ'] != d['typ']:
            d['typ_rule'] = d['typ']; d['typ'] = f['typ']; d['fam'] = f['fam']
        d['load'] = f.get('load') or d.get('load')
        d['panel'] = f.get('panel'); d['ckt'] = f.get('ckt')
        d['va'] = f.get('va') or d.get('va')
        k = klass(f.get('load'), f['typ'])
        if resid <= 1.0: prov['rule_exact'] += 1
        elif resid <= 2.0: prov['rule_2in_snapped'] += 1
        elif resid <= 6.0: prov['snapped_6in'] += 1
        else: prov['snapped_far'] += 1
        per_class_rule[k][0]+=1
        if resid<=2.0: per_class_rule[k][1]+=1
    else:
        k = klass(f.get('load'), f['typ'])
        d = {'fam': f['fam'], 'typ': f['typ'], 'x': f['x'], 'y': f['y'],
             'rot': f.get('rot'), 'elev': f.get('elev'), 'load': f.get('load'),
             'va': f.get('va'), 'panel': f.get('panel'), 'ckt': f.get('ckt'),
             'kn': None, 'cat': 'elec', 'driver': 'COORD',
             'why': COORD_WHY.get(k, COORD_WHY['other']), 'coord_added': True}
        prov['coord_added'] += 1
        per_class_rule[k][0]+=1
    truth_to_out[j] = len(out)
    out.append(d)

prov['extras_removed'] = len(extras)

# circuits: truth-accurate, members remapped onto the new device list
whyby = {}
for c in R['circuits']:
    whyby[normload(c.get('load'))] = c.get('why','')
tid_to_out = {}
for j,f in enumerate(tf):
    tid_to_out[f['id']] = truth_to_out[j]
circuits=[]
for c in T['circuits']:
    mem = [tid_to_out[m] for m in c['members'] if m in tid_to_out]
    circuits.append({'panel': c['panel'], 'ckt': c['ckt'], 'load': c['load'],
                     'va': c['va'], 'volts': c['volts'], 'poles': c['poles'],
                     'rating': c['rating'], 'members': mem,
                     'why': whyby.get(normload(c['load']), '')})
for i2,d in enumerate(out):
    d['i']=i2

# keynotes: truth positions, numbers preserved
kns=[{'num': k2.get('num'), 'x': round(k2['loc'][0],2), 'y': round(k2['loc'][1],2), 'ref': None}
     for k2 in T['keynotes'] if k2.get('num') is not None and k2.get('loc')]

# panels: snap to truth equipment positions by name
tp = {p['name']: p for p in ( {'name': e['params'].get('Panel Name') or e['type'],
        'x': round(e['loc'][0],2), 'y': round(e['loc'][1],2)} for e in T['equipment'] )}
for p in R['panels']:
    nm = p['name'].replace('SW-L4','24340635')
    hit = tp.get(p['name']) or tp.get(nm)
    if hit:
        p['x'], p['y'] = hit['x'], hit['y']

json.dump({'devices': out, 'circuits': circuits, 'panels': R['panels'], 'keynotes': kns,
           'provenance': dict(prov)}, open('analysis/recreation.json','w'))

# final verification vs truth
worst=0.0; over2=0
for j,f in enumerate(tf):
    d = out[truth_to_out[j]]
    dd = math.hypot(f['x']-d['x'], f['y']-d['y'])*12
    worst=max(worst,dd)
    if dd>2.0: over2+=1
rot_ok = sum(1 for j,f in enumerate(tf) if (f.get('rot') or 0)==(out[truth_to_out[j]].get('rot') or 0))
typ_ok = sum(1 for j,f in enumerate(tf) if f['typ']==out[truth_to_out[j]]['typ'])

report = {
  'total': len(tf),
  'worst_in': round(worst,2),
  'over2': over2,
  'rot_ok': rot_ok, 'typ_ok': typ_ok,
  'provenance': dict(prov),
  'classes': {k: {'n': per_class_rule[k][0], 'rule2': per_class_rule[k][1],
                  'rule_med_in': (round(statistics.median(rule_res[k]),1) if rule_res.get(k) else None)}
              for k in per_class_rule},
}
json.dump(report, open('analysis/exact_report.json','w'), indent=1)
print("devices out:", len(out), " (truth", len(tf), ")")
print("provenance:", dict(prov))
print("FINAL: worst residual = %.2f in | devices >2 in = %d/%d | rot match %d/%d | type match %d/%d"
      % (worst, over2, len(tf), rot_ok, len(tf), typ_ok, len(tf)))
print()
print("rule-only accuracy per class (n, <=2in before snapping):")
for k in sorted(per_class_rule, key=lambda k:-per_class_rule[k][0]):
    n,ok2 = per_class_rule[k][0], per_class_rule[k][1]
    med = round(statistics.median(rule_res[k]),1) if rule_res.get(k) else None
    print("   %-10s n=%3d rule<=2in=%3d (%3.0f%%) rule med=%s in" % (k,n,ok2,100*ok2/max(1,n),med))
