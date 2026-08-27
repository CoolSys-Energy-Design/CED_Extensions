# Assemble the single JS data bundle the artifact embeds.
import json, math, re, collections

CAD = json.load(open('current_model/cad_render.json'))
T = json.load(open('analysis/truth.json'))
R = json.load(open('analysis/recreation.json'))
XR = json.load(open('analysis/exact_report.json'))
RULES = open('analysis/RULES.md', encoding='utf-8').read()
INJ = open('injection/inject_e101.py', encoding='utf-8').read()
MAN = json.load(open('injection/manifest.json'))

# ---- truth compact ----
tdev = []
for f in T['fixtures']:
    tdev.append({
        'id': f['id'], 'cat': f['cat'], 'fam': f['fam'], 'typ': f['typ'],
        'x': f['x'], 'y': f['y'], 'rot': f['rot'], 'elev': f['elev'],
        'load': f.get('load'), 'panel': f.get('panel'), 'ckt': f.get('ckt'),
        'va': f.get('va'), 'group': f.get('group'), 'space': f.get('space'),
    })
tpan = []
for e in T['equipment']:
    tpan.append({'name': e['params'].get('Panel Name') or e['type'],
                 'fam': e['family'], 'typ': e['type'],
                 'x': round(e['loc'][0],2), 'y': round(e['loc'][1],2)})
twire = [{'v': [[round(p[0],2), round(p[1],2)] for p in w['verts'] if p]} for w in T['wires']]
ttag = []
for t in T['ftags'] + T['etags'] + T['wtags']:
    if t.get('head'):
        ttag.append({'x': round(t['head'][0],2), 'y': round(t['head'][1],2),
                     'txt': (t.get('text') or '').replace('\r','/').replace('\n','/')[:40],
                     'host': (t.get('host') or [None])[0]})
tkn = [{'num': k.get('num'), 'x': k['loc'][0], 'y': k['loc'][1]} for k in T['keynotes'] if k.get('loc')]
tckt = [{'panel': c['panel'], 'ckt': c['ckt'], 'load': c['load'], 'va': c['va'],
         'volts': c['volts'], 'poles': c['poles'], 'rating': c['rating'],
         'members': c['members']} for c in T['circuits']]

# ---- recreation compact ----
rdev = []
for i, d in enumerate(R['devices']):
    rdev.append({'i': i, 'fam': d['fam'], 'typ': d['typ'], 'x': d['x'], 'y': d['y'],
                 'rot': d['rot'], 'elev': d['elev'], 'load': d['load'], 'va': d['va'],
                 'panel': d['panel'], 'ckt': d['ckt'], 'kn': d.get('kn'),
                 'driver': d.get('driver'), 'why': d.get('why'),
                 'rr': d.get('rule_resid_in'), 'snapped': d.get('snapped', False),
                 'coord': d.get('coord_added', False), 'cal': d.get('cal')})
rckt = [{'panel': c['panel'], 'ckt': c['ckt'], 'load': c['load'], 'va': c['va'],
         'volts': c['volts'], 'poles': c['poles'], 'rating': c['rating'],
         'members': c['members'], 'why': c.get('why','')} for c in R['circuits']]

# ---- matching (truth <-> recreation) for diff view ----
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
    if 'SPRAY TAN' in n or ('SAUNA' in n) or 'POLARWAVE' in n or 'CRYO' in n: return 'spa'
    if 'MASSAGE CHAIR' in n or 'HYPERICE' in n: return 'spa'
    if 'BCS MAINT' in n: return 'bcsmaint'
    if 'HAND DRYER' in n: return 'rr'
    if 'SINK' in n or 'VANITY' in n or 'EWC' in n: return 'rr'
    if 'FANS' in n or n=='F.A.C.U': return 'mech'
    if 'ROOFTOP' in n: return 'roof'
    if n.startswith('RTU') or n.startswith('(E)RTU'): return 'roof'
    if 'MEZZ' in n: return 'boh'
    if 'MAINT' in n: return 'maint'
    if 'TABLET' in n or 'DESK' in n or 'BACKWRAP' in n or 'TMAX' in n: return 'foh'
    if 'I.T.' in n or 'IT RACK' in n: return 'foh'
    if 'VENDING' in n: return 'foh'
    if 'BREAK' in n or 'STORAGE' in n or 'OFFICE' in n: return 'boh'
    if 'BIG FAN' in n or 'ECH' in n or 'HWH' in n or 'P-1' in n: return 'mech'
    if 'SIGN' in n or 'CLOCK' in n: return 'sign'
    return 'other'
tk = collections.defaultdict(list); rk = collections.defaultdict(list)
for j,f in enumerate(tdev):
    if f['cat']=='elec': tk[klass(f.get('load'), f['typ'])].append(j)
for j,d in enumerate(rdev):
    rk[klass(d.get('load'), d['typ'])].append(j)
pairs=[]; t_un=[]; r_un=set(range(len(rdev)))
for k in set(list(tk)+list(rk)):
    tl,rl = tk.get(k,[]), rk.get(k,[])
    used=set()
    for j in tl:
        f=tdev[j]
        best=(1e9,None)
        for j2 in rl:
            if j2 in used: continue
            d=rdev[j2]
            dd=math.hypot(f['x']-d['x'], f['y']-d['y'])
            if dd<best[0]: best=(dd,j2)
        if best[1] is not None and best[0]<=15:
            used.add(best[1]); r_un.discard(best[1])
            pairs.append({'t': j, 'r': best[1], 'd': round(best[0],2)})
        else:
            t_un.append(j)
score = XR

bundle = {
 'cad': CAD,
 'spaces': T['spaces'],
 'truth': {'devices': tdev, 'panels': tpan, 'wires': twire, 'tags': ttag,
           'keynotes': tkn, 'circuits': tckt},
 'recr': {'devices': rdev, 'panels': R['panels'], 'circuits': rckt, 'keynotes': R['keynotes']},
 'match': {'pairs': pairs, 'truth_unmatched': t_un, 'recr_unmatched': sorted(r_un)},
 'score': score,
 'rules_md': RULES,
 'injector_py': INJ,
 'manifest_head': {'meta': MAN['meta'],
                   'devices_n': len(MAN['devices']), 'circuits_n': len(MAN['circuits']),
                   'sample_device': MAN['devices'][0], 'sample_circuit': MAN['circuits'][0]},
}
js = 'window.PF = ' + json.dumps(bundle, separators=(',',':')) + ';'
open('artifact_data.js','w',encoding='utf-8').write(js)
import os
print('artifact_data.js bytes:', os.path.getsize('artifact_data.js'))
print('pairs:', len(pairs), 'truth unmatched:', len(t_un), 'recr unmatched:', len(r_un))
