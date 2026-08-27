import json, os, re, collections, statistics
TRAIN = r"C:/CED_Extensions/Planet Fitness_MCP_AUTOMATION/PF2_Training"
U = 10.7639104
projects = [d for d in os.listdir(TRAIN) if os.path.isdir(os.path.join(TRAIN,d))]

def norm_load(nm):
    if not nm: return None
    n = re.sub(r'\s+',' ',nm.upper().replace('\r',' ').replace('\n',' ')).strip()
    n = re.sub(r'#?\d+$','',n).strip(' -#')
    n = re.sub(r'- \d+[A-Z]?$','',n).strip(' -')
    return n

def norm_space(nm):
    if not nm: return None
    n = re.sub(r'\s+',' ', nm.upper()).strip()
    n = re.sub(r'\s*\d+[A-Z]?$','',n).strip()
    return n

agg_load = collections.defaultdict(lambda: {"cnt":0,"va":[],"volts":[],"poles":[],"rating":[],"nfix":[],"panels":collections.Counter(),"famtypes":collections.Counter(),"ckt_odd":0,"ckt_even":0,"projects":set()})
famtype_elev = collections.defaultdict(list)
space_fixtures = collections.defaultdict(collections.Counter)
space_loads = collections.defaultdict(collections.Counter)
panel_roles = collections.defaultdict(collections.Counter)
proj_ok = []

for p in sorted(projects):
    fp = os.path.join(TRAIN,p,'extract.json')
    if not os.path.exists(fp): continue
    try:
        d = json.load(open(fp))
    except Exception as e:
        print("SKIP",p,e); continue
    proj_ok.append(p)
    fixById = {f['id']:f for f in d.get('fixtures',[])}
    spctrs = [(norm_space(sp.get('name')), sp['loc'][0], sp['loc'][1]) for sp in d.get('spaces',[]) if sp.get('loc')]
    def space_of(loc):
        if not loc or loc[2] < -50 or not spctrs: return None
        best, bd = None, 1e18
        for nm,x,y in spctrs:
            dd = (loc[0]-x)**2 + (loc[1]-y)**2
            if dd < bd: bd, best = dd, nm
        return best if bd < 90**2 else None
    for s in d.get('systems',[]):
        ln = norm_load(s.get('load_name'))
        if not ln or ln in ('SPARE','SPACE','HIGH LEG SPACE'): continue
        members = [fixById.get(m) for m in s.get('members',[]) if m in fixById]
        a = agg_load[ln]
        a['cnt'] += 1; a['projects'].add(p)
        if s.get('app_load_va'): a['va'].append(round(s['app_load_va']/U,1))
        if s.get('volts'): a['volts'].append(round(s['volts']/U))
        if s.get('poles'): a['poles'].append(s['poles'])
        if s.get('rating'): a['rating'].append(s['rating'])
        a['nfix'].append(len(s.get('members',[])))
        pn = (s.get('panel') or '?').strip()
        a['panels'][pn]+=1
        panel_roles[pn][ln]+=1
        ck = s.get('circuit') or ''
        m = re.match(r'^(\d+)', ck)
        if m:
            if int(m.group(1))%2: a['ckt_odd']+=1
            else: a['ckt_even']+=1
        for f in members:
            if not f: continue
            ft = f['family']+' :: '+f['type']
            a['famtypes'][ft]+=1
            sn = space_of(f.get('loc'))
            if sn:
                space_fixtures[sn][ft]+=1
                space_loads[sn][ln]+=1
    for f in d.get('fixtures',[]):
        if f.get('loc') and f['loc'][2] < -50: continue
        ft = f['family']+' :: '+f['type']
        el = f.get('params',{}).get('Elevation from Level')
        if el is not None:
            famtype_elev[ft].append(round(el,2))

def med(l): return round(statistics.median(l),1) if l else None
out = {}
for ln,a in sorted(agg_load.items(), key=lambda kv:-kv[1]['cnt']):
    out[ln] = {
      "circuits_total": a['cnt'], "nproj": len(a['projects']),
      "va_med": med(a['va']), "volts_med": med(a['volts']),
      "poles_mode": collections.Counter(a['poles']).most_common(1)[0][0] if a['poles'] else None,
      "rating_mode": collections.Counter(a['rating']).most_common(1)[0][0] if a['rating'] else None,
      "fix_per_ckt_med": med(a['nfix']),
      "panels": dict(a['panels'].most_common(4)),
      "famtypes": dict(a['famtypes'].most_common(4)),
      "odd_even": [a['ckt_odd'],a['ckt_even']],
    }
json.dump(out, open('analysis/load_rules.json','w'), indent=1)
json.dump({k:dict(v.most_common(20)) for k,v in space_fixtures.items()}, open('analysis/space_fixtures.json','w'), indent=1)
json.dump({k:dict(v.most_common(25)) for k,v in space_loads.items()}, open('analysis/space_loads.json','w'), indent=1)
json.dump({k:dict(v.most_common(40)) for k,v in panel_roles.items() if sum(v.values())>=3}, open('analysis/panel_roles.json','w'), indent=1)
json.dump({k:{"n":len(v),"med":med(v),"modes":collections.Counter(v).most_common(3)} for k,v in famtype_elev.items()}, open('analysis/famtype_elev.json','w'), indent=1)
print("projects:", len(proj_ok), "| loads:", len(out), "| space types:", len(space_fixtures))
print()
print("=== TOP 60 LOAD NAMES (cnt, nproj, VA, V, poles, rating, fix/ckt, panels) ===")
for ln,o in list(out.items())[:60]:
    print(f"  {ln[:38]:38s} n={o['circuits_total']:3d} p={o['nproj']:2d} va={o['va_med']} v={o['volts_med']} {o['poles_mode']}P {o['rating_mode']}A fpc={o['fix_per_ckt_med']} panels={o['panels']}")
