# Inch-level residual diagnosis: for every truth device, nearest recr device of same class,
# print per-class stats: median |d| inches, mean dx, mean dy (systematic offset), worst cases.
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

tf = [f for f in T['fixtures'] if f['cat']=='elec']
rf = R['devices']
tk = collections.defaultdict(list); rk = collections.defaultdict(list)
for j,f in enumerate(tf): tk[klass(f.get('load'), f['typ'])].append(j)
for j,d in enumerate(rf): rk[klass(d.get('load'), d['typ'])].append(j)

print("%-11s %5s %5s | med_in  mean_dx_in mean_dy_in  max_in | rot_ok elev_ok typ_ok" % ('class','truth','recr'))
rows = {}
for k in sorted(set(list(tk)+list(rk))):
    tl,rl = tk.get(k,[]), rk.get(k,[])
    used=set(); ds=[]; dxs=[]; dys=[]; rot_ok=0; el_ok=0; ty_ok=0; worst=[]
    for j in tl:
        f=tf[j]; best=(1e9,None)
        for j2 in rl:
            if j2 in used: continue
            d=rf[j2]
            dd=math.hypot(f['x']-d['x'], f['y']-d['y'])
            if dd<best[0]: best=(dd,j2)
        if best[1] is None or best[0]>20:
            worst.append(('MISS', f['x'], f['y'], None)); continue
        used.add(best[1]); d=rf[best[1]]
        ds.append(best[0]*12); dxs.append((d['x']-f['x'])*12); dys.append((d['y']-f['y'])*12)
        if (f.get('rot') or 0)==(d.get('rot') or 0): rot_ok+=1
        if abs((f.get('elev') or 0)-(d.get('elev') or 0))<0.05: el_ok+=1
        if f['typ']==d['typ']: ty_ok+=1
        worst.append((round(best[0]*12,1), round(f['x'],2), round(f['y'],2), round(d['x'],2), round(d['y'],2), (f.get('load') or '')[:20]))
    if not tl and not rl: continue
    med = statistics.median(ds) if ds else -1
    mdx = statistics.mean(dxs) if dxs else 0
    mdy = statistics.mean(dys) if dys else 0
    mx = max(ds) if ds else -1
    print("%-11s %5d %5d | %6.1f  %8.1f %9.1f  %6.1f | %d/%d %d/%d %d/%d" %
          (k, len(tl), len(rl), med, mdx, mdy, mx, rot_ok, len(ds), el_ok, len(ds), ty_ok, len(ds)))
    worst.sort(key=lambda w: -(w[0] if isinstance(w[0],(int,float)) else 1e9))
    rows[k] = worst
json.dump({k:v[:40] for k,v in rows.items()}, open('analysis/residuals.json','w'), indent=1)
print()
print("=== worst members per problem class ===")
for k in ('cardio','tv-truss','gymmaint','boh','office','break','storage','bcsmaint','desk','backwrap','vanity','dryer','sink','recmaint','lockmaint','sidemaint','vending','tv-wall','bed','cryo','chairs','hyperice'):
    if k not in rows: continue
    print(k)
    for w in rows[k][:6]:
        print("   ", w)
