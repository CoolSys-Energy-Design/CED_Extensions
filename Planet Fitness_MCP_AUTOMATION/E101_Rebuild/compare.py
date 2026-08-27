# Score recreation.json against truth.json class by class.
import json, math, collections, re

T = json.load(open('analysis/truth.json'))
R = json.load(open('analysis/recreation.json'))

def normload(l):
    if not l: return ''
    n = re.sub(r'\s+',' ', l.upper()).strip()
    n = re.sub(r'\s*-\s*1?\d\d[A-Z]?$','',n)   # strip room suffixes
    n = n.replace('RCTPS','RCPTS').replace('SPARY','SPRAY')
    return n

def klass(load, typ):
    n = normload(load)
    if n in ('TREADMILL','STAIRMASTER','POWERED BIKE','STEPMILL'): return 'cardio'
    if 'TV TRUSS' in n: return 'tv-truss'
    if n.startswith('TV') or 'LOCKRM TV' in n or 'RADIANCE' in n: return 'tv-wall'
    if 'HYDROMASSAGE' in n and 'RCPT' not in n: return 'bcs-hydro-240'
    if 'HYDROMASSAGE' in n: return 'bcs-hydro-120'
    if any(k in n for k in ('TANNING BED','HYBRID TANNER','STAND-UP','RED WAVE','TLT')): return 'bcs-bed'
    if 'SPRAY TAN' in n: return 'bcs-spray'
    if 'SAUNA' in n and 'TIMER' not in n: return 'bcs-sauna'
    if 'SAUNA TIMER' in n: return 'bcs-sauna-timer'
    if 'POLARWAVE' in n or 'CRYO' in n: return 'bcs-cryo'
    if 'MASSAGE CHAIR' in n: return 'bcs-chairs'
    if 'HYPERICE' in n: return 'bcs-chairs'
    if 'BCS MAINT' in n or n=='RECEPT - BCS': return 'bcs-maint'
    if 'HAND DRYER' in n: return 'rr-dryer'
    if 'SINK' in n: return 'rr-sink'
    if 'VANITY' in n: return 'rr-vanity'
    if 'EWC' in n or 'DRINK' in n: return 'rr-ewc'
    if 'LOCKER RM FANS' in n or ('FANS' in n and 'EF' in n) or 'BCS FANS' in n: return 'fans'
    if 'ROOFTOP' in n or 'ROOF MAINT' in n: return 'roof-maint'
    if 'MEZZ' in n: return 'mezz'
    if 'MAINT' in n and ('GYM' in n or 'BOH' in n or 'LOCKER' in n or 'RECEPTION' in n or 'MENS' in n or 'WOMENS' in n): return 'maint'
    if 'TABLET' in n: return 'foh-tablets'
    if 'DESK' in n or 'BACKWRAP' in n or 'TMAX' in n or 'T-MAX' in n: return 'foh-desk'
    if 'I.T. RACK' in n or 'IT RACK' in n: return 'foh-it'
    if 'VENDING' in n: return 'vending'
    if 'BREAK' in n or 'STORAGE' in n or 'OFFICE' in n: return 'boh-rooms'
    if 'BIG FAN' in n: return 'bigfan'
    if 'ECH' in n: return 'mech'
    if 'HWH' in n or 'CIRC' in n or 'P-1' in n: return 'mech'
    if 'SIGN' in n: return 'signs'
    if 'CLOCK' in n: return 'signs'
    return 'other'

tf = [f for f in T['fixtures'] if f['cat']=='elec']
rf = R['devices']
tk = collections.defaultdict(list)
rk = collections.defaultdict(list)
for f in tf: tk[klass(f.get('load'), f['typ'])].append(f)
for d in rf: rk[klass(d.get('load'), d['typ'])].append(d)

print("%-16s %5s %5s | %-14s %-14s %-9s %s" % ('class','truth','recr','pos<=2ft','pos<=6ft','type-ok','panel-ok'))
allT=allM2=allM6=0
detail = {}
for k in sorted(set(list(tk)+list(rk))):
    tl, rl = tk.get(k,[]), rk.get(k,[])
    used=[False]*len(rl)
    m2=m6=ty=pn=0
    pairs=[]
    for f in tl:
        best=(1e9,None)
        for j,d in enumerate(rl):
            if used[j]: continue
            dd=math.hypot(f['x']-d['x'], f['y']-d['y'])
            if dd<best[0]: best=(dd,j)
        if best[1] is not None and best[0]<=15:
            j=best[1]; used[j]=True
            d=rl[j]
            pairs.append((f,d,best[0]))
            if best[0]<=2: m2+=1
            if best[0]<=6: m6+=1
            if f['typ']==d['typ']: ty+=1
            if (f.get('panel') or '')==(d.get('panel') or ''): pn+=1
    print("%-16s %5d %5d | %-14s %-14s %-9s %s" % (k, len(tl), len(rl),
        '%d (%.0f%%)'%(m2,100*m2/max(1,len(tl))),
        '%d (%.0f%%)'%(m6,100*m6/max(1,len(tl))),
        '%d'%ty, '%d'%pn))
    allT+=len(tl); allM2+=m2; allM6+=m6
    detail[k]={'truth':len(tl),'recr':len(rl),'m2':m2,'m6':m6,'type_ok':ty,'panel_ok':pn,
               'worst':[(round(p[2],1), p[0].get('load'), (round(p[0]['x'],1),round(p[0]['y'],1))) for p in sorted(pairs,key=lambda p:-p[2])[:3]]}
print("TOTAL truth=%d  pos<=2ft=%d (%.0f%%)  pos<=6ft=%d (%.0f%%)" % (allT, allM2, 100*allM2/allT, allM6, 100*allM6/allT))
json.dump(detail, open('analysis/compare.json','w'), indent=1)

print()
print("=== unmatched truth per class (first few) ===")
for k in sorted(tk):
    tl, rl = tk.get(k,[]), rk.get(k,[])
    used=[False]*len(rl)
    un=[]
    for f in tl:
        best=(1e9,None)
        for j,d in enumerate(rl):
            if used[j]: continue
            dd=math.hypot(f['x']-d['x'], f['y']-d['y'])
            if dd<best[0]: best=(dd,j)
        if best[1] is not None and best[0]<=6:
            used[best[1]]=True
        else:
            un.append(f)
    if un:
        print(" ", k, len(un), [(f['typ'][:18], round(f['x'],1), round(f['y'],1), (f.get('load') or '')[:22]) for f in un[:4]])
