import json, os, math
from collections import defaultdict
RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
d = json.load(open(os.path.join(RUN, "extract.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

# BCS region
X0,X1,Y0,Y1 = 195, 272, -306, -278

# spa equipment clusters (A-N-SPA EQUIPMENT) in BCS
spa = []
for s in g.get("shapes", []):
    lay = (s.get("lay") or "").upper()
    if "SPA EQUIP" in lay and "bb" in s:
        cx, cy = s["c"]
        if X0<=cx<=X1 and Y0<=cy<=Y1:
            spa.append((cx, cy, s["bb"]))
# cluster within 3 ft
clusters = []
for (x,y,bb) in sorted(spa):
    hit=None
    for c in clusters:
        if math.hypot(x-c['x'],y-c['y'])<3.5: hit=c;break
    if hit:
        hit['pts'].append((x,y)); hit['n']+=1
        hit['x']=sum(p[0] for p in hit['pts'])/len(hit['pts']); hit['y']=sum(p[1] for p in hit['pts'])/len(hit['pts'])
    else:
        clusters.append({'x':x,'y':y,'pts':[(x,y)],'n':1})
big = [c for c in clusters if c['n']>=3]
print("BCS spa-equipment clusters (n>=3):", len(big))
for c in sorted(big, key=lambda c:(c['x'])):
    print("  (%.1f, %.1f) n=%d"%(c['x'],c['y'],c['n']))

# walls in BCS region (lines)
walls=[]
WK=("WALL","DEMISING","LEASE","GLAZ")
for s in g.get("shapes", []):
    lay=(s.get("lay") or "").upper()
    if "l" in s and any(k in lay for k in WK) and "DEMO" not in lay:
        x0,y0,x1,y1=s["l"]
        mx,my=(x0+x1)/2,(y0+y1)/2
        if X0-3<=mx<=X1+3 and Y0-5<=my<=Y1+5:
            walls.append([round(x0,2),round(y0,2),round(x1,2),round(y1,2),lay])
print("BCS wall segments:", len(walls))

# staged spa fixtures
spa_loads = ["HYDROMASSAGE - 103A","HYDROMASSAGE RECEPT - 103A","CRYOLOUNGE - 103",
             "MASSAGE CHAIRS - 103","HYBRID TANNER - 103F","STAND-UP TANNER - 103E",
             "TANNING BED - 103L","IT SERVER RACK"]
staged=defaultdict(list)
for f in d.get("fixtures", []):
    ln=f.get("params",{}).get("CKT_Load Name_CEDT")
    if ln in spa_loads and f.get("loc"):
        staged[ln].append({"id":f["id"],"loc":f["loc"][:2],"rot":f.get("rot") or 0})
print("\nstaged spa fixtures:")
for ln in spa_loads:
    print("  %-30s %d"%(ln,len(staged[ln])))

# BCS room centroids
rooms=[s for s in d.get("spaces",[]) if s.get("loc") and s.get("number","").startswith("103")]
print("\nBCS rooms:")
for r in sorted(rooms,key=lambda r:r["number"]):
    print("  %s %s (%.0f,%.0f)"%(r["number"],r["name"],r["loc"][0],r["loc"][1]))

json.dump({"clusters":big,"walls":walls,"staged":{k:v for k,v in staged.items()},
           "rooms":[{"num":r["number"],"name":r["name"],"loc":r["loc"][:2]} for r in rooms]},
          open(os.path.join(RUN,"bcs_ctx.json"),"w"))
