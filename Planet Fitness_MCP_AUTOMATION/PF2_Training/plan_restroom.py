# Restroom/locker plumbing-anchored placement: sinks (GFCI), hand dryers.
import json, os, math
from collections import defaultdict
RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
d = json.load(open(os.path.join(RUN, "extract.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

# staged restroom fixtures
loads=["MENS RESTRM SINKS","BATHROOM SINK","BATHROOM SINKS","HAND DRYER"]
staged=defaultdict(list)
for f in d.get("fixtures",[]):
    ln=f.get("params",{}).get("CKT_Load Name_CEDT")
    if ln in loads and f.get("loc"):
        staged[ln].append({"id":f["id"],"loc":f["loc"][:2]})
for ln in loads:
    if staged[ln]:
        xs=[f["loc"][0] for f in staged[ln]]; ys=[f["loc"][1] for f in staged[ln]]
        print("%-20s n=%d x[%.0f,%.0f] y[%.0f,%.0f]"%(ln,len(staged[ln]),min(xs),max(xs),min(ys),max(ys)))

# toilet rooms: 104A (151,-296), 105A (131,-296). Plumb fixtures (troughs) there.
plumb=[]
for s in g.get("shapes",[]):
    lay=(s.get("lay") or "").upper()
    if "PLUMB FIX" in lay and "bb" in s:
        cx,cy=s["c"]
        if 120<=cx<=175 and -305<=cy<=-288:
            plumb.append((cx,cy,max(s["bb"][2]-s["bb"][0],s["bb"][3]-s["bb"][1])))
# cluster
cl=[]
for x,y,sz in sorted(plumb):
    h=None
    for c in cl:
        if math.hypot(x-c[0],y-c[1])<3: h=c;break
    if h: h[2].append((x,y))
    else: cl.append([x,y,[(x,y)]])
print("toilet-room plumb clusters:", [(round(c[0],1),round(c[1],1),len(c[2])) for c in cl])

# walls near toilet rooms
walls=[]
for s in g.get("shapes",[]):
    lay=(s.get("lay") or "").upper()
    if "l" in s and ("WALL" in lay or "PARTITION" in lay) and "DEMO" not in lay:
        x0,y0,x1,y1=s["l"]
        if math.hypot(x1-x0,y1-y0)<1: continue
        mx,my=(x0+x1)/2,(y0+y1)/2
        if 120<=mx<=175 and -305<=my<=-288: walls.append((x0,y0,x1,y1))
json.dump({"staged":{k:v for k,v in staged.items()},"plumb":plumb,"walls":walls},
          open(os.path.join(RUN,"restroom_ctx.json"),"w"))
print("walls near toilets:", len(walls))
