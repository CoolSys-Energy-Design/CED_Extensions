# Refine restroom GFCIs to Sloan-trough layout: (n-2) on back wall evenly +
# 2 on the alcove side walls at the ends, facing inward, x-aligned with back row.
import json, os, math
from collections import defaultdict
RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
d = json.load(open(os.path.join(RUN, "extract.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

# staged sink ids by load
sinks=defaultdict(list)
for f in d.get("fixtures",[]):
    ln=f.get("params",{}).get("CKT_Load Name_CEDT")
    if ln in ("MENS RESTRM SINKS","BATHROOM SINK") and f.get("loc"):
        sinks[ln].append(f["id"])

# walls near toilets, classify H/V
H=[]; V=[]
for s in g.get("shapes",[]):
    lay=(s.get("lay") or "").upper()
    if "l" not in s: continue
    if not (("WALL" in lay or "PARTITION" in lay) and "DEMO" not in lay): continue
    x0,y0,x1,y1=s["l"]
    L=math.hypot(x1-x0,y1-y0)
    if L<1: continue
    mx,my=(x0+x1)/2,(y0+y1)/2
    if not (115<=mx<=170 and -306<=my<=-290): continue
    if abs(y1-y0)<abs(x1-x0): H.append((x0,y0,x1,y1,L,my))
    else: V.append((x0,y0,x1,y1,L,mx))

def trough_layout(ids, xzone, wall_x_lo, wall_x_hi):
    """place ids: (n-2) on back(south) wall + 2 on side walls at ends, facing in."""
    n=len(ids)
    if n<3: return []
    # back wall = southern-most long horizontal within the x zone
    hz=[w for w in H if wall_x_lo<= (w[0]+w[2])/2 <=wall_x_hi]
    if not hz: return []
    backy=min(w[5] for w in hz)   # most negative y = south
    back=[w for w in hz if abs(w[5]-backy)<1.5]
    bx0=min(min(w[0],w[2]) for w in back); bx1=max(max(w[0],w[2]) for w in back)
    # trim to the requested zone
    bx0=max(bx0,wall_x_lo); bx1=min(bx1,wall_x_hi)
    span=bx1-bx0
    col_x=(bx0+bx1)/2   # the x the outer pair aligns to? we keep back row across span
    out=[]
    nb=n-2  # back wall count
    for i in range(nb):
        t=(i+1)/(nb+1)
        out.append((round(bx0+span*t,3), round(backy+0.5,3), 0.0))  # face north (rot 0)
    # 2 side walls: verticals near the two x-ends of the back wall
    vz=sorted(V, key=lambda w: min(abs((w[0]+w[2])/2-bx0),abs((w[0]+w[2])/2-bx1)))
    # north side of alcove: place the outer pair at the ends, x-aligned to back span ends,
    # a bit north of back wall, facing inward (toward each other)
    yy=backy+2.5
    # left outer faces east (rot=-90/270 -> face +X? ) ; we want them pointing inward across trough
    out.append((round(bx0+0.6,3), round(yy,3), round(3*math.pi/2,5)))   # left, face east(inward)
    out.append((round(bx1-0.6,3), round(yy,3), round(math.pi/2,5)))     # right, face west(inward)
    return out[:n]

plan=[]
# men's zone x[118,145], women's x[145,168]
for load,lo,hi in (("MENS RESTRM SINKS",117,145),("BATHROOM SINK",145,168)):
    ids=sinks.get(load,[])
    pts=trough_layout(ids,None,lo,hi)
    for i,(x,y,r) in enumerate(pts):
        if i<len(ids):
            plan.append({"id":ids[i],"x":x,"y":y,"rot":r,"load":load})

json.dump(plan, open(os.path.join(RUN,"plan_restroom2.json"),"w"))
print("restroom refine entries:", len(plan))
for p in plan: print("  %s (%.1f,%.1f) rot%.2f %s"%(p["id"],p["x"],p["y"],p["rot"],p["load"]))