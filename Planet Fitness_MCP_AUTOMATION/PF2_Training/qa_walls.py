# QA: every fixture must be anchored to its CAD feature. Flag floaters.
import json, os, math
from collections import defaultdict
RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
fx = json.load(open(os.path.join(RUN, "fixtures_now.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

walls=[]; tvblocks=[]; fanblocks=[]; raceway=[]
for s in g.get("shapes", []):
    lay=(s.get("lay") or "").upper()
    if "l" in s and ("WALL" in lay or "DEMISING" in lay or "GLAZ" in lay or "LEASE" in lay) and "DEMO" not in lay:
        x0,y0,x1,y1=s["l"]
        if math.hypot(x1-x0,y1-y0)>=1.0: walls.append((x0,y0,x1,y1))
    if "bb" in s:
        cx,cy=s["c"]; sz=max(s["bb"][2]-s["bb"][0], s["bb"][3]-s["bb"][1])
        if "TELEVISION" in lay: tvblocks.append((cx,cy))
        if "GYM EQUIP" in lay and sz>4 and abs((s["bb"][2]-s["bb"][0])-(s["bb"][3]-s["bb"][1]))<sz*0.4:
            fanblocks.append((cx,cy,sz))
        if "RACEWAY" in lay and 0.1<=sz<=0.6: raceway.append((cx,cy))

def dmin_wall(x,y):
    best=9e9
    for (x0,y0,x1,y1) in walls:
        dx,dy=x1-x0,y1-y0; L2=dx*dx+dy*dy
        if L2==0: continue
        t=max(0,min(1,((x-x0)*dx+(y-y0)*dy)/L2))
        d=math.hypot(x-(x0+t*dx),y-(y0+t*dy))
        if d<best: best=d
    return best
def dmin_pts(x,y,pts):
    best=9e9
    for p in pts:
        d=math.hypot(x-p[0],y-p[1])
        if d<best: best=d
    return best

WALL_TYPES=("Wall",)  # type name contains 'Wall'
SKIP_WALL_LOADS={"TREADMILL","STAIRMASTER","POWERED BIKE","TV TRUSS","MASSAGE CHAIRS - 103","BIG FAN"}
floaters=defaultdict(list)
for f in fx:
    ln=f.get("load"); typ=f.get("type") or ""
    if not ln: continue
    x,y=f["x"],f["y"]
    # TV loads -> should be on a TV block
    if ln and ("TV" in ln or "RADIANCE" in ln) and ln!="TV TRUSS":
        dtv=dmin_pts(x,y,tvblocks)
        if dtv>1.5: floaters["TV-not-on-block"].append((f["id"],ln,round(dtv,1),x,y))
        continue
    if ln=="TV TRUSS":
        continue
    if ln in ("MASSAGE CHAIRS - 103",):
        continue
    if ln in ("TREADMILL","STAIRMASTER","POWERED BIKE"):
        dr=dmin_pts(x,y,raceway)
        if dr>1.2: floaters["cardio-off-square"].append((f["id"],ln,round(dr,1),x,y))
        continue
    if ln=="BIG FAN":
        dfan=dmin_pts(x,y,[(p[0],p[1]) for p in fanblocks])
        if dfan>2.0: floaters["bigfan-not-at-fan"].append((f["id"],ln,round(dfan,1),x,y))
        continue
    # everything else that is a Wall-type -> must be near a wall
    if "Wall" in typ or "Junction Box" in typ or "Disconnect" in typ or "Specialty" in typ or "Quad" in typ:
        dw=dmin_wall(x,y)
        if dw>1.5:
            floaters["FLOATING-wall-recep"].append((f["id"],ln,round(dw,1),x,y))

print("walls:",len(walls)," tvblocks:",len(tvblocks)," fanblocks:",len(fanblocks)," raceway:",len(raceway))
print("\n=== QA RESULTS ===")
for cat,items in floaters.items():
    print("\n%s: %d"%(cat,len(items)))
    for it in items[:40]:
        print("   id %s | %-32s | %.1f ft from target | (%.0f,%.0f)"%(it[0],it[1],it[2],it[3],it[4]))
# fan blocks sample
fanblocks.sort(key=lambda t:-t[2])
print("\nlargest gym-equip (fan candidates):", [(round(x,0),round(y,0),round(sz,1)) for x,y,sz in fanblocks[:8]])
json.dump({"floaters":{k:v for k,v in floaters.items()},"fanblocks":fanblocks,"tvblocks":tvblocks},
          open(os.path.join(RUN,"qa_result.json"),"w"))