# Full BCS wall-snap placement plan (centroid-based, no bbox needed).
import json, os, math
from collections import defaultdict
RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
d = json.load(open(os.path.join(RUN, "extract.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

walls=[]
WK=("WALL","DEMISING","LEASE","GLAZ")
for s in g.get("shapes", []):
    lay=(s.get("lay") or "").upper()
    if "l" in s and any(k in lay for k in WK) and "DEMO" not in lay:
        x0,y0,x1,y1=s["l"]
        if math.hypot(x1-x0,y1-y0) < 1.5: continue
        mx,my=(x0+x1)/2,(y0+y1)/2
        if 150<=mx<=275 and -312<=my<=-258:
            walls.append((x0,y0,x1,y1))

ctr={}
for s in d.get("spaces",[]):
    if s.get("loc"): ctr[s["number"]]=(s["loc"][0],s["loc"][1])

def nearby_walls(cx,cy,rad=9):
    out=[]
    for (x0,y0,x1,y1) in walls:
        dx,dy=x1-x0,y1-y0; L2=dx*dx+dy*dy
        if L2==0: continue
        t=max(0,min(1,((cx-x0)*dx+(cy-y0)*dy)/L2))
        sx,sy=x0+t*dx,y0+t*dy
        dist=math.hypot(cx-sx,cy-sy)
        if dist<=rad: out.append((dist,x0,y0,x1,y1,math.hypot(dx,dy)))
    return out

def snap_point(px,py,cx,cy,rad=9,prefer_len=True):
    """snap (px,py) to best nearby wall of room centered (cx,cy). face into room."""
    cand=nearby_walls(cx,cy,rad)
    if not cand: return None
    best=None
    for (dist,x0,y0,x1,y1,seglen) in cand:
        dx,dy=x1-x0,y1-y0; L2=dx*dx+dy*dy
        t=max(0,min(1,((px-x0)*dx+(py-y0)*dy)/L2))
        sx,sy=x0+t*dx,y0+t*dy
        pdist=math.hypot(px-sx,py-sy)
        score=pdist-(0.2*seglen if prefer_len else 0)
        if best is None or score<best[0]:
            best=(score,sx,sy,dx,dy)
    _,sx,sy,dx,dy=best
    nx,ny=-dy,dx; n=math.hypot(nx,ny); nx,ny=nx/n,ny/n
    if (cx-sx)*nx+(cy-sy)*ny<0: nx,ny=-nx,-ny
    sx2,sy2=sx+0.5*nx,sy+0.5*ny
    r=math.atan2(-nx,ny); r=round(r/(math.pi/2))*(math.pi/2)
    return (round(sx2,3),round(sy2,3),round(r%(2*math.pi),5))

staged=defaultdict(list)
for f in d.get("fixtures",[]):
    ln=f.get("params",{}).get("CKT_Load Name_CEDT")
    if ln and ("HYDROMASSAGE" in ln or "CRYOLOUNGE" in ln or "IT SERVER" in ln
               or "TANNER" in ln or "TANNING BED" in ln or "RED WAVE" in ln):
        if f.get("loc"): staged[ln].append({"id":f["id"],"loc":f["loc"][:2]})

plan=[]; flags=[]

# tanning: 1 disconnect per tanning room, assign by nearest current pos
tanning_rooms=[r for r in ["103B","103C","103D","103E","103F","103K","103L"] if r in ctr]
disc=[f for ln,fs in staged.items() if ("TANNER" in ln or "TANNING BED" in ln or "RED WAVE" in ln) for f in fs]
pairs=[]
for f in disc:
    for r in tanning_rooms:
        pairs.append((math.hypot(f["loc"][0]-ctr[r][0],f["loc"][1]-ctr[r][1]),f["id"],r,f))
pairs.sort()
usedr=set(); usedf=set()
for dist,fid,r,f in pairs:
    if fid in usedf or r in usedr: continue
    usedf.add(fid); usedr.add(r)
    cx,cy=ctr[r]
    sp=snap_point(cx,cy,cx,cy,rad=8)
    if sp: plan.append({"id":fid,"x":sp[0],"y":sp[1],"rot":sp[2],"grp":"tanning","room":r})
    else: flags.append("no wall found for tanning room %s"%r)
empty=[r for r in tanning_rooms if r not in usedr]
if empty: flags.append("tanning rooms without disconnect: %s"%empty)

def distribute_on_wall(cx,cy,items,room,grp,along="auto",second_off=0.0):
    """place items along the longest nearby wall of the room, evenly."""
    cand=nearby_walls(cx,cy,10)
    if not cand:
        flags.append("no wall for %s room %s"%(grp,room)); return
    # pick longest wall
    cand.sort(key=lambda c:-c[5])
    _,x0,y0,x1,y1,seglen=cand[0]
    dx,dy=x1-x0,y1-y0; L=math.hypot(dx,dy); ux,uy=dx/L,dy/L
    nx,ny=-uy,ux
    if (cx-(x0+x1)/2)*nx+(cy-(y0+y1)/2)*ny<0: nx,ny=-nx,-ny
    r=math.atan2(-nx,ny); r=round(r/(math.pi/2))*(math.pi/2); r=round(r%(2*math.pi),5)
    n=len(items)
    for i,f in enumerate(items):
        t=(i+0.5)/n
        bx,by=x0+dx*t, y0+dy*t
        px,py=bx+0.5*nx+second_off*nx, by+0.5*ny+second_off*ny
        plan.append({"id":f["id"],"x":round(px,3),"y":round(py,3),"rot":r,"grp":grp,"room":room})

# hydromassage pairs in 103A (specialty + GFCI side by side)
if "103A" in ctr:
    cx,cy=ctr["103A"]
    hyd=[f for ln,fs in staged.items() if ln.startswith("HYDROMASSAGE - ") for f in fs]
    rec=[f for ln,fs in staged.items() if ln.startswith("HYDROMASSAGE RECEPT") for f in fs]
    distribute_on_wall(cx,cy,hyd,"103A","hydro")
    distribute_on_wall(cx,cy,rec,"103A","hydro_recept",second_off=0.6)
    flags.append("hydro placed along 103A longest wall (no chair blocks) - VERIFY")

# cryo in 103G
if "103G" in ctr:
    cx,cy=ctr["103G"]
    cry=[f for ln,fs in staged.items() if ln.startswith("CRYOLOUNGE") for f in fs]
    distribute_on_wall(cx,cy,cry,"103G","cryo")

# IT quads in 101A
if "101A" in ctr:
    cx,cy=ctr["101A"]
    it=[f for ln,fs in staged.items() if ln.startswith("IT SERVER") for f in fs]
    distribute_on_wall(cx,cy,it,"101A","IT")
    flags.append("IT quads along 101A longest wall (no rack blocks) - VERIFY")

json.dump({"plan":plan,"flags":flags}, open(os.path.join(RUN,"plan_bcs.json"),"w"), indent=0)
print("BCS plan entries:", len(plan))
byg=defaultdict(int)
for p in plan: byg[p["grp"]]+=1
print("by group:", dict(byg))
for fl in flags: print("FLAG:",fl)
for p in plan:
    if p["grp"]=="tanning": print("  tanning %s -> %s (%.1f,%.1f)"%(p["id"],p["room"],p["x"],p["y"]))