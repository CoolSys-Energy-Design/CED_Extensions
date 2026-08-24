# Finalize placement: tighten hydro to 103A, best-guess check-in TVs, big fans, vending.
import json, os, math
from collections import defaultdict
RUN = r"C:\Users\reed.pinterich\.claude\skills\pf-power-plan-2\runs\run-2026-08-21_155402"
d = json.load(open(os.path.join(RUN, "extract.json")))
g = json.load(open(os.path.join(RUN, "cad_geom.json")))

# current fixture ids by load (from extract - ids stable)
byload=defaultdict(list)
for f in d.get("fixtures",[]):
    ln=f.get("params",{}).get("CKT_Load Name_CEDT")
    if ln: byload[ln].append(f["id"])

walls=[]
for s in g.get("shapes",[]):
    lay=(s.get("lay") or "").upper()
    if "l" in s and ("WALL" in lay or "DEMISING" in lay or "GLAZ" in lay or "LEASE" in lay) and "DEMO" not in lay:
        x0,y0,x1,y1=s["l"]
        if math.hypot(x1-x0,y1-y0)<1.5: continue
        walls.append((x0,y0,x1,y1))

def walls_near(cx,cy,rad):
    out=[]
    for (x0,y0,x1,y1) in walls:
        dx,dy=x1-x0,y1-y0; L2=dx*dx+dy*dy
        t=max(0,min(1,((cx-x0)*dx+(cy-y0)*dy)/L2))
        sx,sy=x0+t*dx,y0+t*dy; dist=math.hypot(cx-sx,cy-sy)
        if dist<=rad: out.append((dist,x0,y0,x1,y1,math.hypot(dx,dy)))
    return out

def face_rot(sx,sy,cx,cy,dx,dy):
    nx,ny=-dy,dx; n=math.hypot(nx,ny); nx,ny=nx/n,ny/n
    if (cx-sx)*nx+(cy-sy)*ny<0: nx,ny=-nx,-ny
    r=math.atan2(-nx,ny); return round((round(r/(math.pi/2))*(math.pi/2))%(2*math.pi),5),nx,ny

plan=[]
ctr={s["number"]:(s["loc"][0],s["loc"][1]) for s in d["spaces"] if s.get("loc")}

# --- hydro tighten to 103A: pick the longest wall within 5 ft of 103A centroid ---
cx,cy=ctr["103A"]
cand=sorted(walls_near(cx,cy,6), key=lambda w:-w[5])
if cand:
    _,x0,y0,x1,y1,seglen=cand[0]
    dx,dy=x1-x0,y1-y0; L=math.hypot(dx,dy); ux,uy=dx/L,dy/L
    r,nx,ny=face_rot((x0+x1)/2,(y0+y1)/2,cx,cy,dx,dy)
    hyd=byload["HYDROMASSAGE - 103A"]; rec=byload["HYDROMASSAGE RECEPT - 103A"]
    npair=max(len(hyd),len(rec))
    for i in range(len(hyd)):
        t=(i+0.5)/max(1,len(hyd)); bx,by=x0+dx*t,y0+dy*t
        plan.append({"id":hyd[i],"x":round(bx+0.4*nx,3),"y":round(by+0.4*ny,3),"rot":r,"grp":"hydro"})
    for i in range(len(rec)):
        t=(i+0.5)/max(1,len(rec)); bx,by=x0+dx*t,y0+dy*t
        # GFCI offset 0.6 ft along wall from its specialty
        plan.append({"id":rec[i],"x":round(bx+0.4*nx+0.6*ux,3),"y":round(by+0.4*ny+0.6*uy,3),"rot":r,"grp":"hydro_gfci"})

# --- check-in TVs (7): evenly along a line in front of check-in counter ---
tvs=byload["TV & RADIANCE MONITOR - CHECK-IN 102"]
# check-in counter area ~ x 240..264 at y ~ -250 (reception). place along y=-249
n=len(tvs)
for i,fid in enumerate(sorted(tvs)):
    x=240 + (264-240)*(i+0.5)/n
    plan.append({"id":fid,"x":round(x,3),"y":-249.0,"rot":0.0,"grp":"checkin_tv"})

# --- big fans (2): gym visual centers (open floor) ---
fans=byload["BIG FAN"]
fan_pts=[(140.0,-212.0),(180.0,-212.0)]
for i,fid in enumerate(sorted(fans)):
    if i<len(fan_pts):
        plan.append({"id":fid,"x":fan_pts[i][0],"y":fan_pts[i][1],"rot":math.pi,"grp":"bigfan"})

# --- vending (3): snap to nearest wall near their staging cluster ---
ven=byload["VENDING MACHINE"]
# staging ~ (214,-250). snap each to nearest wall, spread along it
vc=sorted(walls_near(214,-250,8), key=lambda w:-w[5])
if vc:
    _,x0,y0,x1,y1,seglen=vc[0]
    dx,dy=x1-x0,y1-y0; L=math.hypot(dx,dy)
    r,nx,ny=face_rot((x0+x1)/2,(y0+y1)/2,214,-250,dx,dy)
    for i,fid in enumerate(sorted(ven)):
        t=(i+0.5)/max(1,len(ven)); bx,by=x0+dx*t,y0+dy*t
        plan.append({"id":fid,"x":round(bx+0.4*nx,3),"y":round(by+0.4*ny,3),"rot":r,"grp":"vending"})

json.dump(plan, open(os.path.join(RUN,"plan_finalize.json"),"w"))
byg=defaultdict(int)
for p in plan: byg[p["grp"]]+=1
print("finalize plan:", dict(byg), "total", len(plan))