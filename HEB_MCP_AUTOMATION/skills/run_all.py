# SKILL: run_all  (MASTER)  -- on "Go": full placement pipeline for BOTH areas.
# Each area = (label, data dir, view name). Same shared skills, per-area config
# via BA_DATA / BA_VIEW globals. Each area gets its own reset id map and runs
# the 7-step pipeline (devices->keynotes->textnotes->circuits->wires->
# fixture tags->wire tags). Target project is the same (OkmntProfCorr),
# different view per area.
import io as _io
SKILLS = r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills"
AREAS = [
    ("BAKERY",   r"c:\CED_Extensions\HEB_MCP_AUTOMATION\bakery_auto\data",
                 "Power Callout - BAKERY - L1"),
    ("PHARMACY", r"c:\CED_Extensions\HEB_MCP_AUTOMATION\pharmacy_auto\data",
                 "Power Callouts - PHARMACY"),
]
g = globals()
for label, dpath, view in AREAS:
    g["BA_DATA"] = dpath
    g["BA_VIEW"] = view
    print("\n##################  AREA: %s  ##################" % label)
    print("   data=%s" % dpath)
    print("   view=%r" % view)
    _hdr = "BA_DATA = %r\nBA_VIEW = %r\n" % (dpath, view)
    _rp = _hdr + open(SKILLS + r"\run_pipeline.py").read()
    exec(_rp)                            # flags prepended as code; share namespace
print("\n##################  ALL AREAS COMPLETE  ##################")
