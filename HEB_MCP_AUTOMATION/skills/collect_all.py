# SKILL: collect_all  (read-only)  -- gather ALL source data for one area.
# Area is configured by globals BEFORE exec:
#   BA_DATA  = absolute data dir for this area
#   BA_VIEW  = the source/target view name (same name both projects)
# Runs every collector in dependency order against the SOURCE project.
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())

SKILLS = r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills"
g = globals()
# capture SOURCE document once (two-step; no self-referential genexpr)
_docs = list(doc.Application.Documents)
SRC = None
for _d in _docs:
    if _d.Title.startswith("RunUpdateProfiles"):
        SRC = _d; break
print("== COLLECT area: DATA=%s VIEW=%r src=%s" % (DATA, VIEW_NAME, SRC.Title))

ORDER = [
    "collect_linked_elements.py",
    "collect_host_elements.py",
    "collect_keynotes.py",
    "collect_wires.py",
    "collect_circuits.py",
    "collect_textnotes.py",
    "collect_fixture_tags.py",
    "collect_wire_tags.py",
]
import traceback
for i, fn in enumerate(ORDER, 1):
    print("\n--- collect %d/%d : %s ---" % (i, len(ORDER), fn))
    globals()["doc"] = SRC               # re-inject every iteration
    _src = open(SKILLS + "\\" + fn).read()
    try:
        exec(_src)                       # bare: share this frame's namespace (has DB)
    except Exception:
        print("!! COLLECT FAILED at", fn); print(traceback.format_exc()); raise
print("\n== COLLECTION COMPLETE for", DATA)
