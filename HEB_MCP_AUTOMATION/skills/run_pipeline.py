# SKILL: run_pipeline  (one area, full replication into the target)
# 1) PURGE any elements from a previous/aborted run for this area (via the
#    existing id map) so re-runs are clean & idempotent.
# 2) Reset the id map.
# 3) Run the 7 placement steps in order. Flags are PREPENDED as code into
#    each skill source (robust across nested exec scopes).
# Requires BA_DATA / BA_VIEW already defined (run_all prepends them).
exec(open(r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills\_lib.py").read())
import os, traceback

SKILLS = r"c:\CED_Extensions\HEB_MCP_AUTOMATION\skills"
MAP = DATA + r"\place_relative_map.json"
TGT_TITLE = "OkmntProfCorr"
_alldocs = list(doc.Application.Documents)
_tgt = None
for _d in _alldocs:
    if TGT_TITLE in _d.Title:
        _tgt = _d; break

# ---- 1) PURGE prior run for this area ----
purged = 0
if os.path.exists(MAP):
    try:
        oldm = json.load(io.open(MAP))
    except:
        oldm = {}
    ids = []
    for key in ("fixture_tags", "wire_tags", "wire_arrows", "wires",
                "textnotes", "keynotes", "devices"):
        v = oldm.get(key, {})
        if isinstance(v, dict):
            for val in v.values():
                if isinstance(val, list): ids += val
                else: ids.append(val)
    for rec in oldm.get("circuits", {}).values():
        if isinstance(rec, dict) and rec.get("tgt_sys_id"):
            ids.append(rec["tgt_sys_id"])
    if ids and _tgt is not None:
        pt = DB.Transaction(_tgt, "BAKERY/PHARMACY purge prior run")
        pt.Start()
        for i in ids:
            try:
                el = _tgt.GetElement(DB.ElementId(int(i)))
                if el is not None:
                    _tgt.Delete(el.Id); purged += 1
            except: pass
        pt.Commit()
print("== area DATA=%s VIEW=%r  purged=%d prior elements" % (DATA, VIEW_NAME, purged))

# ---- 2) reset map ----
with io.open(MAP, 'w', encoding='utf-8') as f:
    f.write(u"{}")
print("== map reset:", MAP)

# ---- 3) run steps (flags prepended as code) ----
STEPS = [
    ("place_relative.py",     {"BA_DRYRUN": False}),
    ("place_keynotes.py",     {"BA_KN_APPLY": True}),
    ("place_textnotes.py",    {"BA_TN_APPLY": True, "BA_TN_REBUILD": False}),
    ("place_circuits.py",     {"BA_C_APPLY": True}),
    ("place_wires.py",        {"BA_W_APPLY": True, "BA_W_REBUILD": False}),
    ("place_fixture_tags.py", {"BA_FT_APPLY": True}),
    ("place_wire_tags.py",    {"BA_WT_APPLY": True}),
]
for i, (fn, flags) in enumerate(STEPS, 1):
    hdr = "".join("%s = %r\n" % (k, v) for k, v in flags.items())
    print("\n===== STEP %d/%d : %s =====" % (i, len(STEPS), fn))
    _src = hdr + open(SKILLS + "\\" + fn).read()
    try:
        exec(_src)
    except Exception:
        print("!! PIPELINE STOPPED at step %d (%s):" % (i, fn))
        print(traceback.format_exc())
        raise

m = json.load(io.open(MAP))
print("\n========== PIPELINE COMPLETE (%s) ==========" % VIEW_NAME)
for key in ("devices", "keynotes", "textnotes", "circuits", "wires",
            "wire_arrows", "fixture_tags", "wire_tags"):
    print("  %-13s %d" % (key, len(m.get(key, {}))))
print("  flags:", m.get("flags"))
print("============================================")
