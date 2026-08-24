# Auto-detect the current non-template PF project doc, run E101 + CAD extraction into
# a slug folder. Prints the doc title + a suggested slug. Read-only.
import json, os, re, unicodedata

TEMPLATE_HINTS = ("profiles 2.0", "template", "practice", "profiles2")
ROOT = r"c:\CED_Extensions\Planet Fitness_MCP_AUTOMATION\PF2_Training"

def slugify(title):
    t = title.lower()
    t = t.replace("planet fitness", "").replace("corporate", "").replace("takeover", "")
    t = t.replace("grand fitness", "").replace("excel", "").replace("flynn", "").replace("ignite", "")
    t = re.sub(r"[^a-z0-9]+", "_", t).strip("_")
    t = re.sub(r"_+", "_", t)
    return t or "project"

app = uidoc.Application.Application
target = None
cand = []
for d2 in app.Documents:
    if d2.IsLinked:
        continue
    low = d2.Title.lower()
    if any(h in low for h in TEMPLATE_HINTS):
        continue
    if "planet" in low or "fitness" in low:
        cand.append(d2)
# prefer a doc whose slug folder has no extract.json yet (not collected)
for d2 in cand:
    if not os.path.exists(os.path.join(ROOT, slugify(d2.Title), "extract.json")):
        target = d2
if target is None and cand:
    target = cand[-1]
if target is None:
    print(json.dumps({"ok": False, "err": "no non-template PF doc open",
                      "open": [d.Title for d in app.Documents if not d.IsLinked]}))
else:
    slug = slugify(target.Title)
    out_dir = os.path.join(ROOT, slug)
    print(json.dumps({"ok": True, "title": target.Title, "slug": slug, "out_dir": out_dir}))
