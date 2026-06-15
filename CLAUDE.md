# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> Personal profile, working preferences, and data-safety/convention rules live in [.claude/CLAUDE.md](.claude/CLAUDE.md). This file is the **codebase** map.

## What this repo is

A collection of **pyRevit extensions** (Python, runs inside Autodesk Revit) that automate MEP design for CoolSys. There is **no build/compile step** — pyRevit loads the `.extension` folders into Revit at startup. Scripts target the **CPython 3 engine** (note the `#! python3` shebang on `script.py` files), not IronPython.

Two distinct worlds live side by side:
1. **pyRevit tools** — the `*.extension` folders + `CEDLib.lib`, run interactively inside Revit.
2. **Revit MCP automation** — the `*_MCP_AUTOMATION` folders, driven by Claude via the `mcp__revit__*` tools against a live Revit session (no pyRevit involved).

## Running, testing, "building"

- **No compile.** Edit a `script.py`; pyRevit reloads it on next button click. Most pushbuttons call `_dev_reload.purge()` at the top to force a fresh import of the panel's `lib/` during development.
- **Tests** (MEPRFP 2.0 domain logic) live in `AE pyTools.extension/AE pyTools.Tab/MEPRFP Automation 2.0.panel/lib/` as `_test_*.py` and run via the runner there:
  ```
  python "AE pyTools.extension/AE pyTools.Tab/MEPRFP Automation 2.0.panel/lib/_run_tests.py"
  ```
  These are pure-logic tests (no Revit API) — run them in plain CPython. `_roundtrip_test.py` validates lossless YAML import→export.
- **Root-level helper scripts** (`_convert_v4_to_v5.py`, `_dedupe_v5_profiles.py`, `fix_parent_directives_V5_2.py`, `build_HEB_profiles_V5_1.py`, `append_profile_slides.py`, `MEPRFP_Automation_2_Overview.py`) are **one-off CPython utilities** that transform the `*_profiles_*.yaml` files or generate decks. They run outside Revit (`python <script>.py`). `requirements.txt` is empty; deck scripts need `python-pptx`.
- No linter is configured.

## pyRevit layout convention

`<Name>.extension / <Tab>.tab / <Panel>.panel / <Button>.pushbutton / script.py`. A `.pulldown` groups several pushbuttons. `extensions.json` (repo root) is the manifest of which extensions are enabled.

**Extensions:**
- **AE pyTools.extension** — primary/active extension. Holds the flagship **MEPRFP Automation 2.0.panel** plus ~12 other panels (Refrigeration, Mechanical, FireProtection, Selection, QualityChecks, Auto PF, etc.).
- **CED ElecTools.extension** — electrical tools (Circuit Manager, Alerts Manager).
- **CED MechTools.extension** — mechanical/refrigeration tools.
- **H-E-B Tools.extension**, **WM Tools.extension** — client-specific tool sets.
- **CEDLib.lib** — shared library (a pyRevit `.lib`, not an extension), on the path for all extensions.

## The shared library: CEDLib.lib

Reused across pushbuttons. Pushbuttons inject a `lib` dir onto `sys.path` then import flat (no package prefix):
```python
_LIB = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "..", "lib"))
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)
import active_yaml          # flat import from that lib
```
Note each panel has its **own** `lib/`; CEDLib.lib is the cross-extension shared one. Key CEDLib subpackages: `Snippets/` (param_binder.py for shared-parameter binding, family_utils, identity_mark), `CEDElectrical/` (DDD electrical model), `UIClasses/` (WPF dialog bases), `QualityChecks/`, `ExtensibleStorage/`.

## MEPRFP Automation 2.0 — the flagship (read this before touching it)

Path: `AE pyTools.extension/AE pyTools.Tab/MEPRFP Automation 2.0.panel/`. Its `lib/` (~100 modules) is the core engine. The whole panel exists to **capture a repeatable MEP layout once and place it everywhere**.

**Domain model** (`profile_model.py`):
- A **Profile** = a parent fixture (`parent_filter`) + the linked elements that travel with it.
- A **LED** (Linked Element Definition) = one element template: family:type `label`, `parameters`, `offsets` (in the parent's local frame, so they rotate with the parent), and an `annotations` list (tags/keynotes/text_notes).
- **Space profiles** (`space_profile_model.py`) are the room-based analogue: they target a **bucket** (keyword-classified room type) instead of a parent fixture, and each LED has a door-aware `placement_rule` (center, door_relative, wall_opposite_door, corner_furthest_from_door, …) — **no cardinal directions**, anchors are relative to the room's door/walls.

**Parameter directives** (`directives.py`): a LED parameter can be a static literal, `BYPARENT` (read from the host fixture), or `BYSIBLING` (read from another LED in the same set). When comparing values across elements, compare the *same* parameter across siblings — not different parameters on one element.

**Element_Linker** (`element_linker*.py`): every placed element gets a JSON payload stamped onto its `Element_Linker` shared parameter — `led_id`, `set_id`, `parent_element_id`, location/rotation, level, circuit info, and (for spaces) `space_id`/`space_profile_id`. This lineage is the **audit backbone**: it links each live element back to the YAML that produced it.

**Storage**: profiles persist project-local in **Revit Extensible Storage** on a dedicated DataStorage element (schema v4, typed map fields) — NOT on ProjectInformation. `active_yaml.py` is the high-level load/save; `storage.py` is the ES codec; YAML import/export round-trips byte-identically.

**Layered architecture** (keep this separation when editing):
- `*_model.py` — pure data classes wrapping YAML dicts (lossless `to_dict()`).
- `*_workflow.py` — pure logic, no Revit API, no UI.
- `*_apply.py` — wraps Revit transactions; writes Element_Linker.
- `*_window.py` — WPF/pyRevit-forms UI.
- `active_yaml.py` / `storage.py` / `yaml_io.py` / `schema_migrations.py` — I/O, ES, and v3/v4→v100 migration (forward-only).

**Client-specific circuiting** is pluggable: `circuit_clients/base.py` + per-client `heb.py` / `pf.py`. Add a client via a small adapter subclass rather than editing core logic.

**Notable buttons:** `Place from CAD or Linked Model` (matches parents in a linked Revit/DWG/CSV to profiles, drops every LED at its offset), `Circuiting/SuperCircuit V5` (builds Revit ElectricalSystems from captured fixtures), `Misc Ops/QAQC` (nine-category drift audit comparing Element_Linker against the live model; most findings are advisory, a few auto-fix).

## YAML profile files (repo root)

`HEB_profiles_V5*.yaml`, `PF_profiles_V4_23.yaml`, `PF_CorporateONLY_*MEPRFP_2.0.yaml`, etc. are exported profile stores (current schema_version **100**). Versioned/`.bak` copies exist because these are large hand-tuned data assets — **back up before transforming any of them** (see [.claude/CLAUDE.md](.claude/CLAUDE.md) Data Safety).

## Shared parameters

`AE CoolSys Energy Design (CED) Shared Parameters.txt` is the firm-wide Revit shared-parameter definition file (CED Electrical, CKT_* circuit params, Element_Linker, etc.). `Snippets/param_binder.py` binds these into a model by temporarily swapping the app's shared-param file, binding via union (never shrinking existing bindings), then restoring the original path.

## MCP automation (HEB_MCP_AUTOMATION, Planet Fitness_MCP_AUTOMATION)

Separate from pyRevit. These replicate a finished electrical/lighting design from a **source** Revit model onto a **target** model, both open in one Revit session and read/written through the `mcp__revit__execute_revit_code` MCP tool. Pattern: `skills/collect_*.py` read source → JSON in `data/`, then `skills/replicate.py` / `place_relative.py` place devices, keynotes, text, circuits, wires, and tags on the target. The PF folder's `SYNTHESIS.md`, `PF_CIRCUITING_ALGORITHM.md`, and `PF_POWER_PLAN_PLAYBOOK.md` capture the design rules; the `/pf-power-plan` skill orchestrates room-by-room QAQC over this.
