# Documentation Viewer Implementation Record

## Command location

The offline viewer replaces the existing external documentation URL button at:

```text
Extension: AE pyTools
Tab:       AE pyTools
Panel:     CED Tools
Button:    Docs
```

The pyRevit bundle is `Docs.pushbutton`. No additional viewer button is added.

## Runtime and compatibility baseline

- Installed pyRevit baseline: `6.5.4.26228+1146`
- Selected Python engine: pyRevit default IronPython 2.7.12 (`IPY2712PR`)
- WPF target: PresentationFramework on .NET Framework (Revit 2024) and .NET Core (Revit 2025–2026)
- Initial Revit release targets: Revit 2024, Revit 2025, and Revit 2026
- Renderer: native WPF `FlowDocument`; no WebView or network dependency
- Result navigation: ranked list or a metadata-driven extension/topic tree with clickable indexes
- Theme semantics: document canvas and hyperlink colors use shared global CED tokens

The installed pyRevit clone also provides CPython 3.12 (`CPY3123`), but the release verification below uses the default engine that executes the ribbon command. The runtime code remains Python 2 compatible; build tools and runtime-independent tests use standard CPython.

Automated source, catalog, parser, search, path-containment, and navigation checks run outside Revit. Before publication, the renderer compatibility fixture must also be exercised inside each exact Revit build being released. Exact in-host build numbers and results belong in the verification table below; a blank or failing row blocks release.

The in-host smoke test is `tools/docs/revit_renderer_smoke.py`. Set `CED_DOC_TEST_REPO` to the repository root and invoke it with `pyrevit run ... --revit=<year>`.

## Verification record

| Revit version/build | Host runtime | pyRevit / Python engine | Fixture result | Verified date |
|---|---|---|---|---|
| Revit 2024.3.5 — `VersionBuild 24.3.50.51`, `20260518_1515(x64)` | .NET Framework | `6.5.4.26228+1146` / IronPython 2.7.12 | Pass — 22 catalog pages, 26 fixture blocks | 2026-08-24 |
| Revit 2025 — `VersionBuild 25.4.60.9`, runner build `20240307_1300(x64)` | .NET Core | `6.5.4.26228+1146` / IronPython 2.7.12 | Pass — 22 catalog pages, 26 fixture blocks | 2026-08-24 |
| Revit 2026 — `VersionBuild 26.4.20.9`, runner build `20250227_1515(x64)` | .NET Core | `6.5.4.26228+1146` / IronPython 2.7.12 | Pass — 54 catalog pages, 26 fixture blocks | 2026-08-24 |

The 2026 follow-up verification passed with 54 catalog pages and 26 fixture blocks. Synchronous full construction measured approximately 1.81 seconds in the runner; production startup paints the shell first and defers catalog/render work. The tree is built only when selected, and subsequent ribbon clicks reactivate the existing modeless window before importing viewer modules.

## Documentation layout and navigation

User-guide source is flat within each extension folder by default. Physical folders do not mirror ribbon tabs, panels, stacks, pulldowns, or command bundles. The `navigation_path` frontmatter field controls topic grouping independently from `ribbon_path`; index documents attach to their matching tree header. Documentation-agent rules are maintained in `docs/user-guide/AGENTS.md`.

## Release commands

From the repository root:

```powershell
python tools/docs/validate_docs.py
python tools/docs/generate_catalog.py
```

Validation reads Markdown source and `release-manifest.json`. Catalog generation refuses to run when validation has errors. `docs/user-guide/catalog.json` is generated output and is never hand-maintained.
