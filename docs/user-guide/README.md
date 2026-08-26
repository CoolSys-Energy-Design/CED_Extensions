# CED Tools User Guide

This folder is the user-facing documentation source for CED pyRevit tools. It is intentionally separate from extension bundles; bundle-local Markdown remains developer reference material.

Pages describe the currently implemented behavior of a command, including where to find it, when to use it, prerequisites, results, and limitations. Each command page uses the same structure and links to related user-guide pages.

## Extensions

- [AE pyTools](ae-pytools/index.md)
- [CED ElecTools](ced-electools/index.md)
- [CED MechTools](ced-mechtools/index.md)
- [WM Tools](wm-tools/index.md)

## Documentation standards

Use [_tool-template.md](_tool-template.md) for every new user-facing command page. Keep the `id` stable, use the displayed ribbon name as the title, and verify the page against the current command before changing `last_verified`.

Documentation agents must follow [AGENTS.md](AGENTS.md), including the flat-per-extension folder convention and metadata-driven `navigation_path` rules.

## Offline viewer

In Revit, open **AE pyTools > CED Tools > Docs** to search and read the generated offline catalog. The implementation follows [viewer-execution-plan.md](viewer-execution-plan.md).

## Release checks

Run the source validator and regenerate the catalog from the repository root:

```powershell
python tools/docs/validate_docs.py
python tools/docs/generate_catalog.py
```

`catalog.json` is generated output. Do not edit it by hand.
