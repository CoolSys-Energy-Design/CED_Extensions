# Documentation agent instructions

These instructions apply to every file under docs/user-guide/.

## Source of truth

- Edit Markdown source pages; never hand-edit catalog.json.
- Keep an existing page's id stable, even when its title, navigation group, ribbon location, or filename changes.
- Describe implemented behavior only. Inspect the command bundle and supporting code before documenting behavior.
- Update last_verified only after checking the page against the current implementation.

## Folder layout

- Give each extension one folder containing an index.md.
- Keep tool and guide pages flat in that extension folder by default.
- Add a subject subfolder only when it contains at least six closely related pages or shared assets. Do not mirror tabs, panels, stacks, pulldowns, or pushbutton bundles automatically.
- Store shared images in docs/user-guide/assets/; use descriptive kebab-case filenames.

## Navigation

- ribbon_path records the real user-facing ribbon location and remains independent of folders.
- navigation_path controls the viewer tree. It must be a YAML list relative to the extension root.
- Use navigation_path: [] for pages displayed directly under the extension.
- Use a concise user-facing topic such as navigation_path: [Circuit Tools] or navigation_path: [Refrigeration]; do not include the extension name, shared AE pyTools tab, or a single redundant panel.
- Every extension needs one doc_type: index page with navigation_path: []. The extension header opens this page.
- Add a doc_type: index page for a navigation group when clicking that group should display an overview. Its navigation_path must exactly match the group.
- A header without an index page only expands and collapses.

## Required frontmatter

Start from _tool-template.md. Content pages require:

    ---
    id: stable-kebab-case-id
    doc_type: tool
    title: Displayed command name
    summary: One-sentence outcome.
    extension: Extension name
    ribbon_path: Ribbon Tab > Panel > Command
    navigation_path: []
    status: production
    audience: [all]
    model_impact: none
    keywords: [search, terms]
    aliases: []
    last_verified: "YYYY-MM-DD"
    ---

Use doc_type: guide for a page covering several related commands and doc_type: index for an extension or topic overview.

## Links and release mapping

- Use relative Markdown links and keep them valid when moving pages.
- Link related pages by purpose, not merely because they are nearby on the ribbon.
- Map every production tool page to its real command bundle in release-manifest.json. One guide may map multiple bundles.
- Never infer navigation from the physical folder or hardcode the current ribbon structure into viewer code.

## Validation

From the repository root, run:

    python tools/docs/validate_docs.py
    python tools/docs/generate_catalog.py
    python -m unittest discover -s tools/docs/tests -v

If python is unavailable on PATH, use the configured local Python executable. Validation must complete with zero errors and zero warnings before handoff.
