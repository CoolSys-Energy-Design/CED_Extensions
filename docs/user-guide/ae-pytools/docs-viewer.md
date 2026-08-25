---
id: ae-pytools-docs
doc_type: tool
title: Docs
summary: Searches and displays the installed CED tool user guide without requiring internet access.
extension: AE pyTools
ribbon_path: AE pyTools > CED Tools > Docs
navigation_path: []
status: production
audience: [all]
model_impact: none
keywords: [documentation, help, search, user guide, offline]
aliases: [Documentation, Help]
last_verified: "2026-08-24"
---

# Docs

## At a glance

| Use this when | Model impact | Required context |
|---|---|---|
| You need instructions or ribbon locations for a released CED command. | None. | Revit may have a project open, a family open, or no document open. |

## What it does

Opens a modeless, offline browser for the user-guide pages installed with the CED extensions. Search considers titles, keywords, extension names, ribbon paths, headings, and the full text of each indexed page.

## Steps

1. On **AE pyTools > CED Tools**, click **Docs**.
2. Enter one or more search terms.
3. Optionally filter by extension or ribbon location.
4. Choose **List** for ranked search results or **Tree** to browse the catalog by extension and ribbon location.
5. Select a result to read the page.
6. Follow related-page links or use **Back**, **Forward**, and **Home** to navigate.

## Results and verification

The selected Markdown page is rendered inside the viewer. External HTTP or HTTPS links open in the system browser; documentation pages and heading links stay inside the viewer. The Revit model is unchanged.

## Notes and limitations

- The viewer uses the documentation shipped with the installed extension release and does not require internet access.
- Unsupported raw HTML is displayed as plain text.
- Local files other than Markdown pages and referenced images are not opened automatically.
- If a page, image, catalog, or documentation root is unavailable, the viewer reports the problem without affecting Revit or pyRevit startup.

## Related pages

- [AE pyTools](index.md)
- [Selection Tools](selection.md)
- [Orientation Tools](orientation.md)
