---
id: renderer-compatibility-fixture
doc_type: fixture
title: Renderer Compatibility Fixture
extension: Documentation
ribbon_path: Internal > Test Fixtures
navigation_path: []
status: draft
audience: [developers]
keywords: [renderer, markdown, compatibility, fixture]
last_verified: "2026-08-24"
---

# Renderer Compatibility Fixture

This page is the regression fixture for every Markdown feature promised by the offline viewer.

## Inline formatting

Plain text can contain **bold text**, *italic text*, ***bold italic text***, and `inline code`.

## Lists

- First unordered item
- Second unordered item
  - Nested unordered item
  - Another nested item

1. First ordered item
2. Second ordered item
   1. Nested ordered item
   2. Another nested item

## Table

| Feature | Expected result |
|---|---|
| Table header | Emphasized header row |
| Table body | Bordered cells with wrapped content |

## Code block

```python
def open_help(document_id):
    return "Open {}".format(document_id)
```

## Image

![CED documentation fixture](../assets/renderer-fixture.png)

## Blockquote

> A standard blockquote is visually distinct from body text.
> It may continue on a second line.

## Alerts

> [!NOTE]
> Notes provide useful supporting context.

> [!TIP]
> Tips suggest a more effective workflow.

> [!IMPORTANT]
> Important information deserves extra attention.

> [!WARNING]
> Warnings identify a meaningful risk.

> [!CAUTION]
> Cautions identify the highest-risk conditions.

## Navigation targets

Use the [same-page target](#target-heading), open [Zoom to Selection](../ae-pytools/zoom-to-selection.md), or jump to [What it does](../ae-pytools/zoom-to-selection.md#what-it-does).

The [pyRevit website](https://www.pyrevitlabs.io/) is an external HTTP/HTTPS navigation case.

### Target heading

Same-page heading navigation should bring this block into view.
