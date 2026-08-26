# Documentation Viewer Execution Plan

## Goal

Provide an offline, searchable in-Revit documentation browser for the Markdown pages in the user guide without duplicating documentation content inside individual tool bundles.

The viewer should support both centralized documentation discovery and direct navigation between related documentation pages.

## Scope and Source of Truth

- User-guide Markdown under `docs/user-guide/` remains the editable source of truth.
- Every tool page has stable frontmatter:
  - `id`
  - `title`
  - `extension`
  - `ribbon_path`
  - `status`
  - `audience`
  - `keywords`
  - `last_verified`
- Existing bundle-local Markdown remains developer reference material and is not read as viewer content.
- The first viewer release covers only current, user-facing extensions. Legacy or disabled extensions are added only when their pages are ready.
- `catalog.json` and other generated search data are build artifacts and are never manually maintained.

## Implementation Decisions

- **Host:** Build the initial browser as a modeless WPF window. A dockable pane is not part of the first release.
- **Command location:** Choose and document the exact extension, tab, panel, and button name before implementing the command. No documentation-viewer button is created as part of documentation work.
- **Supported environment:** The implementation must record the exact Revit and pyRevit versions tested before release. It must work with the Python engine and .NET capabilities supported by those environments.
- **UI foundation:** Reuse `CEDLib.lib/UIClasses` resource loading, theme bridging, and CED XAML resources. Add viewer-only presentation rules in a dedicated `DocumentationStyles.xaml` resource dictionary; do not create a separate theme system.
- **Renderer:** Implement and test a native WPF `FlowDocument` renderer first. It must pass the renderer compatibility fixture before release. Do not add WebView2 or another browser/third-party dependency in the initial implementation. Reconsider an HTML/WebView renderer only if the FlowDocument approach cannot meet the documented Markdown standard in supported Revit environments.
- **Markdown safety:** Raw HTML is not supported. Render only the documented portable GFM subset and treat unsupported constructs as plain text or a non-blocking viewer warning.
- **Runtime data:** The viewer reads generated `catalog.json` for search and metadata. It does not generate a catalog or scan the entire Markdown source library during normal use.
- **Failure behavior:** If the documentation root, catalog, page, image, link, or frontmatter is invalid or unavailable, show a useful viewer-level message with retry/close actions. Never rebuild automatically at runtime, and never allow documentation failure to interfere with pyRevit or Revit startup.
- **Build tooling:** Add named documentation validation and catalog-generation commands/scripts to the repository and run them as part of the documentation or extension release workflow.
- **Documentation deployment:** The deployed package must retain `docs/user-guide/` in a location resolvable by the documentation-root abstraction.
- **External links:** Open only HTTP/HTTPS links in the system browser. Do not automatically open arbitrary local files.

## Markdown Standard

Documentation should target a portable GitHub-Flavored Markdown (GFM) subset rather than platform-specific Markdown features.

The supported documentation syntax should include:

- Headings
- Bold and italic text
- Ordered and unordered lists
- Nested lists
- Blockquotes
- Fenced code blocks
- Inline code
- Tables
- Images
- Relative links between Markdown files
- Links to headings within the same page
- Links to headings in another Markdown page
- Standard external HTTP/HTTPS links
- GitHub-style alerts:
  - `NOTE`
  - `TIP`
  - `IMPORTANT`
  - `WARNING`
  - `CAUTION`
- YAML frontmatter for page metadata

Prefer standard relative Markdown links:

```md
[Wire Tools](wire-tools.md)
[Individual Homeruns](wire-tools.md#individual-homeruns)
```

over platform-specific syntax such as Obsidian `[[wikilinks]]`.

Obsidian may still be used as an authoring environment, but documentation should not depend on Obsidian-specific features unless support is intentionally added to the documentation standard later.

## Delivery Steps

### 1. Complete the Documentation Library

Add and review current user-facing command pages in small extension/panel batches.

Each released user-facing command should have a corresponding user-guide page with valid frontmatter and a stable document ID.

### 2. Add Documentation Validation

Create a documentation-only validation script.

Validation should check:

- Required frontmatter fields
- Duplicate document IDs
- Invalid or unsupported frontmatter values
- Broken relative Markdown links
- Broken same-page and cross-page heading links
- Missing referenced images/assets
- Invalid or unsupported alert types
- Links or asset paths that improperly escape the documentation root
- Missing expected user-facing command pages
- Orphaned or otherwise invalid documentation entries

Validation should operate against the Markdown source rather than generated catalog data whenever possible.

### 3. Generate the Documentation Catalog and Search Index

Generate `catalog.json` from the Markdown files during the documentation build/release process.

Do not hand-maintain both page metadata and catalog metadata.

The generated catalog should contain enough information to support useful full-text search, including:

- Document ID
- Title
- Extension
- Ribbon path
- Status
- Audience
- Keywords
- Last verified date
- Relative Markdown file path
- Page headings
- Normalized searchable document text

Search should therefore be able to find a document based on text contained in its content even when the term was not manually added to `keywords`.

The initial implementation can load and search this generated JSON in memory. A database or dedicated search engine is not required for the expected documentation library size.

### 4. Define the Documentation Root

The viewer should operate against a resolved **documentation root** rather than being tightly coupled to a particular repository path.

For the initial deployment, the documentation root may resolve to:

```text
<repository>/docs/user-guide/
```

The viewer should then expect resources such as:

```text
docs/user-guide/
├── catalog.json
├── electrical/
├── mechanical/
├── general/
└── assets/
```

This abstraction should allow the documentation to be packaged or deployed differently later without requiring major changes to the viewer.

### 5. Build a Renderer Compatibility Fixture

Create a dedicated Markdown test page containing every feature guaranteed by the documentation standard.

The fixture should include:

- Multiple heading levels
- Bold and italic text
- Ordered and unordered lists
- Nested lists
- Tables
- Inline code
- Fenced code blocks
- Images
- Blockquotes
- `NOTE`
- `TIP`
- `IMPORTANT`
- `WARNING`
- `CAUTION`
- Same-page heading links
- Relative Markdown links
- Relative Markdown links with heading anchors
- External links

Keep this fixture in the repository as a regression test for future renderer or Revit-version changes.

### 6. Validate the FlowDocument Rendering Implementation

Implement the native WPF `FlowDocument` rendering path and test it against the renderer compatibility fixture in all supported Revit environments.

Evaluate:

- Offline operation
- Deployment complexity
- Markdown feature support
- GitHub-style alert rendering
- Table rendering
- Image rendering
- Code formatting
- Internal link handling
- Heading navigation
- Performance
- Compatibility across supported Revit versions

If the FlowDocument renderer cannot reliably support the defined Markdown standard, document the specific failing fixture cases before proposing an HTML/WebView alternative.

Do not require internet access for rendering.

### 7. Implement the Documentation Browser

Add one lightweight pyRevit documentation command.

The viewer should provide:

- Search box
- Search results
- Filtering by extension
- Filtering by ribbon path or tool area
- Tool name in search results
- Extension in search results
- Ribbon location in search results
- Rendered Markdown content
- Back navigation
- Forward navigation
- Home/search navigation

Search should consider:

- Title
- Keywords
- Extension
- Ribbon path
- Headings
- Indexed document content

### 8. Implement Internal Navigation

Treat documentation as a connected collection of pages rather than isolated files.

The viewer should understand:

```text
tool.md
tool.md#heading
#heading
../other-folder/tool.md
../other-folder/tool.md#heading
```

Navigation behavior should be:

- Relative `.md` link → open the target document in the viewer
- Relative `.md#anchor` link → open the document and navigate to the heading
- `#anchor` → navigate within the current document
- Local image → load from the documentation root
- HTTP/HTTPS link → open in the user's normal external browser

Back and Forward should preserve page navigation history.

Resolved local documentation and asset paths must remain inside the documentation root.

Unsupported local file types should not be opened automatically.

### 9. Test Deployed Behavior and Failure Handling

Test at minimum:

- Normal installed extension
- Missing catalog
- Outdated/stale catalog
- Missing Markdown page
- Missing image
- Broken relative link
- Broken heading link
- Invalid Markdown/frontmatter
- Search with no results
- Documentation root unavailable

Failures should produce useful viewer-level messages.

Documentation failures must never prevent normal pyRevit or extension startup.

### 10. Integrate With the Release Workflow

Generate the catalog/search index as part of the documentation or extension release workflow.

Require documentation validation whenever:

- A user-facing command is added
- A user-facing command is removed
- A command is renamed or moved
- Its ribbon location changes
- Its documentation is materially changed

Release validation should prevent publication when required documentation is missing or when duplicate IDs, invalid metadata, broken internal links, or missing required assets are detected.

## Future Integration

### Contextual Tool Help

Stable documentation IDs should allow individual pyRevit tools to open their documentation directly.

For example, a future shared helper could conceptually support:

```text
open_help("wire-tools")
```

Individual tool dialogs could then provide a Help or `?` button that opens the correct page in the same documentation viewer.

This should reuse the central viewer rather than implementing documentation rendering separately in each tool.

### Additional Viewer Features

Potential later enhancements include:

- Recently viewed pages
- Favorites
- Related-tool sections
- Search-term highlighting
- Table of contents generated from headings
- Breadcrumb navigation
- Keyboard navigation
- Copy-link/document-ID actions
- Direct navigation from pyRevit ribbon commands
- Documentation version/update indicators

These are not required for the initial release.

## Definition of Done

The first release is complete when:

- A user can search and open all released user-facing command pages offline.
- Full-text search can find relevant pages without relying solely on manually assigned keywords.
- Search results show the tool name, extension, and ribbon location.
- Markdown renders consistently according to the defined documentation standard.
- GitHub-style alerts render correctly.
- Local images render correctly.
- Users can follow relative links between documentation pages.
- Users can navigate to headings within the same or another page.
- Back and Forward navigation work between visited pages.
- External web links open outside the viewer.
- The viewer contains no duplicated tool prose and reads from the documentation root.
- Viewer/documentation failures cannot interfere with normal pyRevit startup.
- Release validation prevents missing required pages, duplicate IDs, broken internal links, and missing referenced assets.
