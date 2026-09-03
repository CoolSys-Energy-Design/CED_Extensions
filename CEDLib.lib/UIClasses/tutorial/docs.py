# -*- coding: utf-8 -*-
"""Open a tool's user guide from inside the tool.

The written guides live in ``docs/user-guide/`` with one markdown file per tool
and a ``catalog.json`` index carrying each document's path, title and ribbon
path. That catalog is the source of truth here - the Docs button opens the real
guide, not a copy.

Resolution order for a tool:

1. ``DOCS_PATH`` in the tool's ``tutorial_content.py`` - an explicit
   catalog-relative path, for the cases where the folder name and the document
   slug disagree (SetPipeSystem.pushbutton vs set-piping-system.md).
2. The bundle folder name slugified against the catalog's document slugs.
3. The slug compared against document titles.
4. Failing all that, the tool's guided tour rendered as markdown, so the button
   still shows something useful rather than an error.

Output goes to the pyRevit output window, which already renders markdown and
scrolls, rather than another WPF window to maintain.
"""

import io
import json
import os
import re

from pyrevit import script

CATALOG_RELATIVE = os.path.join("docs", "user-guide", "catalog.json")
DEFAULT_CONTENT_NAME = "tutorial_content.py"
BUNDLE_SUFFIXES = ("pushbutton", "pulldown", "splitpushbutton", "stack", "panel")

_LOGGER = script.get_logger()
_CATALOG_CACHE = {}


# ---------------------------------------------------------------------------
# locating things
# ---------------------------------------------------------------------------

def _bundle_dir(caller_file):
    directory = os.path.dirname(os.path.abspath(caller_file))
    if os.path.exists(os.path.join(directory, DEFAULT_CONTENT_NAME)):
        return directory
    # A window class living in lib/ passes its own __file__.
    parent = os.path.dirname(directory)
    if os.path.exists(os.path.join(parent, DEFAULT_CONTENT_NAME)):
        return parent
    return directory


def _catalog_path(start_dir):
    current = os.path.abspath(start_dir)
    while True:
        candidate = os.path.join(current, CATALOG_RELATIVE)
        if os.path.exists(candidate):
            return candidate
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


def _load_catalog(start_dir):
    path = _catalog_path(start_dir)
    if not path:
        return None, None
    stamp = None
    try:
        stamp = os.path.getmtime(path)
    except Exception:
        pass
    cached = _CATALOG_CACHE.get(path)
    if cached and cached[0] == stamp:
        return cached[1], os.path.dirname(path)
    try:
        documents = json.load(io.open(path, encoding="utf-8-sig")).get("documents") or []
    except Exception as exc:
        _LOGGER.debug("Could not read the docs catalog: %s", exc)
        return None, None
    _CATALOG_CACHE[path] = (stamp, documents)
    return documents, os.path.dirname(path)


def slugify_bundle(folder_name):
    """'Place all Coils.pushbutton' -> 'place-all-coils'."""
    name = str(folder_name or "")
    for suffix in BUNDLE_SUFFIXES:
        if name.lower().endswith("." + suffix):
            name = name[: -(len(suffix) + 1)]
            break
    name = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "-", name)   # CamelCase -> dashes
    name = re.sub(r"[^A-Za-z0-9]+", "-", name)
    return name.strip("-").lower()


def _squash(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _explicit_docs_path(bundle):
    """DOCS_PATH declared in the tool's tutorial_content.py, if any."""
    content = os.path.join(bundle, DEFAULT_CONTENT_NAME)
    if not os.path.exists(content):
        return None
    try:
        text = io.open(content, encoding="utf-8-sig").read()
    except Exception:
        return None
    match = re.search(r'^DOCS_PATH\s*=\s*["\']([^"\']+)["\']', text, re.M)
    return match.group(1) if match else None


def resolve_document(caller_file):
    """(document dict or None, docs root or None, bundle dir)."""
    bundle = _bundle_dir(caller_file)
    documents, root = _load_catalog(bundle)
    if not documents:
        return None, root, bundle

    explicit = _explicit_docs_path(bundle)
    if explicit:
        wanted = explicit.replace("\\", "/").lower()
        for entry in documents:
            if str(entry.get("path", "")).replace("\\", "/").lower() == wanted:
                return entry, root, bundle
        _LOGGER.debug("DOCS_PATH %r is not in the catalog", explicit)

    slug = slugify_bundle(os.path.basename(bundle))
    for entry in documents:
        stem = os.path.splitext(os.path.basename(str(entry.get("path", ""))))[0]
        if stem.lower() == slug:
            return entry, root, bundle
    squashed = _squash(slug)
    for entry in documents:
        if _squash(entry.get("title")) == squashed:
            return entry, root, bundle
    return None, root, bundle


def has_docs(caller_file):
    """True when the Docs button has something real to show."""
    entry, root, bundle = resolve_document(caller_file)
    if entry is not None and root:
        return os.path.exists(os.path.join(root, str(entry.get("path", ""))))
    return os.path.exists(os.path.join(bundle, DEFAULT_CONTENT_NAME))


# ---------------------------------------------------------------------------
# rendering
# ---------------------------------------------------------------------------

def tour_markdown(content):
    """Fallback: render a tour content module as a markdown reference."""
    title = str(getattr(content, "TOUR_TITLE", "") or "Tool reference")
    for suffix in (" - Show Me Around", " - What's In Here"):
        if title.endswith(suffix):
            title = title[: -len(suffix)]
    lines = ["# {0}".format(title), ""]
    steps = list(getattr(content, "STEPS", []) or [])
    if not steps:
        lines.append("_No documentation has been written for this tool yet._")
        return "\n".join(lines)
    current = None
    for step in steps:
        chapter = step.get("chapter") or ""
        if chapter != current:
            current = chapter
            lines.extend(["", "## {0}".format(chapter)])
        lines.extend(["", "### {0}".format(step.get("title") or ""), ""])
        body = str(step.get("body") or "").strip()
        if body:
            lines.extend([body, ""])
        for bullet in list(step.get("bullets") or []):
            lines.append("- {0}".format(bullet))
        if step.get("bullets"):
            lines.append("")
    lines.extend(["", "_Rendered from this tool's guided tour - no user guide was "
                  "found for it in docs/user-guide._"])
    return "\n".join(lines)


def build_markdown(caller_file):
    """The document to show for a tool."""
    entry, root, bundle = resolve_document(caller_file)

    if entry is not None and root:
        path = os.path.join(root, str(entry.get("path", "")))
        try:
            text = io.open(path, encoding="utf-8-sig").read().strip()
        except Exception as exc:
            _LOGGER.debug("Could not read %s: %s", path, exc)
            text = ""
        if text:
            header = []
            ribbon = entry.get("ribbon_path")
            status = entry.get("status")
            verified = entry.get("last_verified")
            crumbs = [c for c in (ribbon, ("status: %s" % status) if status else None,
                                  ("last verified %s" % verified) if verified else None) if c]
            if crumbs:
                header = ["_" + "  |  ".join(str(c) for c in crumbs) + "_", ""]
            return "\n".join(header) + text

    content_path = os.path.join(bundle, DEFAULT_CONTENT_NAME)
    if os.path.exists(content_path):
        try:
            from UIClasses.tutorial import _module_name_for
            from UIClasses.tutorial.engine import load_module

            return tour_markdown(load_module(_module_name_for(content_path), content_path))
        except Exception as exc:
            _LOGGER.debug("Could not render tour docs: %s", exc)

    return ("# No documentation\n\nNo user guide was found for this tool in "
            "docs/user-guide, and it has no guided tour to fall back on.")


def show_docs(caller_file):
    """Render the tool's user guide into the pyRevit output window."""
    try:
        entry, _root, _bundle = resolve_document(caller_file)
        output = script.get_output()
        try:
            output.set_title(str((entry or {}).get("title") or "Documentation"))
        except Exception:
            pass
        output.print_md(build_markdown(caller_file))
        try:
            output.show()
        except Exception:
            pass
        return True
    except Exception as exc:
        _LOGGER.warning("Could not open documentation: %s", exc)
        return False
