# -*- coding: utf-8 -*-
"""Documentation-root discovery and contained local-path resolution."""

from __future__ import print_function

import os

try:
    from urllib.parse import unquote
except ImportError:  # IronPython 2 fallback
    from urllib import unquote


class DocumentationPathError(Exception):
    pass


def _inside(root, candidate):
    root = os.path.abspath(root)
    candidate = os.path.abspath(candidate)
    try:
        return os.path.commonpath([root, candidate]) == root
    except (AttributeError, ValueError):
        return candidate == root or candidate.startswith(root.rstrip(os.sep) + os.sep)


def resolve_documentation_root(start_path=None, configured_root=None):
    """Resolve the deployed user-guide root without assuming a repository path."""
    candidates = []
    if configured_root:
        candidates.append(configured_root)
    environment_root = os.getenv("CED_DOCUMENTATION_ROOT")
    if environment_root:
        candidates.append(environment_root)

    start = os.path.abspath(start_path or os.path.dirname(__file__))
    if os.path.isfile(start):
        start = os.path.dirname(start)
    current = start
    for _index in range(10):
        candidates.append(os.path.join(current, "docs", "user-guide"))
        candidates.append(os.path.join(current, "user-guide"))
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent

    checked = []
    for candidate in candidates:
        if not candidate:
            continue
        absolute = os.path.abspath(os.path.expandvars(os.path.expanduser(str(candidate))))
        if absolute in checked:
            continue
        checked.append(absolute)
        if os.path.isdir(absolute):
            return absolute
    raise DocumentationPathError(
        "The documentation root is unavailable. Checked: {}".format(", ".join(checked))
    )


def split_target(target):
    value = unquote(str(target or "").strip())
    path, separator, anchor = value.partition("#")
    return path, anchor if separator else ""


def is_external_http(target):
    lowered = str(target or "").strip().lower()
    return lowered.startswith("http://") or lowered.startswith("https://")


def has_uri_scheme(target):
    value = str(target or "").strip()
    marker = value.find(":")
    if marker <= 0:
        return False
    scheme = value[:marker]
    return scheme[0].isalpha() and all(char.isalnum() or char in "+-." for char in scheme)


def resolve_local_path(root, target_path, current_document=None, must_exist=True):
    """Resolve a decoded documentation-relative path and enforce containment."""
    root = os.path.abspath(root)
    path_text = unquote(str(target_path or "")).replace("/", os.sep)
    if os.path.isabs(path_text):
        candidate = os.path.abspath(os.path.normpath(path_text))
    else:
        base = root
        if current_document:
            current = os.path.abspath(current_document)
            base = os.path.dirname(current) if os.path.isfile(current) or os.path.splitext(current)[1] else current
        candidate = os.path.abspath(os.path.normpath(os.path.join(base, path_text)))
    if not _inside(root, candidate):
        raise DocumentationPathError("The local target escapes the documentation root.")
    if must_exist and not os.path.exists(candidate):
        raise DocumentationPathError("The local target does not exist: {}".format(target_path))
    return candidate


def relative_to_root(root, path):
    root = os.path.abspath(root)
    path = os.path.abspath(path)
    if not _inside(root, path):
        raise DocumentationPathError("The path is outside the documentation root.")
    return os.path.relpath(path, root).replace(os.sep, "/")

