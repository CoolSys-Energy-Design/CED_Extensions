# -*- coding: utf-8 -*-
"""Documentation validation and catalog-generation support.

This module intentionally uses only the Python standard library so the
documentation release checks can run on a clean workstation.
"""

from __future__ import print_function

import argparse
import hashlib
import io
import json
import os
import re

REQUIRED_FIELDS = (
    "id",
    "title",
    "extension",
    "ribbon_path",
    "status",
    "audience",
    "keywords",
    "last_verified",
)
ALLOWED_STATUS = {"production", "beta", "draft", "deprecated"}
ALLOWED_ALERTS = {"NOTE", "TIP", "IMPORTANT", "WARNING", "CAUTION"}
ALLOWED_DOC_TYPES = {"tool", "guide", "index", "fixture"}
NON_CONTENT_NAMES = {"_tool-template.md", "viewer-execution-plan.md", "AGENTS.md"}
LINK_RE = re.compile(r"(!?)\[([^\]]*)\]\(([^)]+)\)")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*#*\s*$")
ALERT_RE = re.compile(r"^\s*>\s*\[!([^\]]+)\]", re.IGNORECASE)
HTML_RE = re.compile(r"^\s*</?[A-Za-z][^>]*>\s*$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ID_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


class Issue(object):
    def __init__(self, severity, path, message, line=None):
        self.severity = severity
        self.path = path
        self.message = message
        self.line = line

    def format(self):
        location = self.path
        if self.line:
            location = "{}:{}".format(location, self.line)
        return "{} {}: {}".format(self.severity.upper(), location, self.message)


def read_text(path):
    with io.open(path, "r", encoding="utf-8-sig") as stream:
        return stream.read()


def write_json(path, value):
    with io.open(path, "w", encoding="utf-8", newline="\n") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2, sort_keys=True)
        stream.write("\n")


def _split_inline_list(value):
    inner = value.strip()[1:-1].strip()
    if not inner:
        return []
    values = []
    current = []
    quote = None
    for char in inner:
        if char in ("'", '"'):
            if quote == char:
                quote = None
            elif quote is None:
                quote = char
            else:
                current.append(char)
        elif char == "," and quote is None:
            values.append(_parse_scalar("".join(current).strip()))
            current = []
        else:
            current.append(char)
    values.append(_parse_scalar("".join(current).strip()))
    return values


def _parse_scalar(value):
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    if value.startswith("[") and value.endswith("]"):
        return _split_inline_list(value)
    lowered = value.lower()
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    if lowered in ("null", "none", "~"):
        return None
    return value


def split_frontmatter(text):
    """Return ``(metadata, body, error)`` for a Markdown document."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines or lines[0].strip() != "---":
        return None, text, None
    end = None
    for index in range(1, len(lines)):
        if lines[index].strip() == "---":
            end = index
            break
    if end is None:
        return None, text, "frontmatter block is not closed"

    metadata = {}
    active_list_key = None
    for line_number, raw in enumerate(lines[1:end], start=2):
        stripped = raw.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped.startswith("-") and active_list_key:
            metadata[active_list_key].append(_parse_scalar(stripped[1:].strip()))
            continue
        if ":" not in raw:
            return None, text, "invalid frontmatter at line {}".format(line_number)
        key, value = raw.split(":", 1)
        key = key.strip()
        value = value.strip()
        if not key:
            return None, text, "empty frontmatter key at line {}".format(line_number)
        if key in metadata:
            return None, text, "duplicate frontmatter key '{}' at line {}".format(key, line_number)
        if value:
            metadata[key] = _parse_scalar(value)
            active_list_key = None
        else:
            metadata[key] = []
            active_list_key = key
    return metadata, "\n".join(lines[end + 1 :]), None


def normalize_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value).strip()
    return [text] if text else []


def strip_inline_markdown(value):
    text = re.sub(r"!\[([^\]]*)\]\([^)]+\)", r"\1", value)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = re.sub(r"[`*_~]", "", text)
    return text.strip()


def github_slug(value):
    text = strip_inline_markdown(value).lower().strip()
    text = re.sub(r"[^\w\- ]", "", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "-", text)
    return text


def extract_headings(body):
    headings = []
    counts = {}
    in_fence = False
    for line_number, line in enumerate(body.splitlines(), start=1):
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        match = HEADING_RE.match(line)
        if not match:
            continue
        title = strip_inline_markdown(match.group(2))
        base = github_slug(title)
        count = counts.get(base, 0)
        counts[base] = count + 1
        anchor = base if count == 0 else "{}-{}".format(base, count)
        headings.append(
            {"level": len(match.group(1)), "title": title, "anchor": anchor, "line": line_number}
        )
    return headings


def searchable_text(body):
    lines = []
    in_fence = False
    for line in body.splitlines():
        stripped = line.strip()
        if stripped.startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            lines.append(stripped)
            continue
        stripped = re.sub(r"^#{1,6}\s+", "", stripped)
        stripped = re.sub(r"^\s*(?:[-+*]|\d+[.)])\s+", "", stripped)
        stripped = re.sub(r"^\s*>\s*(?:\[![A-Z]+\])?\s*", "", stripped)
        stripped = re.sub(r"!\[([^\]]*)\]\([^)]+\)", r"\1", stripped)
        stripped = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", stripped)
        stripped = re.sub(r"[`*_~|]", " ", stripped)
        stripped = re.sub(r"\s+", " ", stripped).strip()
        if stripped and not re.match(r"^[-: ]+$", stripped):
            lines.append(stripped)
    return re.sub(r"\s+", " ", " ".join(lines)).strip()


def ribbon_location(metadata, relative_path):
    """Return a command's containing ribbon location without hardcoded panels."""
    raw = str(metadata.get("ribbon_path", "")).strip()
    segments = [segment.strip() for segment in raw.split(">") if segment.strip()]
    name = os.path.basename(relative_path).lower()
    doc_type = str(metadata.get("doc_type", "index" if name == "readme.md" else "tool")).lower()
    if doc_type != "tool" or len(segments) < 2:
        return raw

    def comparable(value):
        return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()

    command_names = {comparable(metadata.get("title"))}
    command_names.update(comparable(value) for value in normalize_list(metadata.get("aliases")))
    if comparable(segments[-1]) in command_names:
        segments = segments[:-1]
    return " > ".join(segments) if segments else raw


def _target_without_title(raw):
    value = raw.strip()
    if value.startswith("<") and ">" in value:
        return value[1 : value.index(">")]
    # Standard Markdown permits an optional quoted title after the target.
    match = re.match(r"^(\S+)(?:\s+[\"'].*[\"'])?$", value)
    return match.group(1) if match else value


def iter_links(body):
    in_fence = False
    for line_number, line in enumerate(body.splitlines(), start=1):
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        for match in LINK_RE.finditer(line):
            yield {
                "image": bool(match.group(1)),
                "label": match.group(2),
                "target": _target_without_title(match.group(3)),
                "line": line_number,
            }


def _is_inside(root, candidate):
    try:
        return os.path.commonpath([os.path.abspath(root), os.path.abspath(candidate)]) == os.path.abspath(root)
    except (AttributeError, ValueError):
        root_prefix = os.path.abspath(root).rstrip(os.sep) + os.sep
        return os.path.abspath(candidate).startswith(root_prefix)


def _relative(root, path):
    return os.path.relpath(path, root).replace(os.sep, "/")


def discover_markdown(root):
    result = []
    for current, directories, files in os.walk(root):
        directories[:] = sorted(item for item in directories if not item.startswith("."))
        for name in sorted(files):
            if name.lower().endswith(".md"):
                result.append(os.path.join(current, name))
    return result


def load_documents(root):
    documents = {}
    issues = []
    for path in discover_markdown(root):
        rel = _relative(root, path)
        text = read_text(path)
        metadata, body, error = split_frontmatter(text)
        if error:
            issues.append(Issue("error", rel, error))
            continue
        documents[rel] = {
            "path": path,
            "relative_path": rel,
            "metadata": metadata,
            "body": body,
            "headings": extract_headings(body),
            "source_text": text,
        }
    return documents, issues


def _validate_metadata(document, issues):
    rel = document["relative_path"]
    metadata = document["metadata"]
    name = os.path.basename(rel)
    if name in NON_CONTENT_NAMES:
        return
    if metadata is None:
        if name not in NON_CONTENT_NAMES and name.lower() != "readme.md":
            issues.append(Issue("error", rel, "content page is missing YAML frontmatter"))
        return
    for field in REQUIRED_FIELDS:
        if field not in metadata or metadata[field] in (None, "", []):
            issues.append(Issue("error", rel, "required frontmatter field '{}' is missing".format(field)))
    doc_id = str(metadata.get("id", ""))
    if doc_id and not ID_RE.match(doc_id):
        issues.append(Issue("error", rel, "id must be lowercase kebab-case"))
    status = str(metadata.get("status", "")).lower()
    if status and status not in ALLOWED_STATUS:
        issues.append(Issue("error", rel, "unsupported status '{}'".format(status)))
    doc_type = str(metadata.get("doc_type", "tool" if name.lower() != "readme.md" else "index")).lower()
    if doc_type not in ALLOWED_DOC_TYPES:
        issues.append(Issue("error", rel, "unsupported doc_type '{}'".format(doc_type)))
    date = str(metadata.get("last_verified", ""))
    if date and not DATE_RE.match(date):
        issues.append(Issue("error", rel, "last_verified must use YYYY-MM-DD"))
    if not normalize_list(metadata.get("audience")):
        issues.append(Issue("error", rel, "audience must contain at least one value"))
    if not isinstance(metadata.get("keywords"), list):
        issues.append(Issue("error", rel, "keywords must be a YAML inline or block list"))
    if "navigation_path" not in metadata:
        issues.append(Issue("error", rel, "required frontmatter field 'navigation_path' is missing"))
    elif not isinstance(metadata.get("navigation_path"), list):
        issues.append(Issue("error", rel, "navigation_path must be a YAML inline or block list"))


def _validate_markup(document, documents, root, issues):
    rel = document["relative_path"]
    body = document["body"]
    current_dir = os.path.dirname(document["path"])
    heading_anchors = {item["anchor"] for item in document["headings"]}
    in_fence = False
    for line_number, line in enumerate(body.splitlines(), start=1):
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        alert = ALERT_RE.match(line)
        if alert and alert.group(1).upper() not in ALLOWED_ALERTS:
            issues.append(
                Issue("error", rel, "unsupported alert type '{}'".format(alert.group(1)), line_number)
            )
        if HTML_RE.match(line):
            issues.append(Issue("warning", rel, "raw HTML renders as plain text", line_number))

    for link in iter_links(body):
        target = link["target"]
        if not target or target.startswith(("http://", "https://", "mailto:")):
            continue
        if re.match(r"^[A-Za-z][A-Za-z0-9+.-]*:", target):
            issues.append(Issue("error", rel, "unsupported link scheme in '{}'".format(target), link["line"]))
            continue
        path_part, separator, anchor = target.partition("#")
        if not path_part:
            if link["image"]:
                issues.append(Issue("error", rel, "image target cannot be a heading anchor", link["line"]))
            elif anchor and anchor not in heading_anchors:
                issues.append(Issue("error", rel, "missing heading '#{}'".format(anchor), link["line"]))
            continue
        resolved = os.path.abspath(os.path.normpath(os.path.join(current_dir, path_part.replace("/", os.sep))))
        if not _is_inside(root, resolved):
            issues.append(Issue("error", rel, "target escapes documentation root: '{}'".format(target), link["line"]))
            continue
        if not os.path.isfile(resolved):
            noun = "image/asset" if link["image"] else "link target"
            issues.append(Issue("error", rel, "missing {} '{}'".format(noun, target), link["line"]))
            continue
        if link["image"]:
            continue
        extension = os.path.splitext(path_part)[1].lower()
        if extension and extension != ".md":
            issues.append(Issue("error", rel, "unsupported local file link '{}'".format(target), link["line"]))
            continue
        target_rel = _relative(root, resolved)
        target_document = documents.get(target_rel)
        if anchor and target_document:
            anchors = {item["anchor"] for item in target_document["headings"]}
            if anchor not in anchors:
                issues.append(Issue("error", rel, "missing heading '{}' in '{}'".format(anchor, path_part), link["line"]))


def _load_release_manifest(root, repository_root, documents, issues):
    path = os.path.join(root, "release-manifest.json")
    if not os.path.isfile(path):
        issues.append(Issue("error", "release-manifest.json", "release manifest is missing"))
        return
    try:
        manifest = json.loads(read_text(path))
    except Exception as error:
        issues.append(Issue("error", "release-manifest.json", "invalid JSON: {}".format(error)))
        return
    by_id = {}
    for document in documents.values():
        metadata = document["metadata"] or {}
        if metadata.get("id"):
            by_id[str(metadata["id"])] = document
    seen = set()
    manifest_bundle_paths = set()
    for entry in manifest.get("commands", []):
        doc_id = str(entry.get("document_id", ""))
        bundle = str(entry.get("bundle_path", ""))
        if not doc_id or not bundle:
            issues.append(Issue("error", "release-manifest.json", "every command needs document_id and bundle_path"))
            continue
        seen.add(doc_id)
        if doc_id not in by_id:
            issues.append(Issue("error", "release-manifest.json", "released command '{}' has no documentation page".format(doc_id)))
        bundle_path = os.path.abspath(os.path.join(repository_root, bundle.replace("/", os.sep)))
        normalized_bundle_path = os.path.normcase(bundle_path)
        if normalized_bundle_path in manifest_bundle_paths:
            issues.append(
                Issue("error", "release-manifest.json", "duplicate command bundle_path '{}'".format(bundle))
            )
        manifest_bundle_paths.add(normalized_bundle_path)
        if not _is_inside(repository_root, bundle_path) or not os.path.isdir(bundle_path):
            issues.append(Issue("error", "release-manifest.json", "bundle does not exist for '{}': {}".format(doc_id, bundle)))

    command_suffixes = (".pushbutton", ".smartbutton", ".urlbutton")
    for bundle_root in manifest.get("released_bundle_roots", []):
        absolute_root = os.path.abspath(os.path.join(repository_root, str(bundle_root).replace("/", os.sep)))
        if not _is_inside(repository_root, absolute_root) or not os.path.isdir(absolute_root):
            issues.append(
                Issue("error", "release-manifest.json", "released bundle root does not exist: {}".format(bundle_root))
            )
            continue
        for current, directories, _files in os.walk(absolute_root):
            for directory in list(directories):
                if directory.lower().endswith(command_suffixes):
                    command_path = os.path.normcase(os.path.abspath(os.path.join(current, directory)))
                    if command_path not in manifest_bundle_paths:
                        issues.append(
                            Issue(
                                "error",
                                "release-manifest.json",
                                "released command bundle has no manifest documentation entry: {}".format(
                                    _relative(repository_root, command_path)
                                ),
                            )
                        )

    for document in documents.values():
        metadata = document["metadata"] or {}
        name = os.path.basename(document["relative_path"]).lower()
        if os.path.basename(document["relative_path"]) in NON_CONTENT_NAMES:
            continue
        doc_type = str(metadata.get("doc_type", "index" if name == "readme.md" else "tool")).lower()
        if (
            doc_type == "tool"
            and str(metadata.get("status", "")).lower() == "production"
            and str(metadata.get("id", "")) not in seen
        ):
            issues.append(
                Issue(
                    "error",
                    document["relative_path"],
                    "production tool page is orphaned from release-manifest.json",
                )
            )


def validate(root, repository_root=None):
    root = os.path.abspath(root)
    repository_root = os.path.abspath(repository_root or os.path.join(root, "..", ".."))
    documents, issues = load_documents(root)
    ids = {}
    for document in documents.values():
        _validate_metadata(document, issues)
        metadata = document["metadata"] or {}
        doc_id = metadata.get("id")
        if doc_id:
            if doc_id in ids:
                issues.append(
                    Issue("error", document["relative_path"], "duplicate id '{}' also used by {}".format(doc_id, ids[doc_id]))
                )
            else:
                ids[doc_id] = document["relative_path"]
    for document in documents.values():
        _validate_markup(document, documents, root, issues)
    _load_release_manifest(root, repository_root, documents, issues)
    return documents, sorted(issues, key=lambda item: (item.path, item.line or 0, item.message))


def build_catalog(root, repository_root=None):
    documents, issues = validate(root, repository_root=repository_root)
    errors = [item for item in issues if item.severity == "error"]
    if errors:
        raise ValueError("documentation validation failed with {} error(s)".format(len(errors)))
    entries = []
    for rel, document in sorted(documents.items()):
        metadata = document["metadata"]
        if not metadata:
            continue
        if os.path.basename(rel) in NON_CONTENT_NAMES:
            continue
        if str(metadata.get("doc_type", "")).lower() == "fixture":
            continue
        entry = {
            "id": str(metadata["id"]),
            "title": str(metadata["title"]),
            "extension": str(metadata["extension"]),
            "ribbon_path": str(metadata["ribbon_path"]),
            "ribbon_location": ribbon_location(metadata, rel),
            "doc_type": str(
                metadata.get(
                    "doc_type",
                    "index" if os.path.basename(rel).lower() in ("readme.md", "index.md") else "tool",
                )
            ).lower(),
            "navigation_path": normalize_list(metadata.get("navigation_path")),
            "status": str(metadata["status"]),
            "audience": normalize_list(metadata["audience"]),
            "keywords": normalize_list(metadata["keywords"]),
            "aliases": normalize_list(metadata.get("aliases")),
            "last_verified": str(metadata["last_verified"]),
            "path": rel,
            "headings": [
                {"level": item["level"], "title": item["title"], "anchor": item["anchor"]}
                for item in document["headings"]
            ],
            "text": searchable_text(document["body"]),
            "source_sha256": hashlib.sha256(document["source_text"].encode("utf-8")).hexdigest(),
        }
        entries.append(entry)
    return {"schema_version": 1, "documents": entries}, issues


def default_paths():
    script_dir = os.path.abspath(os.path.dirname(__file__))
    repository_root = os.path.abspath(os.path.join(script_dir, "..", ".."))
    documentation_root = os.path.join(repository_root, "docs", "user-guide")
    return repository_root, documentation_root


def run_validate(argv=None):
    repository_root, documentation_root = default_paths()
    parser = argparse.ArgumentParser(description="Validate the CED user-guide source.")
    parser.add_argument("--root", default=documentation_root)
    args = parser.parse_args(argv)
    documents, issues = validate(args.root, repository_root=repository_root)
    for issue in issues:
        print(issue.format())
    errors = [item for item in issues if item.severity == "error"]
    print("Validated {} Markdown files: {} error(s), {} warning(s).".format(
        len(documents), len(errors), len(issues) - len(errors)
    ))
    return 1 if errors else 0


def run_generate(argv=None):
    repository_root, documentation_root = default_paths()
    parser = argparse.ArgumentParser(description="Generate the CED documentation search catalog.")
    parser.add_argument("--root", default=documentation_root)
    parser.add_argument("--output", default=os.path.join(documentation_root, "catalog.json"))
    args = parser.parse_args(argv)
    try:
        catalog, issues = build_catalog(args.root, repository_root=repository_root)
    except ValueError:
        _documents, issues = validate(args.root, repository_root=repository_root)
        for issue in issues:
            print(issue.format())
        return 1
    write_json(args.output, catalog)
    print("Generated {} with {} documents.".format(args.output, len(catalog["documents"])))
    return 0
