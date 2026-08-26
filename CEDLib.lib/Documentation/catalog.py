# -*- coding: utf-8 -*-
"""Read-only generated documentation catalog and in-memory search."""

from __future__ import print_function

import io
import json
import os
import re

from Documentation.pathing import resolve_local_path


class CatalogError(Exception):
    pass


def _text(value):
    return str(value or "").strip()


def _list(value):
    if isinstance(value, list):
        return [_text(item) for item in value if _text(item)]
    return [_text(value)] if _text(value) else []


class Catalog(object):
    REQUIRED = (
        "id",
        "title",
        "extension",
        "ribbon_path",
        "path",
        "headings",
        "text",
        "source_sha256",
    )

    def __init__(self, root, data):
        self.root = os.path.abspath(root)
        if not isinstance(data, dict) or data.get("schema_version") != 1:
            raise CatalogError("The documentation catalog schema is unsupported.")
        raw_documents = data.get("documents")
        if not isinstance(raw_documents, list):
            raise CatalogError("The documentation catalog has no document list.")
        self.documents = []
        self.by_id = {}
        self.by_path = {}
        for raw in raw_documents:
            if not isinstance(raw, dict):
                raise CatalogError("The documentation catalog contains an invalid entry.")
            missing = [field for field in self.REQUIRED if field not in raw]
            if missing:
                raise CatalogError("A catalog entry is missing: {}".format(", ".join(missing)))
            item = dict(raw)
            item["id"] = _text(item["id"])
            item["title"] = _text(item["title"])
            item["extension"] = _text(item["extension"])
            item["ribbon_path"] = _text(item["ribbon_path"])
            item["ribbon_location"] = _ribbon_location(item)
            item["doc_type"] = _text(item.get("doc_type") or "tool").lower()
            item["navigation_path"] = _list(item.get("navigation_path"))
            item["path"] = _text(item["path"]).replace("\\", "/")
            item["keywords"] = _list(item.get("keywords"))
            item["audience"] = _list(item.get("audience"))
            item["headings"] = item.get("headings") or []
            item["text"] = _text(item.get("text"))
            item["source_sha256"] = _text(item.get("source_sha256")).lower()
            if not item["id"] or item["id"] in self.by_id:
                raise CatalogError("The documentation catalog contains a missing or duplicate id.")
            try:
                resolve_local_path(self.root, item["path"], must_exist=False)
            except Exception as error:
                raise CatalogError("Unsafe catalog path for '{}': {}".format(item["id"], error))
            self.documents.append(item)
            self.by_id[item["id"]] = item
            self.by_path[item["path"].lower()] = item

    @classmethod
    def load(cls, root):
        path = os.path.join(root, "catalog.json")
        if not os.path.isfile(path):
            raise CatalogError("The documentation catalog is missing: {}".format(path))
        try:
            with io.open(path, "r", encoding="utf-8-sig") as stream:
                data = json.load(stream)
        except Exception as error:
            raise CatalogError("The documentation catalog could not be read: {}".format(error))
        return cls(root, data)

    @property
    def extensions(self):
        return sorted(set(item["extension"] for item in self.documents), key=lambda value: value.lower())

    @property
    def ribbon_paths(self):
        return sorted(set(item["ribbon_location"] for item in self.documents), key=lambda value: value.lower())

    def get_by_path(self, relative_path):
        return self.by_path.get(str(relative_path or "").replace("\\", "/").lower())

    def search(self, query="", extension="", ribbon_path=""):
        query = _text(query).lower()
        terms = [item for item in re.split(r"\s+", query) if item]
        extension = _text(extension).lower()
        ribbon_path = _text(ribbon_path).lower()
        matches = []
        for item in self.documents:
            if extension and item["extension"].lower() != extension:
                continue
            if ribbon_path and item["ribbon_location"].lower() != ribbon_path:
                continue
            if not terms:
                matches.append((0, item["title"].lower(), item))
                continue
            cached = item.get("_search_fields")
            if cached is None:
                title = item["title"].lower()
                keywords = " ".join(item["keywords"]).lower()
                headings = " ".join(
                    _text(value.get("title")) for value in item["headings"]
                ).lower()
                metadata = " ".join(
                    (item["extension"], item["ribbon_path"], item["ribbon_location"], item["id"])
                ).lower()
                content = item["text"].lower()
                cached = (title, keywords, headings, metadata, content)
                item["_search_fields"] = cached
            title, keywords, headings, metadata, content = cached
            searchable = " ".join((title, keywords, headings, metadata, content))
            if terms and not all(term in searchable for term in terms):
                continue
            score = 0
            if query:
                if title == query:
                    score += 10000
                elif title.startswith(query):
                    score += 5000
                elif query in title:
                    score += 2500
                for term in terms:
                    score += 300 if term in title else 0
                    score += 60 if term in keywords else 0
                    score += 40 if term in headings else 0
                    score += 20 if term in metadata else 0
                    score += 5 if term in content else 0
            matches.append((score, item["title"].lower(), item))
        matches.sort(key=lambda value: (-value[0], value[1]))
        return [value[2] for value in matches]


def build_navigation_tree(documents, index_documents=None):
    """Build extension/navigation hierarchy independent of folders and ribbon layout."""
    root = {"groups": {}, "documents": [], "index": None}
    for document in list(documents or []):
        extension = _text(document.get("extension"))
        segments = [extension] if extension else []
        segments.extend(_list(document.get("navigation_path")))
        cursor = root
        for segment in segments:
            cursor = cursor["groups"].setdefault(
                segment,
                {"groups": {}, "documents": [], "index": None},
            )
        if _text(document.get("doc_type")).lower() == "index":
            cursor["index"] = document
        else:
            cursor["documents"].append(document)

    for document in list(index_documents or []):
        if _text(document.get("doc_type")).lower() != "index":
            continue
        extension = _text(document.get("extension"))
        segments = [extension] if extension else []
        segments.extend(_list(document.get("navigation_path")))
        cursor = root
        for segment in segments:
            cursor = cursor["groups"].get(segment)
            if cursor is None:
                break
        if cursor is not None and cursor.get("index") is None:
            cursor["index"] = document

    def materialize(node):
        groups = []
        for label in sorted(node["groups"], key=lambda value: value.lower()):
            child = node["groups"][label]
            groups.append(
                {
                    "label": label,
                    "groups": materialize(child),
                    "documents": sorted(child["documents"], key=lambda item: item["title"].lower()),
                    "index": child.get("index"),
                }
            )
        return groups

    return {
        "groups": materialize(root),
        "documents": sorted(root["documents"], key=lambda item: item["title"].lower()),
    }


def build_ribbon_tree(documents, index_documents=None):
    """Backward-compatible name for the metadata-driven navigation tree."""
    return build_navigation_tree(documents, index_documents=index_documents)


def _ribbon_location(item):
    """Read generated location metadata or derive it for older catalogs."""
    generated = _text(item.get("ribbon_location"))
    if generated:
        return generated
    raw = _text(item.get("ribbon_path"))
    segments = [segment.strip() for segment in raw.split(">") if segment.strip()]
    if len(segments) < 2:
        return raw

    def comparable(value):
        return re.sub(r"[^a-z0-9]+", " ", _text(value).lower()).strip()

    command_names = {comparable(item.get("title"))}
    command_names.update(comparable(value) for value in _list(item.get("aliases")))
    if comparable(segments[-1]) in command_names:
        segments = segments[:-1]
    return " > ".join(segments) if segments else raw
