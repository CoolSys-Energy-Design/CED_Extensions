# -*- coding: utf-8 -*-
"""Dependency-free parser for the documented portable GFM subset."""

from __future__ import print_function

import re

HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*#*\s*$")
LIST_RE = re.compile(r"^(\s*)([-+*]|\d+[.)])\s+(.+)$")
TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?\s*$")
ALERT_RE = re.compile(r"^\s*>\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$", re.IGNORECASE)
HTML_RE = re.compile(r"^\s*</?[A-Za-z][^>]*>\s*$")
REQUIRED_FRONTMATTER = (
    "id",
    "title",
    "extension",
    "ribbon_path",
    "status",
    "audience",
    "keywords",
    "last_verified",
)


def runtime_frontmatter(text):
    """Return minimal page metadata or raise a useful runtime error."""
    normalized = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    if not lines or lines[0].strip() != "---":
        raise ValueError("The documentation page is missing YAML frontmatter.")
    end = None
    for index in range(1, len(lines)):
        if lines[index].strip() == "---":
            end = index
            break
    if end is None:
        raise ValueError("The documentation page has an unclosed frontmatter block.")
    metadata = {}
    for line_number, line in enumerate(lines[1:end], start=2):
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or stripped.startswith("-"):
            continue
        if ":" not in line:
            raise ValueError("Invalid frontmatter at line {}.".format(line_number))
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip()
        if key in metadata:
            raise ValueError("Duplicate frontmatter field '{}'.".format(key))
        if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
            value = value[1:-1]
        metadata[key] = value
    missing = [field for field in REQUIRED_FRONTMATTER if field not in metadata]
    if missing:
        raise ValueError("The documentation page is missing frontmatter: {}.".format(", ".join(missing)))
    return metadata


def strip_frontmatter(text):
    normalized = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    if not lines or lines[0].strip() != "---":
        return normalized
    for index in range(1, len(lines)):
        if lines[index].strip() == "---":
            return "\n".join(lines[index + 1 :])
    return normalized


def slugify(value):
    text = re.sub(r"[`*_~]", "", str(value or "")).lower().strip()
    text = re.sub(r"[^\w\- ]", "", text, flags=re.UNICODE)
    return re.sub(r"\s+", "-", text)


def _table_cells(line):
    value = line.strip()
    if value.startswith("|"):
        value = value[1:]
    if value.endswith("|"):
        value = value[:-1]
    return [cell.strip() for cell in value.split("|")]


def parse(text):
    """Return a small block AST consumed by the native FlowDocument renderer."""
    body = strip_frontmatter(text)
    lines = body.split("\n")
    blocks = []
    heading_counts = {}
    index = 0
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped:
            index += 1
            continue

        if stripped.startswith("```"):
            language = stripped[3:].strip()
            code = []
            index += 1
            while index < len(lines) and not lines[index].lstrip().startswith("```"):
                code.append(lines[index])
                index += 1
            if index < len(lines):
                index += 1
            blocks.append({"type": "code", "language": language, "text": "\n".join(code)})
            continue

        heading = HEADING_RE.match(line)
        if heading:
            title = heading.group(2)
            base = slugify(title)
            count = heading_counts.get(base, 0)
            heading_counts[base] = count + 1
            anchor = base if count == 0 else "{}-{}".format(base, count)
            blocks.append({"type": "heading", "level": len(heading.group(1)), "text": title, "anchor": anchor})
            index += 1
            continue

        alert = ALERT_RE.match(line)
        if alert:
            alert_type = alert.group(1).upper()
            values = []
            index += 1
            while index < len(lines) and lines[index].lstrip().startswith(">"):
                value = lines[index].lstrip()[1:]
                values.append(value[1:] if value.startswith(" ") else value)
                index += 1
            blocks.append({"type": "alert", "alert_type": alert_type, "text": "\n".join(values).strip()})
            continue

        if line.lstrip().startswith(">"):
            values = []
            while index < len(lines) and lines[index].lstrip().startswith(">"):
                value = lines[index].lstrip()[1:]
                values.append(value[1:] if value.startswith(" ") else value)
                index += 1
            blocks.append({"type": "blockquote", "text": "\n".join(values).strip()})
            continue

        if index + 1 < len(lines) and "|" in line and TABLE_SEPARATOR_RE.match(lines[index + 1]):
            headers = _table_cells(line)
            rows = []
            index += 2
            while index < len(lines) and "|" in lines[index] and lines[index].strip():
                rows.append(_table_cells(lines[index]))
                index += 1
            blocks.append({"type": "table", "headers": headers, "rows": rows})
            continue

        list_match = LIST_RE.match(line)
        if list_match:
            items = []
            while index < len(lines):
                match = LIST_RE.match(lines[index])
                if not match:
                    break
                marker = match.group(2)
                items.append(
                    {
                        "level": max(0, len(match.group(1).replace("\t", "    ")) // 2),
                        "ordered": marker[0].isdigit(),
                        "marker": marker,
                        "text": match.group(3),
                    }
                )
                index += 1
            blocks.append({"type": "list", "items": items})
            continue

        if re.match(r"^\s*(?:---+|___+|\*\*\*+)\s*$", line):
            blocks.append({"type": "rule"})
            index += 1
            continue

        paragraph = [stripped]
        unsupported_html = bool(HTML_RE.match(line))
        index += 1
        while index < len(lines):
            candidate = lines[index]
            candidate_stripped = candidate.strip()
            if not candidate_stripped:
                break
            if (
                candidate_stripped.startswith("```")
                or HEADING_RE.match(candidate)
                or ALERT_RE.match(candidate)
                or candidate.lstrip().startswith(">")
                or LIST_RE.match(candidate)
                or (index + 1 < len(lines) and "|" in candidate and TABLE_SEPARATOR_RE.match(lines[index + 1]))
            ):
                break
            paragraph.append(candidate_stripped)
            unsupported_html = unsupported_html or bool(HTML_RE.match(candidate))
            index += 1
        blocks.append(
            {"type": "paragraph", "text": " ".join(paragraph), "unsupported_html": unsupported_html}
        )
    return blocks


INLINE_RE = re.compile(
    r"(!\[([^\]]*)\]\(([^)]+)\))"
    r"|(\[([^\]]+)\]\(([^)]+)\))"
    r"|(\*\*\*([^*]+)\*\*\*)"
    r"|(\*\*([^*]+)\*\*)"
    r"|(`([^`]+)`)|(?<!\*)\*([^*]+)\*(?!\*)"
)


def parse_inlines(text):
    """Return inline tokens while leaving unsupported constructs as text."""
    result = []
    position = 0
    for match in INLINE_RE.finditer(str(text or "")):
        if match.start() > position:
            result.append({"type": "text", "text": text[position : match.start()]})
        if match.group(1):
            result.append({"type": "image", "alt": match.group(2), "target": match.group(3).strip()})
        elif match.group(4):
            result.append({"type": "link", "text": match.group(5), "target": match.group(6).strip()})
        elif match.group(7):
            result.append({"type": "bold_italic", "text": match.group(8)})
        elif match.group(9):
            result.append({"type": "bold", "text": match.group(10)})
        elif match.group(11):
            result.append({"type": "code", "text": match.group(12)})
        else:
            result.append({"type": "italic", "text": match.group(13)})
        position = match.end()
    if position < len(text):
        result.append({"type": "text", "text": text[position:]})
    return result
