# -*- coding: utf-8 -*-
"""Dependency-free .xlsx reader for the PFAI import.

The panel already writes workbooks with `zipfile` alone (see Print Profiles), so
reading one the same way keeps the button free of openpyxl - which is not
installed with the pyRevit CPython engine on every machine.

Only what a PFAI workbook needs is supported: shared strings, inline strings,
numbers, and booleans, on the sheets named in the workbook part. Formulas are
read as their cached value; a workbook written by `make_pfai_workbook` has none.

    book = read_workbook(path)          # {sheet_name: [[cell, ...], ...]}
    rows = sheet_dicts(book, "Devices") # [{header: value}, ...]
"""

import re
import zipfile
import xml.etree.ElementTree as ET

__all__ = ["read_workbook", "sheet_dicts", "XlsxError"]

_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


class XlsxError(Exception):
    pass


def _col_index(ref):
    """'AB12' -> 27 (0-based column)."""
    m = _CELL_RE.match(ref or "")
    if not m:
        return None
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _shared_strings(zf):
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall("m:si", _NS):
        # a shared string may be split across several runs
        out.append("".join(t.text or "" for t in si.iter(
            "{%s}t" % _NS["m"])))
    return out


def _sheet_parts(zf):
    """[(sheet name, zip path)] in workbook order."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = {}
    for rel in rels.findall("p:Relationship", _NS):
        t = rel.get("Target") or ""
        if t.startswith("/"):
            t = t[1:]
        elif not t.startswith("xl/"):
            t = "xl/" + t
        target[rel.get("Id")] = t
    out = []
    for sh in wb.findall("m:sheets/m:sheet", _NS):
        rid = sh.get("{%s}id" % _NS["r"])
        if rid in target:
            out.append((sh.get("name"), target[rid]))
    return out


def _cell_value(c, strings):
    t = c.get("t")
    if t == "inlineStr":
        return "".join(x.text or "" for x in c.iter("{%s}t" % _NS["m"]))
    v = c.find("m:v", _NS)
    if v is None or v.text is None:
        return ""
    raw = v.text
    if t == "s":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return ""
    if t == "b":
        return raw not in ("0", "", "false", "FALSE")
    if t in ("str", "e"):
        return raw
    try:
        f = float(raw)
    except ValueError:
        return raw
    return int(f) if f == int(f) and abs(f) < 1e15 else f


def read_workbook(path):
    """{sheet name: list of rows, each a list of cell values}."""
    try:
        zf = zipfile.ZipFile(path, "r")
    except (IOError, OSError, zipfile.BadZipfile) as exc:
        raise XlsxError("Could not open %s: %s" % (path, exc))
    try:
        strings = _shared_strings(zf)
        book = {}
        for name, part in _sheet_parts(zf):
            try:
                root = ET.fromstring(zf.read(part))
            except KeyError:
                continue
            rows = []
            for row in root.findall("m:sheetData/m:row", _NS):
                cells = []
                for c in row.findall("m:c", _NS):
                    i = _col_index(c.get("r"))
                    if i is None:
                        i = len(cells)
                    # honour the r= reference so blank cells keep their column
                    while len(cells) < i:
                        cells.append("")
                    cells.append(_cell_value(c, strings))
                rows.append(cells)
            book[name] = rows
        return book
    finally:
        zf.close()


def sheet_dicts(book, sheet_name):
    """Rows of `sheet_name` as dicts keyed by the header row.

    Blank rows are dropped. Headers are lower-cased and stripped so the caller
    is not at the mercy of how the workbook was capitalised.
    """
    rows = book.get(sheet_name)
    if rows is None:
        raise XlsxError(
            "The workbook has no '%s' sheet (found: %s)."
            % (sheet_name, ", ".join(sorted(book)) or "none"))
    if not rows:
        return []
    head = [str(h).strip().lower() for h in rows[0]]
    out = []
    for r in rows[1:]:
        if not any(str(x).strip() for x in r):
            continue
        d = {}
        for i, key in enumerate(head):
            if key:
                d[key] = r[i] if i < len(r) else ""
        out.append(d)
    return out
