#! python3
# -*- coding: utf-8 -*-
"""MEPRFP Automation 2.0 :: Print Profiles

Export every profile in the active Extensible-Storage YAML to an Excel
(.xlsx) workbook. Layout (one row per child element):

    Profile | parent-filter data | Set | Element | ... | <parameter columns>

The profile name and its parent-filter data appear on the first child row
of each profile; the profile's children (linked-element-definitions) are
listed one per row beneath it, each with its structural data and its
parameters spread across their own columns. A blank row separates profiles.

The workbook is written with the standard library only (zipfile + XML).
pyRevit's bundled CPython 3 engine exposes no ``site-packages``, so
``xlsxwriter`` / ``pyrevit.interop.xl`` are NOT importable here.
"""

import os
import sys
import zipfile

_LIB = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "..", "lib")
)
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import _dev_reload
_dev_reload.purge()

from pyrevit import revit, script

import forms_compat as forms
import active_yaml
import profile_model

TITLE = "Print Profiles (MEPRFP 2.0)"

# Profile-level columns — filled only on each profile's first child row.
PROFILE_HEADERS = [
    "Profile",
    "Parent Category",
    "Family Pattern",
    "Type Pattern",
    "Aliases",
    "Profile ID",
    "Version",
    "Truth Source",
    "Allow Parentless",
    "Allow Unmatched Parents",
    "Prompt On Mismatch",
    "Equipment Properties",
]

# Child (linked-element) columns — filled on every row.
CHILD_HEADERS = [
    "Set",
    "Element (Family : Type)",
    "Category",
    "Group?",
    "Anchor?",
    "X Offset (in)",
    "Y Offset (in)",
    "Z Offset (in)",
    "Rotation (deg)",
    "# Annotations",
]

# Structural columns that precede the dynamic parameter columns.
STRUCT_HEADERS = PROFILE_HEADERS + CHILD_HEADERS

# Index of the "Element" cell — used to count child rows for the report.
_ELEMENT_COL = len(PROFILE_HEADERS) + 1


# ---------------------------------------------------------------------
# Value helpers
# ---------------------------------------------------------------------

def _fmt_param_value(value):
    """Serialize a LED parameter value for a single cell.

    Static values pass through (numbers stay numeric). BYPARENT / BYSIBLING
    directives render as readable tokens.
    """
    if isinstance(value, dict):
        if "parent_parameter" in value:
            return "=PARENT:{}".format(value.get("parent_parameter"))
        if "sibling_parameter" in value:
            return "=SIBLING:{}".format(value.get("sibling_parameter"))
        # Unknown dict shape — show its items rather than dropping data.
        return "; ".join("{}={}".format(k, v) for k, v in value.items())
    return value


def _round(value):
    try:
        return round(float(value), 4)
    except (TypeError, ValueError):
        return value


# ---------------------------------------------------------------------
# Build the table (list of rows; each row is a list of cell values)
# ---------------------------------------------------------------------

def _profile_cells(profile):
    """Profile-level cells (aliases + metadata), in PROFILE_HEADERS order."""
    d = profile.to_dict()
    pf = profile.parent_filter
    aliases = d.get("merged_aliases") or []
    props = profile.equipment_properties or {}
    return [
        profile.name or "<unnamed>",
        pf.category,
        pf.family_name_pattern,
        pf.type_name_pattern,
        " | ".join(str(a) for a in aliases) if aliases else None,
        profile.id,
        d.get("version"),
        profile.truth_source_name,
        bool(profile.allow_parentless),
        bool(profile.allow_unmatched_parents),
        bool(profile.prompt_on_parent_mismatch),
        "; ".join("{}={}".format(k, v) for k, v in props.items())
        if props else None,
    ]


def _child_cells(set_name, led):
    """Child-element cells, in CHILD_HEADERS order."""
    offs = led.offsets
    off0 = offs[0] if offs else None
    return [
        set_name,
        led.label,
        led.category,
        bool(led.is_group),
        bool(led.is_parent_anchor),
        _round(off0.x_inches) if off0 is not None else None,
        _round(off0.y_inches) if off0 is not None else None,
        _round(off0.z_inches) if off0 is not None else None,
        _round(off0.rotation_deg) if off0 is not None else None,
        len(led.annotations),
    ]


def build_table(profiles):
    """Return (headers, rows). ``rows`` cells are str / number / bool / None."""
    # Emit profiles alphabetically by name (case-insensitive; unnamed last).
    profiles = sorted(
        profiles, key=lambda p: ((p.name or "").strip().lower() or "￿")
    )

    # First pass: collect the union of parameter names across every child.
    param_names = set()
    for profile in profiles:
        for lset in profile.linked_sets:
            for led in lset.leds:
                param_names.update((led.parameters or {}).keys())
    param_cols = sorted(param_names)

    headers = list(STRUCT_HEADERS) + param_cols
    blank_profile = [None] * len(PROFILE_HEADERS)
    rows = []

    def _param_cells(led):
        cells = [None] * len(param_cols)
        for name, value in (led.parameters or {}).items():
            cells[param_cols.index(name)] = _fmt_param_value(value)
        return cells

    for p_idx, profile in enumerate(profiles):
        # Separate consecutive profiles with a blank row (skip before the first).
        if p_idx > 0:
            rows.append([None] * len(headers))

        first_child_of_profile = True
        emitted_any = False

        for lset in profile.linked_sets:
            set_name = lset.name or lset.id or ""
            for led in lset.leds:
                # Profile-level data only on the profile's first child row.
                prof = _profile_cells(profile) if first_child_of_profile \
                    else list(blank_profile)
                first_child_of_profile = False
                rows.append(prof + _child_cells(set_name, led) + _param_cells(led))
                emitted_any = True

        # A profile with no children still gets a row so it isn't invisible.
        if not emitted_any:
            rows.append(_profile_cells(profile)
                        + [None] * (len(headers) - len(PROFILE_HEADERS)))

    return headers, rows


# ---------------------------------------------------------------------
# Minimal .xlsx writer (stdlib only)
# ---------------------------------------------------------------------

def _col_letter(idx):
    """0-based column index -> Excel column letter(s)."""
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _xml_escape(text):
    out = []
    for ch in text:
        code = ord(ch)
        # Drop characters not permitted in XML 1.0.
        if not (code in (0x9, 0xA, 0xD)
                or 0x20 <= code <= 0xD7FF
                or 0xE000 <= code <= 0xFFFD):
            continue
        if ch == "&":
            out.append("&amp;")
        elif ch == "<":
            out.append("&lt;")
        elif ch == ">":
            out.append("&gt;")
        elif ch == '"':
            out.append("&quot;")
        else:
            out.append(ch)
    return "".join(out)


def _cell_xml(ref, value):
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return ('<c r="{}" t="inlineStr"><is><t>{}</t></is></c>'
                .format(ref, "TRUE" if value else "FALSE"))
    if isinstance(value, (int, float)):
        return '<c r="{}"><v>{}</v></c>'.format(ref, repr(value))
    return ('<c r="{}" t="inlineStr"><is><t xml:space="preserve">{}</t>'
            '</is></c>'.format(ref, _xml_escape(str(value))))


def _sheet_xml(headers, rows):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/'
        'spreadsheetml/2006/main"><sheetData>',
    ]
    all_rows = [headers] + rows
    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, value in enumerate(row):
            cell = _cell_xml("{}{}".format(_col_letter(c_idx), r_idx), value)
            if cell:
                cells.append(cell)
        parts.append('<row r="{}">{}</row>'.format(r_idx, "".join(cells)))
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
    'content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
    'package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
    'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '</Types>'
)

_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
    'relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)

_WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/'
    'main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships"><sheets>'
    '<sheet name="Profiles" sheetId="1" r:id="rId1"/>'
    '</sheets></workbook>'
)

_WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
    'relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/worksheet" '
    'Target="worksheets/sheet1.xml"/>'
    '</Relationships>'
)


def write_xlsx(path, headers, rows):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("xl/workbook.xml", _WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", _WORKBOOK_RELS)
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml(headers, rows))


# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------

def main():
    output = script.get_output()
    output.close_others()

    doc = revit.doc
    if doc is None:
        forms.alert("No active document.", title=TITLE)
        return

    data = active_yaml.load_active_data(doc)
    eqdefs = (data or {}).get("equipment_definitions") or []
    if not eqdefs:
        forms.alert("No profiles in the active store.", title=TITLE)
        return

    profiles = profile_model.ProfileDocument(data).profiles

    headers, rows = build_table(profiles)

    save_path = forms.save_file(
        file_ext="xlsx",
        title="Save profiles workbook",
        default_name="profiles",
    )
    if not save_path:
        return

    try:
        write_xlsx(save_path, headers, rows)
    except Exception as ex:  # pragma: no cover - surfaced to the user
        forms.alert(
            "Could not write the workbook:\n{}".format(ex), title=TITLE
        )
        raise

    n_children = sum(1 for r in rows if r[_ELEMENT_COL] not in (None, ""))
    output.print_md(
        "**Print Profiles complete** — wrote {} profile(s), {} child "
        "element(s), {} parameter column(s) to:\n\n`{}`".format(
            len(profiles),
            n_children,
            len(headers) - len(STRUCT_HEADERS),
            save_path,
        )
    )


if __name__ == "__main__":
    main()
