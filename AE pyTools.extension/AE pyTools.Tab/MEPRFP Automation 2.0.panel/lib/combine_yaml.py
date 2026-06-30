# -*- coding: utf-8 -*-
"""
Pure-logic merge of two equipment-definition documents (Combine Yaml).

File A keeps its numbering. File B's profiles are appended to A with all
of their identifiers renumbered so nothing collides with A and B's
internal cross-references stay intact.

Identifier scheme (see WFMELECPROFILES.yaml / HEB_profiles_*.yaml)::

    profile      id: EQ-NNN          ced_truth_source_id: EQ-NNN
    linked set   id: SET-NNN
    LED          id: SET-NNN-LED-MMM
    annotation   id: SET-NNN-LED-MMM-ANN-PPP

References embedded inside LED ``parameters`` (and directives)::

    Set Definition ID:            SET-NNN
    Linked Element Definition ID: SET-NNN-LED-MMM
    sibling_parameter:            SET-NNN-LED-MMM:<param>   (BYSIBLING)

``EQ`` and ``SET`` numbers are *independent* namespaces — a profile
``EQ-392`` can own ``SET-208`` — so they are remapped with separate
counters, each continuing from File A's respective maximum.

This module has no Revit / pyRevit / UI dependency. It operates on the
plain dicts produced by ``yaml_io.parse`` so it can be unit-tested in
plain CPython.
"""

import re


# Matches an ``EQ-<digits>`` or ``SET-<digits>`` token. The lookbehind
# stops it firing inside a larger word (e.g. the "SET-100" in a stray
# "RESET-100"); legitimate ids/refs are always preceded by start-of-
# string, whitespace, or ``:``. Only the prefix + leading number run is
# captured, so "SET-208-LED-001" rewrites the "SET-208" and leaves
# "-LED-001" untouched.
_ID_TOKEN_RE = re.compile(r'(?<![A-Za-z0-9-])(EQ-|SET-)(\d+)')

_EQ_ID_RE = re.compile(r'^EQ-(\d+)$')
_SET_ID_RE = re.compile(r'^SET-(\d+)$')


def _num(value, regex):
    if value is None:
        return None
    m = regex.match(str(value))
    return int(m.group(1)) if m else None


def _linked_sets(profile):
    raw = profile.get("linked_sets")
    if isinstance(raw, list):
        return raw
    if isinstance(raw, dict):
        return [raw]
    return []


def collect_numbers(eqdefs):
    """Return ``(eq_numbers, set_numbers)`` as int sets present in ``eqdefs``."""
    eq_nums = set()
    set_nums = set()
    for profile in eqdefs or []:
        if not isinstance(profile, dict):
            continue
        e = _num(profile.get("id"), _EQ_ID_RE)
        if e is not None:
            eq_nums.add(e)
        for linked_set in _linked_sets(profile):
            if not isinstance(linked_set, dict):
                continue
            s = _num(linked_set.get("id"), _SET_ID_RE)
            if s is not None:
                set_nums.add(s)
    return eq_nums, set_nums


def build_renumber_maps(a_defs, b_defs):
    """Build ``(eq_map, set_map)`` of old-number -> new-number for File B.

    New numbers continue from File A's maxima (EQ and SET independently).
    EQ numbers are assigned in profile order; SET numbers in the order
    sets first appear. Numbers already mapped are not reassigned.
    """
    a_eq, a_set = collect_numbers(a_defs)
    eq_next = max(a_eq) if a_eq else 0
    set_next = max(a_set) if a_set else 0

    eq_map = {}
    set_map = {}
    for profile in b_defs or []:
        if not isinstance(profile, dict):
            continue
        e = _num(profile.get("id"), _EQ_ID_RE)
        if e is not None and e not in eq_map:
            eq_next += 1
            eq_map[e] = eq_next
        for linked_set in _linked_sets(profile):
            if not isinstance(linked_set, dict):
                continue
            s = _num(linked_set.get("id"), _SET_ID_RE)
            if s is not None and s not in set_map:
                set_next += 1
                set_map[s] = set_next
    return eq_map, set_map


def _remap_string(text, eq_map, set_map):
    def repl(match):
        prefix, num = match.group(1), int(match.group(2))
        mapping = eq_map if prefix == "EQ-" else set_map
        new = mapping.get(num)
        if new is None:
            return match.group(0)
        return "{}{:03d}".format(prefix, new)

    return _ID_TOKEN_RE.sub(repl, text)


def _remap_node(node, eq_map, set_map):
    """Rewrite every EQ-/SET- id token in string *values* of ``node``.

    Mutates in place. Dict keys (parameter names) are never touched, so
    ``"Set Definition ID"`` stays a key while its value is remapped.
    """
    if isinstance(node, dict):
        for key in list(node.keys()):
            value = node[key]
            if isinstance(value, str):
                node[key] = _remap_string(value, eq_map, set_map)
            else:
                _remap_node(value, eq_map, set_map)
    elif isinstance(node, list):
        for i in range(len(node)):
            value = node[i]
            if isinstance(value, str):
                node[i] = _remap_string(value, eq_map, set_map)
            else:
                _remap_node(value, eq_map, set_map)


def combine(data_a, data_b):
    """Merge File B into File A and return ``(combined_dict, summary)``.

    File A's content (including any top-level keys beyond
    ``equipment_definitions``) is preserved verbatim; File B's profiles
    are renumbered and appended. ``data_b`` is mutated in place during
    renumbering — pass a freshly parsed copy.

    ``summary`` is a dict::

        {
          "a_profile_count": int,
          "b_profile_count": int,
          "combined_profile_count": int,
          "eq_map": {old: new, ...},
          "set_map": {old: new, ...},
        }
    """
    if not isinstance(data_a, dict):
        raise ValueError("File A did not parse to a YAML mapping.")
    if not isinstance(data_b, dict):
        raise ValueError("File B did not parse to a YAML mapping.")

    a_defs = list(data_a.get("equipment_definitions") or [])
    b_defs = list(data_b.get("equipment_definitions") or [])

    eq_map, set_map = build_renumber_maps(a_defs, b_defs)
    for profile in b_defs:
        _remap_node(profile, eq_map, set_map)

    combined = dict(data_a)
    combined["equipment_definitions"] = a_defs + b_defs

    summary = {
        "a_profile_count": len(a_defs),
        "b_profile_count": len(b_defs),
        "combined_profile_count": len(a_defs) + len(b_defs),
        "eq_map": eq_map,
        "set_map": set_map,
    }
    return combined, summary
