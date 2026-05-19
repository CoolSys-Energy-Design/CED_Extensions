# -*- coding: utf-8 -*-
"""
Synced-relationship audit.

For every placed child element with an ``Element_Linker`` payload, we
walk the LED's ``parameters`` dict; any ``parent_parameter`` or
``sibling_parameter`` directive defines an *expected* value (read off
the parent / sibling). If the child's actual value diverges, we record
a Conflict carrying the **percentage difference** between the source
and the child. The UI groups conflicts as ``Profile -> LED ->
Conflict``.

This audit is **flag-only**: ``detect_conflicts`` reports divergence;
nothing here ever mutates the model. There is intentionally no
resolution / write path.

Zero-source rule: a ``parent`` directive whose source parameter reads
0 while the child reads non-0 is NOT flagged — a 0 on the parent means
the directive source is unconfigured / N/A for that instance.
"""

import re

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    FamilyInstance,
    FilteredElementCollector,
    Group,
    StorageType,
)

import directives as _dir
import element_linker_io as _el_io
import profile_model


_FIRST_NUMERIC_RE = re.compile(r"\d+")


def _first_numeric_token(s):
    """First run of digits in ``s``, e.g. ``"SET-323-LED-002" -> "323"``,
    ``"EQ-464" -> "464"``. Returns ``None`` when nothing matches."""
    if s is None:
        return None
    m = _FIRST_NUMERIC_RE.search(str(s))
    return m.group(0) if m else None


# ---------------------------------------------------------------------
# Conflict / Decision data classes
# ---------------------------------------------------------------------

class Conflict(object):
    """One detected mismatch.

    This audit is **flag-only** — there is no correction path. A
    Conflict is a report row: the directive source value (``expected``,
    read off the parent/sibling) vs the child's ``actual``, plus the
    ``percent_difference`` between them when both are numeric.
    """

    def __init__(self, profile_id, profile_name, led_id, led_label,
                 element_id, parameter_name, kind, expected_value,
                 actual_value, target_param_name, target_element_id,
                 percent_difference=None):
        self.profile_id = profile_id
        self.profile_name = profile_name
        self.led_id = led_id
        self.led_label = led_label
        self.element_id = element_id
        self.parameter_name = parameter_name
        self.kind = kind  # "parent" | "sibling"
        self.expected_value = expected_value
        self.actual_value = actual_value
        self.target_param_name = target_param_name
        self.target_element_id = target_element_id
        # float percent (e.g. 20.0 for 50 -> 60), or None when one of
        # the values isn't numeric / the source is 0.
        self.percent_difference = percent_difference

    @property
    def key(self):
        return (self.profile_id, self.led_id, self.element_id, self.parameter_name)

    @property
    def percent_display(self):
        if self.percent_difference is None:
            return "n/a (non-numeric)"
        return "{:.1f}%".format(self.percent_difference)

    def to_display_dict(self):
        return {
            "profile": "{} ({})".format(self.profile_name, self.profile_id),
            "led": "{} ({})".format(self.led_label, self.led_id),
            "element_id": self.element_id,
            "parameter": self.parameter_name,
            "kind": self.kind,
            "expected": self.expected_value,
            "actual": self.actual_value,
            "percent_difference": self.percent_difference,
        }


# ---------------------------------------------------------------------
# Parameter read / write
# ---------------------------------------------------------------------

def _read_param_value(elem, name):
    if elem is None:
        return None
    param = elem.LookupParameter(name)
    if param is None:
        return None
    storage = param.StorageType
    if storage == StorageType.String:
        return param.AsString()
    if storage == StorageType.Integer:
        return param.AsInteger()
    if storage == StorageType.Double:
        return param.AsDouble()
    if storage == StorageType.ElementId:
        eid = param.AsElementId()
        return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)
    return param.AsValueString() or param.AsString()


_NUMBER_RE = re.compile(r"[-+]?\d*\.?\d+")


def _to_number(value):
    """Best-effort numeric coercion. ``50`` / ``50.0`` / ``"50 A"`` /
    ``"-1' - 6\""``-ish all yield the leading signed number; anything
    with no parseable digits yields ``None``.

    Used so the audit can compute a percentage difference even when a
    parameter reads back as a unit-bearing string (``"50 A"``).
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    m = _NUMBER_RE.search(str(value).strip())
    if not m:
        return None
    try:
        return float(m.group(0))
    except (TypeError, ValueError):
        return None


def _percent_difference(expected, actual):
    """Percent the child (``actual``) diverges from the directive source
    (``expected``): ``abs(actual - expected) / abs(expected) * 100``.

    e.g. expected 50, actual 60 -> 20.0. Returns ``None`` when either
    side isn't numeric or ``expected`` is 0 (undefined / handled by the
    caller's zero-source rule).
    """
    e = _to_number(expected)
    a = _to_number(actual)
    if e is None or a is None or e == 0:
        return None
    return abs(a - e) / abs(e) * 100.0


# ---------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------

def _values_match(a, b):
    if a is None and b is None:
        return True
    if isinstance(a, float) or isinstance(b, float):
        try:
            return abs(float(a) - float(b)) < 1e-6
        except (TypeError, ValueError):
            pass
    return str(a or "").strip().lower() == str(b or "").strip().lower()


def _collect_placed_elements_with_linker(doc):
    """Return ``{(profile_id_or_None, set_id, led_id): [(elem, linker), ...]}``."""
    out = {}
    for klass in (FamilyInstance, Group):
        collector = FilteredElementCollector(doc).OfClass(klass).WhereElementIsNotElementType()
        for elem in collector:
            linker = _el_io.read_from_element(elem)
            if linker is None or not linker.set_id or not linker.led_id:
                continue
            key = (linker.set_id, linker.led_id)
            out.setdefault(key, []).append((elem, linker))
    return out


def _build_sibling_lookup(siblings_in_set):
    """Given ``[(elem, linker), ...]`` for one set, build a callable
    ``(led_id, param_name) -> value | None``."""
    by_led = {}
    for elem, linker in siblings_in_set:
        by_led.setdefault(linker.led_id, []).append(elem)

    def lookup(led_id, param_name):
        for elem in by_led.get(led_id, []):
            v = _read_param_value(elem, param_name)
            if v is not None:
                return v
        return None
    return lookup


def detect_conflicts(doc, profile_data):
    """Return a list of ``Conflict``."""
    conflicts = []
    pdoc = profile_model.ProfileDocument(profile_data)

    # Group all placed elements by set so we can resolve sibling references.
    placed = _collect_placed_elements_with_linker(doc)
    siblings_in_set = {}
    for (set_id, led_id), entries in placed.items():
        siblings_in_set.setdefault(set_id, []).extend(entries)

    # ``(set_id, led_id) -> owning profile id``. A placed element's
    # Element_Linker carries set_id + led_id but not profile_id, so we
    # have to infer the owner by walking the YAML. In normal data the
    # set_id is unique to a profile and there's exactly one match. In
    # data with overlapping ids (duplicated / merged profiles that
    # didn't get their LED ids re-stamped), multiple profiles can claim
    # the same (set_id, led_id).
    #
    # Tiebreaker: prefer the profile whose ``id`` numeric token matches
    # the set's numeric token (e.g. ``EQ-323`` wins over ``EQ-464`` for
    # ``SET-323-LED-002``). Capture-time naming uses matching numeric
    # tokens between profile and its captured sets, so this picks the
    # original owner rather than a later duplicate. When no profile
    # matches numerically (or set/profile ids don't carry numeric
    # tokens), first-in-YAML wins, which is stable and matches the
    # original capture order.
    claimants_by_set_led = {}
    for profile in pdoc.profiles:
        for linked_set in profile.linked_sets:
            for led in linked_set.leds:
                key = (linked_set.id, led.id)
                claimants_by_set_led.setdefault(key, []).append(profile.id)
    owner_by_set_led = {}
    for (set_id, led_id), claimants in claimants_by_set_led.items():
        set_num = _first_numeric_token(set_id)
        chosen = None
        if set_num is not None:
            for pid in claimants:
                if _first_numeric_token(pid) == set_num:
                    chosen = pid
                    break
        owner_by_set_led[(set_id, led_id)] = chosen if chosen is not None else claimants[0]

    for profile in pdoc.profiles:
        for linked_set in profile.linked_sets:
            sibling_lookup = _build_sibling_lookup(
                siblings_in_set.get(linked_set.id, [])
            )
            for led in linked_set.leds:
                # Skip non-owning profiles for shared (set_id, led_id)
                # pairs — the placed element only belongs to one
                # profile in the user's mental model, so emit the
                # conflict against that one and not its duplicates.
                if owner_by_set_led.get((linked_set.id, led.id)) != profile.id:
                    continue
                params = led.parameters or {}
                placed_entries = placed.get((linked_set.id, led.id), [])
                if not placed_entries:
                    continue
                # Each placed instance gets its own audit pass — the
                # parent for resolution is whichever element the linker
                # points at as the parent.
                for elem, linker in placed_entries:
                    parent_elem = None
                    if linker.parent_element_id:
                        parent_elem = doc.GetElement(_make_element_id(doc, linker.parent_element_id))
                    parent_lookup = _build_parent_lookup(parent_elem)
                    for param_name, value in params.items():
                        kind = _dir.directive_kind(value)
                        if kind == "static":
                            continue
                        found, expected = _dir.resolve_expected_value(
                            value, parent_lookup, sibling_lookup
                        )
                        if not found:
                            continue
                        actual = _read_param_value(elem, param_name)
                        if _values_match(actual, expected):
                            continue

                        exp_num = _to_number(expected)
                        act_num = _to_number(actual)

                        # Zero-source rule: when the monitored PARENT
                        # parameter is 0 but the child's value is not 0,
                        # do NOT flag. A 0 on the parent means the
                        # directive source is unconfigured / N/A for
                        # this instance, so a non-zero child isn't a
                        # real divergence worth reporting. (Only applies
                        # to parent directives, per spec; the inverse —
                        # parent set, child 0 — IS still flagged.)
                        if (
                            kind == "parent"
                            and exp_num is not None
                            and exp_num == 0
                            and act_num is not None
                            and act_num != 0
                        ):
                            continue

                        conflicts.append(Conflict(
                            profile_id=profile.id,
                            profile_name=profile.name,
                            led_id=led.id,
                            led_label=led.label,
                            element_id=_id_value(elem),
                            parameter_name=param_name,
                            kind=kind,
                            expected_value=expected,
                            actual_value=actual,
                            target_param_name=_target_param_name(value, kind),
                            target_element_id=(
                                _id_value(parent_elem) if kind == "parent"
                                else _sibling_target_id(value, siblings_in_set.get(linked_set.id, []))
                            ),
                            percent_difference=_percent_difference(
                                expected, actual
                            ),
                        ))
    return conflicts


def _make_element_id(doc, value):
    from Autodesk.Revit.DB import ElementId
    try:
        return ElementId(int(value))
    except Exception:
        return ElementId.InvalidElementId


def _id_value(elem):
    if elem is None:
        return None
    eid = elem.Id
    return getattr(eid, "Value", None) or getattr(eid, "IntegerValue", None)


def _build_parent_lookup(parent_elem):
    def lookup(param_name):
        return _read_param_value(parent_elem, param_name)
    return lookup


def _target_param_name(directive_value, kind):
    if kind == "parent":
        return _dir.parent_param_name(directive_value)
    if kind == "sibling":
        target = _dir.sibling_target(directive_value)
        return target[1] if target else None
    return None


def _sibling_target_id(directive_value, set_entries):
    target = _dir.sibling_target(directive_value)
    if target is None:
        return None
    led_id, _ = target
    for elem, linker in set_entries:
        if linker.led_id == led_id:
            return _id_value(elem)
    return None


# ---------------------------------------------------------------------
# Resolution
# ---------------------------------------------------------------------
#
# Intentionally none. This audit is flag-only: it reports divergence on
# a percentage basis and never mutates the model. The old
# apply_resolution / _write_param_value pair (Update child / Update
# parent) was removed — "flagging is our only mission".
