# -*- coding: utf-8 -*-
"""
Stage 6 — Spaces annotation placement.

Walks every placed family instance in the active view whose
Element_Linker carries the ``space_id`` lineage marker, looks up the
source LED in ``space_profiles[*]``, and emits an
``AnnotationCandidate`` for each ``annotations[*]`` entry. Reuses the
equipment-side ``annotation_placement._place_tag`` /
``_place_keynote`` / ``_place_text_note`` for the Revit-API edge so
tag-vs-text-note routing matches what the equipment pipeline does.
"""

import math

import clr  # noqa: F401

from Autodesk.Revit.DB import (  # noqa: E402
    FamilyInstance,
    FilteredElementCollector,
    LocationPoint,
    XYZ,
)

import annotation_placement as _ap  # noqa: E402
import element_linker_io as _el_io  # noqa: E402
import geometry  # noqa: E402


# ---------------------------------------------------------------------
# LED index
# ---------------------------------------------------------------------

def _build_space_led_index(profile_data):
    """Return ``(by_id, by_label)`` indexes over ``space_profiles``.

    ``by_id`` is the strong link: ``{led_id: (profile, set, led)}`` —
    used first because it matches the exact LED that was stamped onto
    the placed fixture's Element_Linker.

    ``by_label`` is the fallback used when the fixture's stamped
    ``led_id`` no longer exists (typical after a duplicate-profile,
    delete-and-recreate, or import that re-issued LED IDs). It maps
    the LED's label (``"Family : Type"``) to a list of matching LEDs;
    the workflow only adopts the fallback when the list has exactly
    one entry, so we never silently bind to the wrong LED when two
    profiles share a family/type.
    """
    by_id = {}
    by_label = {}
    for profile in profile_data.get("space_profiles") or []:
        if not isinstance(profile, dict):
            continue
        for set_dict in profile.get("linked_sets") or []:
            if not isinstance(set_dict, dict):
                continue
            for led in set_dict.get("linked_element_definitions") or []:
                if not isinstance(led, dict):
                    continue
                lid = led.get("id")
                if lid:
                    by_id[lid] = (profile, set_dict, led)
                label = (led.get("label") or "").strip()
                if label:
                    by_label.setdefault(label, []).append(
                        (profile, set_dict, led),
                    )
    return by_id, by_label


def _fixture_label(fixture):
    """``"Family : Type"`` for a FamilyInstance, or ``""`` if unreadable."""
    sym = getattr(fixture, "Symbol", None)
    if sym is None:
        return ""
    fam = getattr(sym, "FamilyName", "") or ""
    typ = getattr(sym, "Name", "") or ""
    if fam and typ:
        return "{} : {}".format(fam, typ)
    return fam or typ or ""


def _fixture_pt_and_rot(fixture):
    loc = getattr(fixture, "Location", None)
    pt = None
    rad = 0.0
    if isinstance(loc, LocationPoint):
        try:
            pt = loc.Point
        except Exception:
            pt = None
        try:
            rad = loc.Rotation
        except Exception:
            rad = 0.0
    if pt is None:
        try:
            bbox = fixture.get_BoundingBox(None)
        except Exception:
            bbox = None
        if bbox is not None:
            pt = XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
    if pt is None:
        return None, 0.0
    return (pt.X, pt.Y, pt.Z), geometry.normalize_angle(math.degrees(rad))


# ---------------------------------------------------------------------
# Collection
# ---------------------------------------------------------------------

def collect_space_candidates(doc, view, profile_data, kinds=None,
                             skip_duplicates=True):
    """Return ``(candidates, diagnostics)`` for space-based fixtures.

    Filters fixtures to those where ``Element_Linker.is_space_based``
    is true (i.e. ``space_id`` was stamped at placement time).
    ``kinds`` defaults to all three (tag / keynote / text note).

    When ``skip_duplicates`` is True (the default), each candidate is
    run through ``annotation_placement.mark_duplicates`` so already-
    placed tags / keynotes / text notes get ``skip = True`` and a
    ``duplicate_reason`` populated. The UI shows the reason and
    ``execute_placement`` skips them at apply time.

    ``diagnostics`` is a counter dict the UI uses to explain WHY the
    candidate list is empty when it is — each filter stage records how
    many fixtures it dropped, so the user can see whether the
    bottleneck is "no Element_Linker", "Element_Linker present but no
    space_id", "led_id doesn't match anything in space_profiles", or
    "matched a LED with no annotations".
    """
    if kinds is None:
        kinds = set(_ap.ALL_KINDS)
    else:
        kinds = set(kinds)

    led_by_id, led_by_label = _build_space_led_index(profile_data)
    out = []
    diag = {
        "fixtures_scanned": 0,
        "with_linker": 0,
        "space_based": 0,
        "with_led_id": 0,
        "led_in_index": 0,
        "led_has_annotations": 0,
        "passed_kind_filter": 0,
        "leds_in_space_profiles": len(led_by_id),
        # Counts fixtures whose stamped led_id was orphaned but whose
        # family:type label still uniquely matched a current space LED.
        # Surface this in the status bar so the user knows their data is
        # in a "rebound by label" state and may want to re-place to
        # refresh the Element_Linker.
        "rebound_by_label": 0,
        # Fixtures whose label matched more than one current LED — we
        # refuse to guess in that case.
        "label_ambiguous": 0,
    }

    if view is not None:
        host_collector = FilteredElementCollector(doc, view.Id)
    else:
        host_collector = FilteredElementCollector(doc)

    for fixture in host_collector.OfClass(FamilyInstance).WhereElementIsNotElementType():
        diag["fixtures_scanned"] += 1
        linker = _el_io.read_from_element(fixture)
        if linker is None:
            continue
        diag["with_linker"] += 1
        if not linker.is_space_based:
            continue
        diag["space_based"] += 1
        if not linker.led_id:
            continue
        diag["with_led_id"] += 1
        entry = led_by_id.get(linker.led_id)
        if entry is None:
            # Stamped led_id no longer exists in space_profiles. Try
            # rebinding by the fixture's family:type label — if exactly
            # one current LED has that label, treat it as the same LED.
            label = _fixture_label(fixture)
            candidates_by_label = led_by_label.get(label) or []
            if len(candidates_by_label) == 1:
                entry = candidates_by_label[0]
                diag["rebound_by_label"] += 1
            elif len(candidates_by_label) > 1:
                diag["label_ambiguous"] += 1
                continue
            else:
                continue
        diag["led_in_index"] += 1
        profile, _set, led = entry
        annotations = led.get("annotations") or []
        if not annotations:
            continue
        diag["led_has_annotations"] += 1
        fixture_pt, fixture_rot = _fixture_pt_and_rot(fixture)
        if fixture_pt is None:
            continue

        for ann in annotations:
            if not isinstance(ann, dict):
                continue
            kind = ann.get("kind") or _ap.KIND_TAG
            if kind not in kinds:
                continue
            diag["passed_kind_filter"] += 1
            offset = ann.get("offsets") or {}
            if isinstance(offset, list):
                offset = offset[0] if offset else {}
            target_pt = geometry.target_point_from_offsets(
                fixture_pt, fixture_rot, offset,
            )
            target_rot = geometry.child_rotation_from_offsets(
                fixture_rot, offset,
            )
            out.append(_ap.AnnotationCandidate(
                fixture=fixture,
                fixture_pt=fixture_pt,
                fixture_rot=fixture_rot,
                led_id=led.get("id") or "",
                led_label=led.get("label") or "",
                profile_id=profile.get("id") or "",
                profile_name=profile.get("name") or "",
                annotation=ann,
                target_pt=target_pt,
                target_rot=target_rot,
            ))

    if skip_duplicates and out:
        try:
            _ap.mark_duplicates(doc, view, out)
        except Exception:
            # Dedup is a UX nicety — never let it block placement preview.
            pass

    return out, diag


def explain_empty_candidates(diag):
    """Return a human-readable string explaining why no candidates were
    produced. Reads the diagnostic dict from
    ``collect_space_candidates`` and points at the first stage that
    dropped everything. Returns ``""`` when the dict's
    ``passed_kind_filter`` is non-zero (i.e. there's nothing to
    explain).
    """
    if diag.get("passed_kind_filter", 0) > 0:
        return ""
    n_total = diag.get("fixtures_scanned", 0)
    if n_total == 0:
        return "No FamilyInstances visible in this view."
    if diag.get("with_linker", 0) == 0:
        return (
            "{} fixture(s) in view, but none carry an Element_Linker. "
            "These weren't placed by the MEPRFP 2.0 pipeline.".format(n_total)
        )
    if diag.get("space_based", 0) == 0:
        return (
            "{} fixture(s) carry an Element_Linker, but none were placed "
            "by the Spaces pipeline (no space_id stamped). Use 'Place "
            "Element Annotations' for equipment-pipeline fixtures.".format(
                diag.get("with_linker", 0),
            )
        )
    n_leds = diag.get("leds_in_space_profiles", 0)
    if diag.get("led_in_index", 0) == 0:
        ambig = diag.get("label_ambiguous", 0)
        ambig_note = (
            " ({} fixture(s) had a label that matched more than one LED "
            "and were skipped to avoid binding to the wrong one — make "
            "the LED labels unique or re-place those fixtures.)".format(
                ambig,
            )
            if ambig else ""
        )
        return (
            "{} space-based fixture(s) found, but none of their LED IDs "
            "(or fallback family:type labels) match any of the {} LED(s) "
            "in space_profiles. The space LED may have been renamed or "
            "deleted since the fixture was placed.{}".format(
                diag.get("space_based", 0), n_leds, ambig_note,
            )
        )
    if diag.get("led_has_annotations", 0) == 0:
        return (
            "{} space-based fixture(s) matched a space LED, but none of "
            "those LEDs have any saved annotations. Add an annotation in "
            "Manage Space Profiles -> Params/Offsets/Annotations.".format(
                diag.get("led_in_index", 0),
            )
        )
    return (
        "{} candidate(s) had annotations but none matched the selected "
        "annotation kind(s). Tick a different kind checkbox.".format(
            diag.get("led_has_annotations", 0),
        )
    )


# ---------------------------------------------------------------------
# Apply (delegates to the equipment-side machinery)
# ---------------------------------------------------------------------

def execute_placement(doc, view, candidates):
    return _ap.execute_placement(doc, view, candidates)


def is_view_eligible(view):
    return _ap.is_view_eligible(view)
