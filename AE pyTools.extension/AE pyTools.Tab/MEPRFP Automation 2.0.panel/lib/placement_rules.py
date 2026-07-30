"""Pure placement-decision rules (no Revit API imports).

``placement.py`` imports the Revit API at module top, which makes its
helpers untestable in plain CPython. The label-shape decisions —
"is this LED a model group or a family, and which names should the
group lookup try" — are pure string logic, so they live here where
``_run_tests.py`` can exercise them directly.

Kinds returned by :func:`resolve_placement_kind`:

``"group"``
    The LED is explicitly flagged ``is_group`` by capture (trusted).
    Try every entry in ``group_candidates`` against the GroupType
    index; on total miss the caller must surface ``group_missing`` —
    never fall through to the family path, the flag is authoritative.

``"maybe_group"``
    Legacy heuristic: the label split as ``"X : X"`` (groups have no
    type axis, so V5-era captures serialized them family==type with
    ``is_group: false``). Try the group lookup first, but fall back
    to family resolution on miss — the label might genuinely be a
    family whose type shares its name.

``"family"``
    Normal family LED; ``group_candidates`` is empty.
"""


def split_label(label):
    """``"Family : Type"`` -> ``("Family", "Type")``. Single-word labels
    fall back to ``(label, "")``."""
    if not label:
        return "", ""
    if " : " in label:
        family, type_name = label.split(" : ", 1)
        return family.strip(), type_name.strip()
    return label.strip(), ""


def resolve_placement_kind(label, is_group):
    """Classify a LED label for placement.

    Returns ``(kind, group_candidates, family, type_name)``.

    For explicit groups the raw label is tried FIRST: a group whose
    own name contains ``" : "`` would be destroyed by the family/type
    split, so the un-split string must win when it names a real
    GroupType. The family half is kept as a secondary candidate for
    the common ``"X : X"`` serialization of older captures.
    """
    raw = (label or "").strip()
    family, type_name = split_label(label)

    if is_group:
        candidates = []
        if raw:
            candidates.append(raw)
        if family and family != raw:
            candidates.append(family)
        return "group", candidates, family, type_name

    if family and type_name and family == type_name:
        return "maybe_group", [family], family, type_name

    return "family", [], family, type_name
