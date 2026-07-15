# -*- coding: utf-8 -*-
"""Reusable Revit category helpers."""

import re

from pyrevit import DB

from Snippets import revit_helpers

_INVALID_CATEGORY_ID_VALUE = revit_helpers.get_elementid_value(DB.ElementId.InvalidElementId, default=-1)


def normalize_category_name(text):
    """Normalize a display/API category name for tolerant equality checks."""
    return re.sub(r"[^a-z0-9]+", "", str(text or "").strip().lower())


def category_id_value(category_or_id, default=_INVALID_CATEGORY_ID_VALUE):
    """Return a numeric category id from Category/ElementId-like inputs."""
    if category_or_id is None:
        return int(default)
    try:
        maybe_id = getattr(category_or_id, "Id", None)
    except Exception:
        maybe_id = None
    if isinstance(maybe_id, DB.ElementId):
        return revit_helpers.get_elementid_value(maybe_id, default=default)
    return revit_helpers.get_elementid_value(category_or_id, default=default)


def is_valid_category_id_value(value):
    """Return True when value is a real category id."""
    try:
        numeric = int(value)
    except Exception:
        return False
    return int(numeric) != int(_INVALID_CATEGORY_ID_VALUE)


def built_in_category_from_name(name):
    """Return DB.BuiltInCategory by API name such as OST_ElectricalFixtures."""
    try:
        return getattr(DB.BuiltInCategory, str(name or ""), None)
    except Exception:
        return None


def category_id_from_bic(bic):
    """Return an ElementId for a BuiltInCategory enum value."""
    if bic is None:
        return None
    try:
        return revit_helpers.elementid_from_value(int(bic))
    except Exception:
        return None


def _category_exists_in_doc(doc, category_id):
    if doc is None or category_id is None:
        return True
    target_value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
    if not is_valid_category_id_value(target_value):
        return False
    try:
        for cat in list(doc.Settings.Categories or []):
            if cat is None:
                continue
            if category_id_value(cat, default=_INVALID_CATEGORY_ID_VALUE) == int(target_value):
                return True
    except Exception:
        pass
    return False


def unique_category_ids(category_ids):
    """Return category ids de-duplicated by numeric ElementId value."""
    unique = []
    seen = set()
    for category_id in list(category_ids or []):
        value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if (not is_valid_category_id_value(value)) or value in seen:
            continue
        seen.add(value)
        unique.append(category_id)
    return unique


def category_ids_from_bics(doc, bics):
    """Return valid category ElementIds for BuiltInCategory values."""
    ids = []
    for bic in list(bics or []):
        category_id = category_id_from_bic(bic)
        if category_id is None:
            continue
        if not _category_exists_in_doc(doc, category_id):
            continue
        ids.append(category_id)
    return unique_category_ids(ids)


def category_map_by_display_name(doc):
    """Return project categories keyed by normalized display name."""
    category_map = {}
    try:
        for cat in list(doc.Settings.Categories or []):
            if cat is None:
                continue
            category_map[normalize_category_name(cat.Name)] = cat
    except Exception:
        category_map = {}
    return category_map


def category_by_display_name(doc, name):
    """Return a Category from the project by display name."""
    key = normalize_category_name(name)
    if not key:
        return None
    return category_map_by_display_name(doc).get(key)


def category_ids_from_display_names(doc, names):
    """Resolve display category names to ElementIds; return (ids, missing_names)."""
    resolved = []
    missing = []
    category_map = category_map_by_display_name(doc)

    for name in list(names or []):
        key = normalize_category_name(name)
        if not key:
            continue
        cat = category_map.get(key)
        if cat is None:
            missing.append(str(name))
            continue
        resolved.append(getattr(cat, "Id", None))

    return unique_category_ids(resolved), missing


def split_category_names(categories_value):
    """Split a comma-separated display category list."""
    return [x.strip() for x in str(categories_value or "").split(",") if x and str(x).strip()]


def category_id_values(category_ids):
    values = set()
    for category_id in list(category_ids or []):
        value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if not is_valid_category_id_value(value):
            continue
        values.add(int(value))
    return values


def category_id_values_from_categories(categories):
    """Return numeric id values for Category collections."""
    return category_id_values([getattr(cat, "Id", None) for cat in list(categories or []) if cat is not None])


def build_category_set(doc, category_ids):
    """Build a DB.CategorySet from Category/ElementId inputs."""
    category_set = DB.CategorySet()
    missing = []
    inserted = 0

    category_map = {}
    try:
        for cat in list(doc.Settings.Categories or []):
            if cat is None:
                continue
            cat_id_value = category_id_value(cat, default=_INVALID_CATEGORY_ID_VALUE)
            if is_valid_category_id_value(cat_id_value):
                category_map[cat_id_value] = cat
    except Exception:
        category_map = {}

    for category_id in list(unique_category_ids(category_ids) or []):
        cat_id_value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if not is_valid_category_id_value(cat_id_value):
            continue
        category = category_map.get(cat_id_value)
        if category is None:
            missing.append(str(cat_id_value))
            continue
        category_set.Insert(category)
        inserted += 1

    return category_set, inserted, missing


def build_category_set_from_display_names(doc, names):
    """Build a DB.CategorySet from project category display names."""
    category_ids, missing_names = category_ids_from_display_names(doc, names)
    category_set, inserted, missing_ids = build_category_set(doc, category_ids)
    missing = list(missing_names or []) + list(missing_ids or [])
    return category_set, inserted, missing


def merge_category_sets(doc, first_categories, second_categories):
    """Return CategorySet union of two category collections."""
    merged_values = set()
    merged_values.update(category_id_values(first_categories))
    merged_values.update(category_id_values(second_categories))
    merged_ids = [revit_helpers.elementid_from_value(int(v)) for v in sorted(list(merged_values or []))]
    return build_category_set(doc, merged_ids)


# Compatibility helpers for electrical parameter-table binding scopes.
_ELECTRICAL_CIRCUIT_TOKENS = set(
    ["electricalcircuits", "electricalcircuit", "eelctricalcircuits", "eelctricalcircuit"]
)
_ALL_ELECTRICAL_TOKENS = set(["allelectrical", "allelectricalcategories"])
_ELECTRICAL_EQUIPMENT_TOKENS = set(["electricalequipment"])
_ELECTRICAL_FIXTURE_TOKENS = set(["electricalfixtures", "electricaldevices"])

_FIXTURE_DEVICE_BIC_NAMES = (
    "OST_ElectricalFixtures",
    "OST_LightingFixtures",
    "OST_LightingDevices",
    "OST_DataDevices",
    "OST_FireAlarmDevices",
    "OST_SecurityDevices",
)
_FIXTURE_DEVICE_OPTIONAL_BIC_BY_MIN_VERSION = (
    ("OST_MechanicalControlDevices", 2024),
)


def _revit_major_version(doc=None, version=None):
    candidate = version
    if candidate is None and doc is not None:
        try:
            candidate = doc.Application.VersionNumber
        except Exception:
            candidate = None
    match = re.search(r"\d{4}", str(candidate or ""))
    return int(match.group(0)) if match else 0


def get_fixture_device_bic_names(doc=None, version=None):
    """Return fixture/device BuiltInCategory names supported by this Revit version."""
    names = list(_FIXTURE_DEVICE_BIC_NAMES)
    revit_version = _revit_major_version(doc=doc, version=version)
    for bic_name, min_version in _FIXTURE_DEVICE_OPTIONAL_BIC_BY_MIN_VERSION:
        if revit_version and revit_version < int(min_version):
            continue
        if built_in_category_from_name(bic_name) is not None:
            names.append(bic_name)
    return tuple(names)


def get_fixture_device_categories(doc=None, version=None):
    """Return fixture/device BuiltInCategory values supported by this Revit version."""
    categories = []
    for bic_name in get_fixture_device_bic_names(doc=doc, version=version):
        bic = built_in_category_from_name(bic_name)
        if bic is not None:
            categories.append(bic)
    return tuple(categories)


# Compatibility names retained for callers that use the former QC-specific API.
def get_electrical_qc_device_bic_names(version=None, doc=None):
    return get_fixture_device_bic_names(doc=doc, version=version)


def get_electrical_qc_device_categories(version=None, doc=None):
    return get_fixture_device_categories(doc=doc, version=version)


def _category_ids_from_bic_names(doc, bic_names):
    bics = []
    for bic_name in list(bic_names or []):
        bic = built_in_category_from_name(bic_name)
        if bic is not None:
            bics.append(bic)
    return category_ids_from_bics(doc, bics)


def get_circuit_category_ids(doc=None):
    return _category_ids_from_bic_names(doc, ("OST_ElectricalCircuit",))


def get_equipment_category_ids(doc=None):
    return _category_ids_from_bic_names(doc, ("OST_ElectricalEquipment",))


def get_fixture_category_ids(doc=None, version=None):
    return _category_ids_from_bic_names(
        doc,
        get_fixture_device_bic_names(doc=doc, version=version),
    )


def get_electrical_qc_device_category_ids(doc=None, version=None):
    return get_fixture_category_ids(doc=doc, version=version)


def get_all_electrical_category_ids(doc=None, version=None):
    category_ids = []
    category_ids.extend(get_circuit_category_ids(doc))
    category_ids.extend(get_equipment_category_ids(doc))
    category_ids.extend(get_fixture_category_ids(doc, version=version))
    return unique_category_ids(category_ids)


def resolve_binding_category_ids(doc, categories_value):
    """Resolve parameter-table category scopes and explicit category names."""
    resolved = []
    missing = []
    category_map = category_map_by_display_name(doc)

    for token in split_category_names(categories_value):
        normalized = normalize_category_name(token)
        token_ids = []
        if normalized in _ELECTRICAL_CIRCUIT_TOKENS:
            token_ids = get_circuit_category_ids(doc)
        elif normalized in _ALL_ELECTRICAL_TOKENS:
            token_ids = get_all_electrical_category_ids(doc)
        elif normalized in _ELECTRICAL_EQUIPMENT_TOKENS:
            token_ids = get_equipment_category_ids(doc)
        elif normalized in _ELECTRICAL_FIXTURE_TOKENS:
            token_ids = get_fixture_category_ids(doc)
        else:
            category = category_map.get(normalized)
            if category is not None:
                token_ids = [getattr(category, "Id", None)]
            else:
                bic = built_in_category_from_name(token)
                token_ids = category_ids_from_bics(doc, [bic]) if bic is not None else []

        token_ids = unique_category_ids(token_ids)
        if not token_ids:
            missing.append(str(token))
            continue
        resolved.extend(token_ids)

    return unique_category_ids(resolved), missing


def apply_writeback_filter(doc, category_ids, write_equipment_results, write_fixture_results):
    equipment_values = category_id_values(get_equipment_category_ids(doc))
    fixture_values = category_id_values(get_fixture_category_ids(doc))
    filtered = []

    for category_id in list(category_ids or []):
        value = category_id_value(category_id, default=_INVALID_CATEGORY_ID_VALUE)
        if not is_valid_category_id_value(value):
            continue
        if value in equipment_values and not bool(write_equipment_results):
            continue
        if value in fixture_values and not bool(write_fixture_results):
            continue
        filtered.append(category_id)

    return unique_category_ids(filtered)
