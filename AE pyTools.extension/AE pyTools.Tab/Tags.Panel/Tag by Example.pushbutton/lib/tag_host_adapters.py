# -*- coding: utf-8 -*-
"""Phase 1 host adapter and target validation logic."""

from Autodesk.Revit.UI.Selection import ISelectionFilter
from pyrevit import DB

from Snippets import revit_helpers
from Snippets.tag_host_transform import get_host_placement_frame
from tag_api_compat import get_single_local_reference
from tag_api_compat import get_tag_type_id
from tag_api_compat import id_value
from tag_api_compat import same_id
from tag_api_compat import valid_id


def is_supported_tag_view(view):
    """Return True for views where Phase 1 tag placement is meaningful."""
    if view is None:
        return False
    view_type = getattr(view, "ViewType", None)
    for name in ("FloorPlan", "CeilingPlan", "DraftingView"):
        supported_type = getattr(DB.ViewType, name, None)
        if supported_type is not None and view_type == supported_type:
            return True
    return False


def supported_view_description(view):
    try:
        return str(view.ViewType)
    except Exception:
        return "This view"


def category_id(element):
    try:
        return id_value(element.Category.Id)
    except Exception:
        return 0


def type_id(element):
    try:
        return id_value(element.GetTypeId())
    except Exception:
        return 0


def family_name(element, document):
    try:
        symbol = document.GetElement(element.GetTypeId())
        family = getattr(symbol, "Family", None)
        if family is not None:
            return str(family.Name)
    except Exception:
        pass
    return "<No family>"


def type_name(element, document):
    try:
        symbol = document.GetElement(element.GetTypeId())
        if symbol is not None:
            return str(DB.Element.Name.__get__(symbol))
    except Exception:
        pass
    return revit_helpers.get_family_symbol_name(element, doc=document, fallback="<No type>")


def is_multi_category_tag(document, tag):
    try:
        tag_type = document.GetElement(tag.GetTypeId())
        tag_category = tag_type.Category
        multi_category = getattr(DB.BuiltInCategory, "OST_MultiCategoryTags", None)
        if multi_category is None:
            return False
        return id_value(tag_category.Id) == int(multi_category)
    except Exception:
        return False


def host_description(element, document):
    category_name = "<No category>"
    try:
        category_name = str(element.Category.Name)
    except Exception:
        pass
    return {
        "category": category_name,
        "category_id": category_id(element),
        "family": family_name(element, document),
        "type": type_name(element, document),
        "type_id": type_id(element),
    }


def is_phase_one_host(element):
    return isinstance(element, DB.FamilyInstance)


def get_tag_owner_view(document, tag):
    """Return the view that actually owns an annotation tag."""
    owner_view_id = getattr(tag, "OwnerViewId", None)
    if not valid_id(owner_view_id):
        raise ValueError("The tag has no valid owner view.")
    owner_view = document.GetElement(owner_view_id)
    if owner_view is None:
        raise ValueError("The tag owner view is no longer valid.")
    return owner_view


def _append_view_id(view_ids, view_id):
    if not valid_id(view_id):
        return
    view_value = id_value(view_id)
    if view_value not in view_ids:
        view_ids.append(view_value)


def related_view_ids(document, view):
    """Return the primary/dependent view family for a view.

    Dependent-view annotations can be visible from a related view while their
    OwnerViewId resolves to the primary view (or, in some cases, another
    dependent view).  Keep this compatibility logic in one place.
    """
    view_ids = []
    if view is None:
        return view_ids

    _append_view_id(view_ids, view.Id)
    primary_view = view
    try:
        primary_view_id = view.GetPrimaryViewId()
    except Exception:
        primary_view_id = None
    if valid_id(primary_view_id):
        _append_view_id(view_ids, primary_view_id)
        primary_view = document.GetElement(primary_view_id)

    if primary_view is not None:
        try:
            dependent_ids = primary_view.GetDependentViewIds()
        except Exception:
            dependent_ids = []
        for dependent_id in list(dependent_ids or []):
            _append_view_id(view_ids, dependent_id)
    return view_ids


def tag_owner_is_related(document, view, tag):
    """Return whether a tag belongs to the active view family."""
    owner_view_id = getattr(tag, "OwnerViewId", None)
    if not valid_id(owner_view_id):
        return False
    owner_value = id_value(owner_view_id)
    return owner_value in related_view_ids(document, view)


def validate_example_tag(document, view, tag):
    if tag is None or not isinstance(tag, DB.IndependentTag):
        raise ValueError("Select one supported IndependentTag.")
    if not tag_owner_is_related(document, view, tag):
        raise ValueError(
            "The example tag must belong to the active view or its related "
            "primary/dependent views."
        )
    reference, host = get_single_local_reference(document, tag)
    if not is_phase_one_host(host):
        raise ValueError(
            "Phase 1 supports IndependentTag elements hosted on loadable "
            "FamilyInstance elements only."
        )
    owner_view = get_tag_owner_view(document, tag)
    frame = get_host_placement_frame(host, owner_view)
    return reference, host, frame


def is_nested_instance(element):
    try:
        return getattr(element, "SuperComponent", None) is not None
    except Exception:
        return False


def is_visible_candidate(document, view, element, example_host, include_nested):
    if element is None or not isinstance(element, DB.FamilyInstance):
        return False, "Unsupported host class."
    if same_id(element.Id, example_host.Id):
        return False, "Example host excluded."
    if is_nested_instance(element) and not include_nested:
        return False, "Nested family instance excluded by settings."
    if is_nested_instance(element) and include_nested:
        try:
            DB.Reference(element)
        except Exception:
            return False, "Nested instance has no stable project reference."
    try:
        if view.IsHidden(element.Id):
            return False, "Target is hidden in the owner view."
    except Exception:
        pass
    # View-scoped collectors and Revit selection already constrain the
    # element to this document.  Avoid Python object-identity comparisons here
    # because IronPython can wrap the same Revit document more than once.
    return True, ""


def matches_target_mode(element, example_host, mode, document):
    if mode == "type":
        return type_id(element) == type_id(example_host)
    if mode == "family":
        return family_name(element, document) == family_name(example_host, document)
    if mode == "category":
        return category_id(element) == category_id(example_host)
    return False


class ExampleTagSelectionFilter(ISelectionFilter):
    def __init__(self, document, view):
        self.document = document
        self.view = view

    def AllowElement(self, element):
        if not isinstance(element, DB.IndependentTag):
            return False
        try:
            validate_example_tag(self.document, self.view, element)
            return True
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


class TargetSelectionFilter(ISelectionFilter):
    def __init__(self, document, view, example_host, include_nested,
                 allow_any_family=False):
        self.document = document
        self.view = view
        self.example_host = example_host
        self.include_nested = include_nested
        self.allow_any_family = bool(allow_any_family)

    def AllowElement(self, element):
        allowed, unused_reason = is_visible_candidate(
            self.document,
            self.view,
            element,
            self.example_host,
            self.include_nested,
        )
        del unused_reason
        if not allowed or not is_phase_one_host(element):
            return False
        if self.allow_any_family:
            return True
        return category_id(element) == category_id(self.example_host)

    def AllowReference(self, reference, position):
        return False


def target_reference(element):
    try:
        return DB.Reference(element)
    except Exception as error:
        raise ValueError("No stable reference available: {}".format(error))


def tag_type_matches(tag, tag_type_id):
    return id_value(get_tag_type_id(tag)) == id_value(tag_type_id)
