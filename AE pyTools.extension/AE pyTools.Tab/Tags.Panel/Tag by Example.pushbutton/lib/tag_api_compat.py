# -*- coding: utf-8 -*-
"""Small compatibility layer for tag and reference APIs."""

from pyrevit import DB

from Snippets import revit_helpers


def id_value(element_id):
    return revit_helpers.get_elementid_value(element_id)


def same_id(first_id, second_id):
    return id_value(first_id) == id_value(second_id)


def valid_id(element_id):
    return element_id is not None and id_value(element_id) != id_value(DB.ElementId.InvalidElementId)


def get_tagged_local_ids(tag):
    getter = getattr(tag, "GetTaggedLocalElementIds", None)
    if getter is not None:
        try:
            return list(getter() or [])
        except Exception:
            pass
    local_id = getattr(tag, "TaggedLocalElementId", None)
    if valid_id(local_id):
        return [local_id]
    return []


def get_tagged_references(tag):
    getter = getattr(tag, "GetTaggedReferences", None)
    if getter is None:
        return []
    try:
        return list(getter() or [])
    except Exception:
        return []


def get_single_local_reference(doc, tag):
    """Return ``(reference, host)`` for one local, resolvable host."""
    references = get_tagged_references(tag)
    local_ids = get_tagged_local_ids(tag)
    if len(references) == 0 and len(local_ids) == 1:
        local_host = doc.GetElement(local_ids[0])
        if local_host is not None:
            try:
                references = [DB.Reference(local_host)]
            except Exception:
                references = []
    if len(references) != 1 or len(local_ids) != 1:
        raise ValueError("Only single-reference local tags are supported.")

    reference = references[0]
    host_id = getattr(reference, "ElementId", None)
    if not valid_id(host_id):
        raise ValueError("Tag reference is not a valid local reference.")
    if not same_id(host_id, local_ids[0]):
        raise ValueError("Tag reference does not resolve to the tagged local element.")
    host = doc.GetElement(host_id)
    if host is None:
        raise ValueError("The tag host no longer exists.")

    linked_id = getattr(reference, "LinkedElementId", None)
    if linked_id is not None and valid_id(linked_id):
        raise ValueError("Linked-model references are not supported in Phase 1.")
    return reference, host


def get_owner_view_id(tag):
    return getattr(tag, "OwnerViewId", None)


def get_tag_type_id(tag):
    try:
        return tag.GetTypeId()
    except Exception:
        return None


def get_leader_snapshot(tag, reference):
    snapshot = {
        "has_leader": False,
        "end_condition": None,
        "elbow": None,
        "end": None,
    }
    try:
        snapshot["has_leader"] = bool(tag.HasLeader)
    except Exception:
        return snapshot
    if not snapshot["has_leader"]:
        return snapshot

    snapshot["end_condition"] = getattr(tag, "LeaderEndCondition", None)
    try:
        getter = getattr(tag, "GetLeaderElbow", None)
        if getter is not None:
            snapshot["elbow"] = getter(reference)
        elif bool(getattr(tag, "HasElbow", False)):
            snapshot["elbow"] = tag.LeaderElbow
    except Exception:
        snapshot["elbow"] = None

    try:
        getter = getattr(tag, "GetLeaderEnd", None)
        if getter is not None:
            snapshot["end"] = getter(reference)
        else:
            snapshot["end"] = tag.LeaderEnd
    except Exception:
        snapshot["end"] = None
    return snapshot


def create_independent_tag(document, view_id, tag_type_id, reference,
                           has_leader, orientation, head_position):
    """Create a tag using repository and version-compatible overloads."""
    creator = DB.IndependentTag.Create
    errors = []

    # This is the overload used by the repository's working tag placement
    # code: (document, tag type, owner view, reference, leader, orientation,
    # point).
    try:
        return creator(document, tag_type_id, view_id, reference,
                       bool(has_leader), orientation, head_position)
    except Exception as error:
        errors.append(error)

    # Some supported versions expose the type-aware overload with the view
    # before the tag type.
    try:
        return creator(document, view_id, tag_type_id, reference,
                       bool(has_leader), orientation, head_position)
    except Exception as error:
        errors.append(error)

    # Newer APIs may expose the tag type as the final argument alongside the
    # older TagMode signature.
    try:
        return creator(document, view_id, reference, bool(has_leader),
                       DB.TagMode.TM_ADDBY_CATEGORY, orientation,
                       head_position, tag_type_id)
    except Exception as error:
        errors.append(error)

    try:
        created_tag = creator(
            document,
            view_id,
            reference,
            bool(has_leader),
            DB.TagMode.TM_ADDBY_CATEGORY,
            orientation,
            head_position,
        )
        if valid_id(tag_type_id):
            created_tag.ChangeTypeId(tag_type_id)
        return created_tag
    except Exception as fallback_error:
        raise RuntimeError(
            "IndependentTag.Create failed for supported overloads: {}; "
            "fallback: {}".format(" | ".join([str(item) for item in errors]),
                                   fallback_error)
        )


def set_leader_end_condition(tag, condition):
    if condition is None:
        return False
    try:
        tag.LeaderEndCondition = condition
        return True
    except Exception:
        return False


def set_leader_elbow(tag, reference, point):
    if point is None:
        return False
    try:
        setter = getattr(tag, "SetLeaderElbow", None)
        if setter is not None:
            setter(reference, point)
        else:
            tag.LeaderElbow = point
        return True
    except Exception:
        return False


def set_leader_end(tag, reference, point):
    if point is None:
        return False
    try:
        setter = getattr(tag, "SetLeaderEnd", None)
        if setter is not None:
            setter(reference, point)
        else:
            tag.LeaderEnd = point
        return True
    except Exception:
        return False


def get_tag_orientation(tag):
    return getattr(tag, "TagOrientation", None)


def get_rotation_angle(tag):
    try:
        return float(tag.RotationAngle)
    except Exception:
        return None
