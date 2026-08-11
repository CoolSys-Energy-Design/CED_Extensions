# -*- coding: utf-8 -*-
"""Move monitored children by their monitored parent's location delta.

The transform mechanic mirrors the MEPRFP Follow Parent tool: translate by
the parent's (current - accepted) vector, then rotate about the parent's
current point Z-axis by the rotation delta. After the move the child's own
location is auto-accepted (the move was deliberate) and its Element_Linker
pose fields are rewritten so the MEPRFP tools stay consistent.
"""

from __future__ import print_function

import copy

try:
    from pyrevit import DB
except Exception:
    DB = None

import comparison_engine
import location_service
import mep_linker_bridge
import models
import parent_move_math
import tracking_service


def _location_tolerances(tracking_set):
    defaults = (tracking_set or {}).get("location_defaults") or {}
    return (
        float(defaults.get("translation_tolerance", 0.001) or 0.001),
        float(defaults.get("angular_tolerance", 0.0017453292519943296)
              or 0.0017453292519943296),
    )


def _resolve_child_element(document, record):
    unique_id = str((record or {}).get("source_element_unique_id") or "")
    if not unique_id:
        return None
    try:
        return document.GetElement(unique_id)
    except Exception:
        return None


def move_children_with_parent(document, store, set_id, persistent_ids, logger=None):
    """Apply the parent follow-move to the selected child records.

    Returns ``(updated_store, message)``. Raises ValueError when nothing
    is movable.
    """
    tracking_set = copy.deepcopy(models.find_set(store, set_id))
    if tracking_set is None:
        raise ValueError("Select a Tracking Set first.")
    elements = tracking_set.get("elements") or {}
    translation_tolerance, angular_tolerance = _location_tolerances(tracking_set)

    moves = []
    warnings = []
    for persistent_id in list(persistent_ids or []):
        key = str(persistent_id or "")
        record = elements.get(key)
        if record is None or record.get("state") == models.ELEMENT_REMOVED:
            warnings.append("{}: record unavailable.".format(key))
            continue
        parent_key = str(record.get("parent_persistent_id") or "")
        parent_record = elements.get(parent_key) if parent_key else None
        if parent_record is None:
            warnings.append(
                "{}: no monitored Element Linker parent.".format(_label(record, key))
            )
            continue
        if comparison_engine.locations_equal(
            parent_record.get("accepted_location"),
            parent_record.get("current_location"),
            translation_tolerance,
            angular_tolerance,
        ):
            warnings.append(
                "{}: parent has not moved.".format(_label(record, key))
            )
            continue
        delta = parent_move_math.compute_follow_delta(
            parent_record.get("accepted_location"),
            parent_record.get("current_location"),
        )
        if not parent_move_math.significant_delta(
            delta, translation_tolerance, angular_tolerance
        ):
            warnings.append(
                "{}: parent move is below tolerance.".format(_label(record, key))
            )
            continue
        element = _resolve_child_element(document, record)
        if element is None:
            warnings.append(
                "{}: element no longer exists in the model.".format(_label(record, key))
            )
            continue
        moves.append((key, record, parent_key, parent_record, element, delta))

    if not moves:
        raise ValueError(
            "None of the selected elements have a monitored parent with a "
            "pending move.\n" + "\n".join(warnings)
        )

    moved_keys = []
    transaction = DB.Transaction(document, "Parameter Monitor - Move with Parent")
    transaction.Start()
    try:
        for key, record, _parent_key, parent_record, element, delta in moves:
            try:
                dx, dy, dz = delta["translation"]
                if any(abs(component) > 1e-12 for component in (dx, dy, dz)):
                    DB.ElementTransformUtils.MoveElement(
                        document, element.Id, DB.XYZ(dx, dy, dz)
                    )
                rotation_delta = float(delta.get("rotation_delta") or 0.0)
                if abs(rotation_delta) > 1e-12:
                    px, py, pz = delta["pivot"]
                    axis = DB.Line.CreateBound(
                        DB.XYZ(px, py, pz), DB.XYZ(px, py, pz + 1.0)
                    )
                    DB.ElementTransformUtils.RotateElement(
                        document, element.Id, axis, rotation_delta
                    )
            except Exception as ex:
                warnings.append(
                    "{}: move failed ({}).".format(_label(record, key), ex)
                )
                continue
            new_location = location_service.read_location(element)
            try:
                mep_linker_bridge.update_linker(
                    element,
                    parent_move_math.linker_pose_updates(
                        new_location, parent_record.get("current_location")
                    ),
                )
            except Exception as ex:
                warnings.append(
                    "{}: Element_Linker pose not updated ({}).".format(
                        _label(record, key), ex
                    )
                )
            moved_keys.append((key, new_location))
        transaction.Commit()
    except Exception:
        try:
            transaction.RollBack()
        except Exception:
            pass
        raise

    if not moved_keys:
        raise ValueError(
            "No elements could be moved.\n" + "\n".join(warnings)
        )

    accepted_parents = []
    touched_parent_keys = set()
    for key, new_location in moved_keys:
        record = elements.get(key)
        record["current_location"] = copy.deepcopy(new_location)
        record["accepted_location"] = copy.deepcopy(new_location)
        comparison_engine.recompute_record(record, tracking_set)
        touched_parent_keys.add(str(record.get("parent_persistent_id") or ""))
    for parent_key in touched_parent_keys:
        if not parent_key:
            continue
        if parent_move_math.accept_parent_location_if_children_clean(
            tracking_set, parent_key
        ):
            parent_record = elements.get(parent_key) or {}
            accepted_parents.append(_label(parent_record, parent_key))

    comparison_engine.refresh_set_status(tracking_set)
    tracking_set["updated_at"] = models.utc_now_text()
    updated_store = tracking_service._replace_set(store, tracking_set)

    lines = ["Moved {} element(s) with their monitored parent.".format(len(moved_keys))]
    if accepted_parents:
        lines.append(
            "Parent location auto-accepted (all children in sync): {}.".format(
                ", ".join(accepted_parents)
            )
        )
    for warning in warnings:
        lines.append("Warning: {}".format(warning))
    if logger is not None:
        try:
            logger.info(
                "Parameter Monitor move-with-parent: %s moved, %s parents accepted",
                len(moved_keys), len(accepted_parents),
            )
        except Exception:
            pass
    return updated_store, "\n".join(lines)


def _label(record, fallback):
    metadata = (record or {}).get("metadata") or {}
    return str(metadata.get("friendly_name") or fallback)
