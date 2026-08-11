# -*- coding: utf-8 -*-
"""Read-only location extraction for supported element classes."""

from __future__ import print_function

import math

import models

try:
    from pyrevit import DB
except Exception:
    DB = None


def read_location(element, transform=None):
    """Location of an element as a JSON-safe dict.

    ``transform`` (a Revit Transform, e.g. a link instance's total
    transform) converts a linked element's local point/rotation into
    host/world coordinates so deltas against host elements are valid and
    a moved link instance registers as a moved element.
    """
    if element is None:
        return {
            "state": models.VALUE_UNSUPPORTED,
            "message": "Location Not Supported",
        }
    try:
        location = element.Location
    except Exception:
        location = None
    if DB is None or location is None or not isinstance(location, DB.LocationPoint):
        return {
            "state": models.VALUE_UNSUPPORTED,
            "message": "Location Not Supported",
        }
    try:
        point = location.Point
        rotation = float(location.Rotation)
        coordinate_system = "source_document_internal"
        if transform is not None:
            point = transform.OfPoint(point)
            facing = transform.OfVector(
                DB.XYZ(math.cos(rotation), math.sin(rotation), 0.0)
            )
            rotation = math.atan2(float(facing.Y), float(facing.X))
            coordinate_system = "host_world"
        return {
            "state": models.VALUE_VALID,
            "x": float(point.X),
            "y": float(point.Y),
            "z": float(point.Z),
            "rotation": rotation,
            "coordinate_system": coordinate_system,
        }
    except Exception as ex:
        return {
            "state": models.VALUE_READ_ERROR,
            "message": str(ex),
        }

