# -*- coding: utf-8 -*-
"""Read-only location extraction for supported element classes."""

from __future__ import print_function

import models

try:
    from pyrevit import DB
except Exception:
    DB = None


def read_location(element):
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
        return {
            "state": models.VALUE_VALID,
            "x": float(point.X),
            "y": float(point.Y),
            "z": float(point.Z),
            "rotation": float(location.Rotation),
            "coordinate_system": "source_document_internal",
        }
    except Exception as ex:
        return {
            "state": models.VALUE_READ_ERROR,
            "message": str(ex),
        }

