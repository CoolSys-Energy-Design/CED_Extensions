# -*- coding: utf-8 -*-
"""ForgeTypeId-backed Revit unit conversion helpers."""

import Autodesk.Revit.DB as DB


def project_unit_id(doc, spec_type_id):
    """Return the project's display unit for a Forge spec."""
    options = doc.GetUnits().GetFormatOptions(spec_type_id)
    return options.GetUnitTypeId()


def display_to_internal(doc, value, spec_type_id):
    """Convert a project-display value to Revit internal units."""
    unit_id = project_unit_id(doc, spec_type_id)
    return float(DB.UnitUtils.ConvertToInternalUnits(float(value), unit_id))


def display_to_unit(doc, value, spec_type_id, target_unit_type_id):
    """Convert a project-display value to a target Forge unit."""
    internal = display_to_internal(doc, value, spec_type_id)
    return float(DB.UnitUtils.ConvertFromInternalUnits(internal, target_unit_type_id))


def electrical_potential_display_to_volts(doc, value):
    """Convert a project-formatted electrical-potential number to volts."""
    return display_to_unit(
        doc,
        value,
        DB.SpecTypeId.ElectricalPotential,
        DB.UnitTypeId.Volts,
    )


def internal_to_unit(value, unit_type_id):
    """Convert a Revit internal value to a Forge target unit."""
    return float(DB.UnitUtils.ConvertFromInternalUnits(float(value), unit_type_id))
