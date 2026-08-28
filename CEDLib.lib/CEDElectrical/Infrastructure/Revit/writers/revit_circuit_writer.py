# -*- coding: utf-8 -*-
"""Revit-backed circuit writer adapter."""

from pyrevit import DB

from Snippets import categories as category_utils
from Snippets import design_options
from Snippets import revit_helpers


class RevitCircuitWriter(object):
    """Writes calculated circuit and downstream parameter values."""

    def write_circuit_parameters(self, circuit, param_values):
        """Write calculated parameter map to a circuit element."""
        if not design_options.is_main_model_element(circuit):
            return
        for param_name, value in param_values.items():
            param = circuit.LookupParameter(param_name)
            if not param:
                continue
            try:
                revit_helpers.set_parameter_if_changed(param, value)
            except Exception:
                continue

    def write_connected_elements(
            self,
            branch,
            param_values,
            settings,
            locked_ids=None,
            connected_elements=None,
    ):
        """Write calculated values to connected fixtures/devices and equipment."""
        circuit = getattr(branch, 'circuit', branch)
        fixture_count = 0
        equipment_count = 0
        locked_ids = locked_ids or set()
        if not design_options.is_main_model_element(circuit):
            return fixture_count, equipment_count
        if bool(getattr(branch, 'is_special', False)):
            return fixture_count, equipment_count

        write_fixtures = getattr(settings, 'write_fixture_results', False)
        write_equipment = getattr(settings, 'write_equipment_results', False)
        if not (write_fixtures or write_equipment):
            return fixture_count, equipment_count

        doc = getattr(circuit, 'Document', None)
        fixture_category_values = set(
            [
                category_utils.category_id_value(bic)
                for bic in category_utils.get_fixture_device_categories(doc=doc)
            ]
        )
        equipment_category_values = category_utils.category_id_values(
            category_utils.get_equipment_category_ids()
        )

        if connected_elements is None:
            try:
                elements = list(circuit.Elements)
            except Exception:
                elements = []
        else:
            elements = list(connected_elements or [])

        for el in elements:
            if not design_options.is_main_model_element(el):
                continue
            if not isinstance(el, DB.FamilyInstance):
                continue
            if el.Id in locked_ids:
                continue

            cat = el.Category
            if not cat:
                continue

            cat_id_value = category_utils.category_id_value(cat.Id)
            is_fixture = cat_id_value in fixture_category_values
            is_equipment = cat_id_value in equipment_category_values

            if not (is_fixture or is_equipment):
                continue
            if is_fixture and not write_fixtures:
                continue
            if is_equipment and not write_equipment:
                continue

            for param_name, value in param_values.items():
                if value is None:
                    continue
                param = el.LookupParameter(param_name)
                if not param:
                    continue
                try:
                    revit_helpers.set_parameter_if_changed(param, value)
                except Exception:
                    continue

            if is_fixture:
                fixture_count += 1
            elif is_equipment:
                equipment_count += 1

        return fixture_count, equipment_count
