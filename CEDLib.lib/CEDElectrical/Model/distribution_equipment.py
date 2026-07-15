# -*- coding: utf-8 -*-
"""Revit-agnostic distribution equipment models."""

from CEDElectrical.part_types import PART_TYPE_OTHER_PANEL, PART_TYPE_PANELBOARD, PART_TYPE_SWITCHBOARD

PANEL_OR_SWITCHGEAR_PART_TYPES = set([PART_TYPE_PANELBOARD, PART_TYPE_SWITCHBOARD, PART_TYPE_OTHER_PANEL])


class DistributionEquipment(object):
    """Domain model for electrical distribution equipment metadata."""

    def __init__(self, **kwargs):
        self.id = int(kwargs.get("id", 0) or 0)
        self.name = kwargs.get("name")
        self.element_name = kwargs.get("element_name")
        self.panel_name = kwargs.get("panel_name")
        self.equipment_type = kwargs.get("equipment_type")
        self.part_type = kwargs.get("part_type")
        self.parameter_values = dict(kwargs.get("parameter_values") or {})
        self.parameter_sources = dict(kwargs.get("parameter_sources") or {})

        self.voltage = kwargs.get("voltage")
        self.poles = kwargs.get("poles")
        self.distribution_system = kwargs.get("distribution_system")
        self.distribution_system_secondary = kwargs.get("distribution_system_secondary")

        self.supply_connections = list(kwargs.get("supply_connections") or [])
        self.supply_circuits = list(kwargs.get("supply_circuits") or [])
        if not self.supply_circuits and self.supply_connections:
            self.supply_circuits = [
                int(x.get("circuit_id") or 0)
                for x in self.supply_connections
                if int(x.get("circuit_id") or 0) > 0
            ]
        self.branch_circuits = list(kwargs.get("branch_circuits") or [])
        self.branch_circuit_options = list(kwargs.get("branch_circuit_options") or [])

        self.mains_rating = kwargs.get("mains_rating")
        self.mains_type = kwargs.get("mains_type")
        self.has_ocp = kwargs.get("has_ocp")
        self.ocp_type = kwargs.get("ocp_type")
        self.ocp_rating = kwargs.get("ocp_rating")

        self.has_feed_thru_lugs = kwargs.get("has_feed_thru_lugs")
        self.has_neutral_bus = kwargs.get("has_neutral_bus")
        self.has_ground_bus = kwargs.get("has_ground_bus")
        self.has_isolated_ground_bus = kwargs.get("has_isolated_ground_bus")

        self.max_poles = kwargs.get("max_poles")
        self.short_circuit_rating = kwargs.get("short_circuit_rating")

        self.power_connected_total = kwargs.get("power_connected_total")
        self.current_connected_total = kwargs.get("current_connected_total")
        self.power_demand_total = kwargs.get("power_demand_total")
        self.current_demand_total = kwargs.get("current_demand_total")
        self.branch_current_phase_a = kwargs.get("branch_current_phase_a")
        self.branch_current_phase_b = kwargs.get("branch_current_phase_b")
        self.branch_current_phase_c = kwargs.get("branch_current_phase_c")
        self.branch_load_phase_a = kwargs.get("branch_load_phase_a")
        self.branch_load_phase_b = kwargs.get("branch_load_phase_b")
        self.branch_load_phase_c = kwargs.get("branch_load_phase_c")

    @property
    def ID(self):
        """Compatibility alias for id."""
        return self.id

    @property
    def Name(self):
        """Compatibility alias for name."""
        return self.name

    @property
    def EquipmentType(self):
        """Compatibility alias for equipment_type."""
        return self.equipment_type

    @property
    def is_mlo(self):
        """Return True when mains type indicates main-lugs-only equipment."""
        text = str(self.mains_type or "").strip().upper()
        return "MLO" in text

    @property
    def is_panel_or_switchgear(self):
        """Return True when equipment is a panelboard/switchboard-style bus."""
        return self.part_type in PANEL_OR_SWITCHGEAR_PART_TYPES

    @property
    def supply_circuit_quantity(self):
        """Return count of assigned supply circuits."""
        return len([x for x in list(self.supply_circuits or []) if int(x or 0) > 0])

    @property
    def primary_supply(self):
        """Return the connector-backed primary supply record, when known."""
        for connection in list(self.supply_connections or []):
            if bool(connection.get("is_primary")):
                return connection
        if self.supply_connections:
            return self.supply_connections[0]
        if self.supply_circuits:
            return {"circuit_id": int(self.supply_circuits[0] or 0), "is_primary": True}
        return None

    @property
    def primary_supply_circuit_id(self):
        supply = self.primary_supply
        if not supply:
            return 0
        try:
            return int(supply.get("circuit_id") or 0)
        except Exception:
            return 0

    @property
    def secondary_supplies(self):
        """Return non-primary supply records."""
        secondary = []
        for connection in list(self.supply_connections or []):
            if bool(connection.get("is_primary")):
                continue
            secondary.append(connection)
        if secondary:
            return secondary
        primary_id = self.primary_supply_circuit_id
        return [
            {"circuit_id": int(circuit_id or 0), "is_primary": False}
            for circuit_id in list(self.supply_circuits or [])
            if int(circuit_id or 0) > 0 and int(circuit_id or 0) != int(primary_id or 0)
        ]

    @property
    def secondary_supply_circuit_ids(self):
        ids = []
        for supply in list(self.secondary_supplies or []):
            try:
                circuit_id = int(supply.get("circuit_id") or 0)
            except Exception:
                circuit_id = 0
            if circuit_id > 0:
                ids.append(circuit_id)
        return ids

    @property
    def has_multiple_supplies(self):
        return self.supply_circuit_quantity > 1

    def distribution_voltage_for_poles(self, poles, secondary=False):
        """Return LN voltage for 1P and LL voltage for 2P/3P from distribution profile."""
        profile = self.distribution_system_secondary if bool(secondary) else self.distribution_system
        if not profile:
            return None
        try:
            pole_count = int(poles or 1)
        except Exception:
            pole_count = 1
        key = "lg_voltage" if pole_count <= 1 else "ll_voltage"
        value = profile.get(key)
        if value is not None:
            return value
        return profile.get("ll_voltage") or profile.get("lg_voltage")

    def to_dict(self):
        """Serialize model data for UI/report consumers."""
        return dict(self.__dict__)


class Transformer(DistributionEquipment):
    """Distribution equipment specialization for transformers."""

    def __init__(self, **kwargs):
        DistributionEquipment.__init__(self, **kwargs)
        self.xfmr_rating = kwargs.get("xfmr_rating")
        self.xfmr_impedance = kwargs.get("xfmr_impedance")
        self.xfmr_kfactor = kwargs.get("xfmr_kfactor")


class PowerBus(DistributionEquipment):
    """Distribution equipment specialization for panel/switch/data buses."""

    def __init__(self, **kwargs):
        DistributionEquipment.__init__(self, **kwargs)
        self.has_panel_schedule = bool(kwargs.get("has_panel_schedule", False))
        self.panel_configuration = kwargs.get("panel_configuration")

    def has_panelschedule(self):
        """Return True when this bus has a panel schedule instance in the model."""
        return bool(self.has_panel_schedule)
