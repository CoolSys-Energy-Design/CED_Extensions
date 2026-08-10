# -*- coding: utf-8 -*-
"""Tracked element to host electrical device/circuit context."""

from __future__ import print_function

try:
    from pyrevit import DB
    import Autodesk.Revit.DB.Electrical as DBE
    from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
except Exception:
    DB = None
    DBE = None
    ISelectionFilter = object
    ObjectType = None

from Snippets import revit_helpers

try:
    from Snippets import _elecutils
except Exception:
    _elecutils = None


def _id_value(value):
    return revit_helpers.get_elementid_value(value)


class ElectricalDeviceSelectionFilter(ISelectionFilter):
    def __init__(self):
        allowed = []
        if DB is not None:
            allowed = [
                DB.BuiltInCategory.OST_ElectricalEquipment,
                DB.BuiltInCategory.OST_ElectricalFixtures,
                DB.BuiltInCategory.OST_LightingFixtures,
                DB.BuiltInCategory.OST_MechanicalEquipment,
            ]
        self._allowed = set([
            _id_value(DB.ElementId(item)) for item in allowed
        ]) if DB is not None else set()

    def AllowElement(self, element):  # noqa: N802
        category = getattr(element, "Category", None)
        return category is not None and _id_value(category.Id) in self._allowed

    def AllowReference(self, reference, position):  # noqa: N802
        return False


def pick_device(uidocument):
    if uidocument is None or ObjectType is None:
        return None
    try:
        reference = uidocument.Selection.PickObject(
            ObjectType.Element,
            ElectricalDeviceSelectionFilter(),
            "Pick a host electrical device or equipment element",
        )
    except Exception as ex:
        # Revit reports Escape during PickObject as OperationCanceledException
        # (and some pyRevit engines surface it as a generic aborted-pick error).
        type_name = ""
        try:
            type_name = str(ex.GetType().FullName or "")
        except Exception:
            type_name = ex.__class__.__name__
        message = str(ex or "").lower()
        if "operationcanceled" in type_name.lower() or "aborted the pick" in message:
            return None
        raise
    return uidocument.Document.GetElement(reference.ElementId)


def relationship_from_device(device):
    if device is None:
        return None
    return {
        "device_unique_id": str(getattr(device, "UniqueId", "") or ""),
        "device_id": _id_value(getattr(device, "Id", None)),
        "device_name": str(getattr(device, "Name", "") or "Electrical Device"),
    }


def _candidate_circuits(device):
    if device is None:
        return []
    candidates = []
    if DBE is not None and isinstance(device, DBE.ElectricalSystem):
        candidates.append(device)
    mep_model = getattr(device, "MEPModel", None)
    if mep_model is not None:
        for method_name in ("GetElectricalSystems", "GetAssignedElectricalSystems"):
            try:
                candidates.extend(list(getattr(mep_model, method_name)() or []))
            except Exception:
                pass
        try:
            candidates.extend(list(getattr(mep_model, "ElectricalSystems", None) or []))
        except Exception:
            pass
    if _elecutils is not None:
        try:
            candidates = list(_elecutils.filter_circuits(candidates))
        except Exception:
            pass
    unique = []
    seen = set()
    for candidate in candidates:
        value = _id_value(getattr(candidate, "Id", None))
        if value <= 0 or value in seen:
            continue
        seen.add(value)
        unique.append(candidate)
    return unique


def circuit_context(device):
    if device is None:
        return {
            "status": "device_missing",
            "status_text": "Linked Device Missing",
            "circuits": [],
        }
    circuits = []
    for circuit in _candidate_circuits(device):
        panel = getattr(circuit, "BaseEquipment", None)
        circuits.append({
            "circuit_id": _id_value(getattr(circuit, "Id", None)),
            "circuit_unique_id": str(getattr(circuit, "UniqueId", "") or ""),
            "circuit_number": str(getattr(circuit, "CircuitNumber", "") or ""),
            "circuit_name": str(getattr(circuit, "Name", "") or ""),
            "panel_name": str(getattr(panel, "Name", "") or ""),
        })
    if not circuits:
        status = "device_not_circuited"
        status_text = "Device Linked - Not Circuited"
    elif len(circuits) == 1:
        status = "circuited"
        circuit = circuits[0]
        status_text = "{} - {}".format(
            circuit.get("panel_name") or "Panel",
            circuit.get("circuit_number") or circuit.get("circuit_name") or "Circuit",
        )
    else:
        status = "multiple_circuits"
        status_text = "Device Linked - {} Circuits".format(len(circuits))
    return {"status": status, "status_text": status_text, "circuits": circuits}


def resolve_relationship(host_document, relationship):
    relationship = relationship or {}
    unique_id = str(relationship.get("device_unique_id") or "")
    device = None
    if host_document is not None and unique_id:
        try:
            device = host_document.GetElement(unique_id)
        except Exception:
            device = None
    context = circuit_context(device)
    context["device_unique_id"] = unique_id
    context["device_id"] = _id_value(getattr(device, "Id", None)) if device is not None else 0
    context["device_name"] = str(getattr(device, "Name", "") or relationship.get("device_name") or "")
    return device, context
