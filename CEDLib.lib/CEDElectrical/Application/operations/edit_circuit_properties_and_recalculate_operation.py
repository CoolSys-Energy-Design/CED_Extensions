# -*- coding: utf-8 -*-
"""Apply staged circuit property edits, then run calculate operation."""

from datetime import datetime

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Model.circuit_settings import CircuitVDMethod
from Snippets import design_options, revit_helpers

CIRCUIT_NOTES_KEY = "__bip_circuit_notes__"
CIRCUIT_NAME_KEY = "__bip_circuit_name__"
CIRCUIT_DATA_VD_METHOD_KEY = "circuit_vd_method"


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


def _calc_options_from_request(request):
    options = {
        "show_output": bool(request.options.get("show_output", False)),
        "use_existing_transaction_group": True,
        "calc_preview_enabled": False,
        "transaction_name": "Edit Circuit Properties + Calculate Circuits",
        "write_transaction_name": "Edit Circuit Properties + Write Circuit Parameters",
    }
    return options


class EditCircuitPropertiesAndRecalculateOperation(object):
    """Writes staged user edits on circuits, then recalculates affected circuits."""

    key = "edit_circuit_properties_and_recalculate"

    def __init__(self, calculate_operation):
        self._calculate_operation = calculate_operation
        self._alert_store = getattr(calculate_operation, "alert_store", None)

    def execute(self, request, doc):
        circuits = self._get_target_circuits(doc, request.circuit_ids)
        if not circuits:
            return {"status": "cancelled", "reason": "no_circuits"}

        updates_by_id = self._normalize_updates(request.options.get("updates"))
        if not updates_by_id:
            return {"status": "cancelled", "reason": "no_updates"}

        changed_ids = []
        locked_rows = []

        tg = DB.TransactionGroup(doc, "Edit Circuit Properties + Calculate Circuits")
        tg.Start()
        tx = DB.Transaction(doc, "Edit Circuit Properties")
        tx.Start()
        try:
            for circuit in circuits:
                cid = _elid_value(circuit.Id)
                update = dict(updates_by_id.get(cid) or {})
                param_values = dict(update.get("param_values") or {})
                circuit_data_values = dict(update.get("circuit_data_values") or {})
                force_recalculate = bool(update.get("force_recalculate", False))
                if not param_values and not circuit_data_values and not force_recalculate:
                    continue

                if self._is_locked(doc, circuit.Id):
                    locked_rows.append(self._locked_row(circuit, doc))
                    continue

                did_change = False
                for param_name, value in list(param_values.items()):
                    did_change = self._set_param_value(circuit, param_name, value) or did_change
                if circuit_data_values:
                    did_change = self._set_circuit_data_values(circuit, circuit_data_values) or did_change
                if did_change or force_recalculate:
                    changed_ids.append(cid)

            tx.Commit()
        except Exception:
            tx.RollBack()
            try:
                tg.RollBack()
            except Exception:
                pass
            raise

        if not changed_ids:
            try:
                tg.RollBack()
            except Exception:
                pass
            return {
                "status": "cancelled",
                "reason": "no_changes",
                "locked_rows": locked_rows,
                "runtime_alert_rows": [],
            }

        calc_request = OperationRequest(
            operation_key="calculate_circuits",
            circuit_ids=changed_ids,
            source=request.source,
            options=_calc_options_from_request(request),
        )
        try:
            calc_result = self._calculate_operation.execute(calc_request, doc) or {}
            if (
                str(calc_result.get("status") or "").strip().lower() == "cancelled"
                and str(calc_result.get("reason") or "").strip().lower() == "calc_preview_skipped"
            ):
                try:
                    tg.RollBack()
                except Exception:
                    pass
                if locked_rows:
                    existing = list(calc_result.get("locked_rows") or [])
                    calc_result["locked_rows"] = existing + locked_rows
                calc_result["edited_circuits"] = len(changed_ids)
                return calc_result
            if str(calc_result.get("status") or "").strip().lower() == "preview_required":
                try:
                    tg.RollBack()
                except Exception:
                    pass
                if locked_rows:
                    existing = list(calc_result.get("locked_rows") or [])
                    calc_result["locked_rows"] = existing + locked_rows
                calc_result["edited_circuits"] = len(changed_ids)
                return calc_result
            tg.Assimilate()
        except Exception:
            try:
                tg.RollBack()
            except Exception:
                pass
            raise

        if locked_rows:
            existing = list(calc_result.get("locked_rows") or [])
            calc_result["locked_rows"] = existing + locked_rows
        calc_result["edited_circuits"] = len(changed_ids)
        return calc_result

    def _normalize_updates(self, updates):
        by_id = {}
        for row in list(updates or []):
            if not isinstance(row, dict):
                continue
            try:
                cid = int(row.get("circuit_id") or 0)
            except Exception:
                cid = 0
            if cid <= 0:
                continue
            param_values = dict(row.get("param_values") or {})
            circuit_data_values = dict(row.get("circuit_data_values") or {})
            force_recalculate = bool(row.get("force_recalculate", False))
            if not param_values and not circuit_data_values and not force_recalculate:
                continue
            by_id[cid] = {
                "param_values": param_values,
                "circuit_data_values": circuit_data_values,
                "force_recalculate": force_recalculate,
            }
        return by_id

    def _set_circuit_data_values(self, circuit, values):
        if self._alert_store is None:
            return False
        payload = self._alert_store.read_alert_payload(circuit) or {}
        if not isinstance(payload, dict):
            payload = {}
        changed = False

        if CIRCUIT_DATA_VD_METHOD_KEY in values:
            current = CircuitVDMethod.normalize(payload.get(CIRCUIT_DATA_VD_METHOD_KEY), CircuitVDMethod.GLOBAL)
            target = CircuitVDMethod.normalize(values.get(CIRCUIT_DATA_VD_METHOD_KEY), CircuitVDMethod.GLOBAL)
            if current != target:
                payload[CIRCUIT_DATA_VD_METHOD_KEY] = target
                changed = True

        if not changed:
            return False

        alerts = payload.get("alerts")
        payload["alerts"] = alerts if isinstance(alerts, list) else []
        hidden = payload.get("hidden_definition_ids")
        payload["hidden_definition_ids"] = hidden if isinstance(hidden, list) else []
        payload["version"] = payload.get("version") or 1
        payload["updated_utc"] = datetime.utcnow().isoformat() + "Z"
        return bool(self._alert_store.write_alert_payload(circuit, payload))

    def _get_target_circuits(self, doc, circuit_ids):
        circuits = []
        for raw_id in list(circuit_ids or []):
            try:
                el = doc.GetElement(_elid_from_value(raw_id))
            except Exception:
                el = None
            if isinstance(el, DBE.ElectricalSystem) and design_options.is_main_model_element(el):
                circuits.append(el)
        return circuits

    def _set_param_value(self, circuit, param_name, value):
        key_text = str(param_name or "")
        changed = False

        # Rating/Frame edits must update true circuit properties + built-ins,
        # not just shared parameter mirrors.
        if key_text == "CKT_Rating_CED":
            changed = self._set_circuit_numeric(
                circuit,
                "Rating",
                DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
                value,
            ) or changed
            return changed
        elif key_text == "CKT_Frame_CED":
            changed = self._set_circuit_numeric(
                circuit,
                "Frame",
                DB.BuiltInParameter.RBS_ELEC_CIRCUIT_FRAME_PARAM,
                value,
            ) or changed
            return changed

        if key_text == CIRCUIT_NOTES_KEY:
            try:
                param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
            except Exception:
                param = None
        elif key_text == CIRCUIT_NAME_KEY:
            try:
                param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
            except Exception:
                param = None
        else:
            try:
                param = circuit.LookupParameter(key_text)
            except Exception:
                param = None
        if not param:
            return changed

        try:
            storage_type = param.StorageType
            if storage_type == DB.StorageType.Integer:
                current = param.AsInteger()
                target = int(round(float(value or 0)))
            elif storage_type == DB.StorageType.Double:
                current = param.AsDouble()
                target = float(value or 0.0)
                if abs(float(current) - float(target)) < 0.000001:
                    return changed
            elif storage_type == DB.StorageType.String:
                current = param.AsString() or ""
                target = str(value or "")
            else:
                return changed
        except Exception:
            return changed

        try:
            if storage_type == DB.StorageType.Integer:
                if int(current) == int(target):
                    return changed
                return bool(param.Set(int(target))) or changed
            if storage_type == DB.StorageType.Double:
                return bool(param.Set(float(target))) or changed
            if storage_type == DB.StorageType.String:
                if str(current or "") == str(target or ""):
                    return changed
                return bool(param.Set(str(target))) or changed
        except Exception:
            return changed
        return changed

    def _set_circuit_numeric(self, circuit, prop_name, bip, value):
        try:
            numeric = float(value)
        except Exception:
            return False

        changed = False
        try:
            current = getattr(circuit, prop_name)
            if current is None or abs(float(current) - numeric) > 0.0001:
                setattr(circuit, prop_name, numeric)
                changed = True
        except Exception:
            pass

        try:
            param = circuit.get_Parameter(bip)
        except Exception:
            param = None
        if param:
            try:
                if param.StorageType == DB.StorageType.Double:
                    cur = param.AsDouble()
                    if cur is None or abs(float(cur) - numeric) > 0.0001:
                        param.Set(float(numeric))
                        changed = True
                elif param.StorageType == DB.StorageType.Integer:
                    iv = int(round(numeric))
                    cur = param.AsInteger()
                    if cur != iv:
                        param.Set(iv)
                        changed = True
            except Exception:
                pass
        return changed

    def _is_locked(self, doc, eid):
        if not getattr(doc, "IsWorkshared", False):
            return False
        try:
            return DB.WorksharingUtils.GetCheckoutStatus(doc, eid) == DB.CheckoutStatus.OwnedByOtherUser
        except Exception:
            return False

    def _locked_row(self, circuit, doc):
        panel = ""
        try:
            panel = circuit.BaseEquipment.Name if circuit.BaseEquipment else ""
        except Exception:
            panel = ""
        number = ""
        try:
            number = circuit.CircuitNumber or ""
        except Exception:
            number = ""
        owner = ""
        try:
            owner = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, circuit.Id).Owner or ""
        except Exception:
            owner = ""
        return {
            "circuit_id": _elid_value(circuit.Id),
            "circuit": "{}-{}".format(panel, number),
            "load_name": getattr(circuit, "LoadName", "") or "",
            "circuit_owner": owner,
            "device_owner": "",
            "sync_writeback": False,
        }
