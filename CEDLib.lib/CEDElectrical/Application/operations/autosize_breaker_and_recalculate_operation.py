# -*- coding: utf-8 -*-
"""Apply breaker/frame updates and run calculate operation."""

import Autodesk.Revit.DB.Electrical as DBE
from pyrevit import DB

from CEDElectrical.Application.dto.operation_request import OperationRequest
from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from Snippets import revit_helpers


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


def _calc_options_from_request(request):
    options = {
        'show_output': bool(request.options.get('show_output', False)),
    }
    if request.options.get('calc_preview_enabled') is not None:
        options['calc_preview_enabled'] = bool(request.options.get('calc_preview_enabled', False))
    preview_decision = str(request.options.get('calc_preview_decision') or '').strip().lower()
    if preview_decision:
        options['calc_preview_decision'] = preview_decision
    return options


class AutosizeBreakerAndRecalculateOperation(object):
    """Stages selected breaker/frame values and recalculates circuits."""

    key = 'autosize_breaker_and_recalculate'

    _ALLOWED_TYPES = set(['BRANCH', 'FEEDER', 'XFMR PRI', 'XFMR SEC'])

    def __init__(self, calculate_operation):
        self._calculate_operation = calculate_operation

    def execute(self, request, doc):
        updates = list(request.options.get('updates') or [])
        if not updates:
            return {'status': 'cancelled', 'reason': 'no_updates'}

        by_id = {}
        for row in updates:
            try:
                cid = int(row.get('circuit_id'))
            except Exception:
                continue
            by_id[cid] = row

        circuits = []
        for cid in by_id.keys():
            try:
                el = doc.GetElement(_elid_from_value(cid))
            except Exception:
                el = None
            if isinstance(el, DBE.ElectricalSystem):
                circuits.append(el)
        if not circuits:
            return {'status': 'cancelled', 'reason': 'no_circuits'}

        changed_ids = []
        staged_builtin_values_by_id = {}
        staged_preview_values_by_id = {}
        locked_rows = []

        for circuit in circuits:
            branch_type = self._branch_type(circuit)
            if branch_type not in self._ALLOWED_TYPES:
                continue
            if self._is_locked(doc, circuit.Id):
                locked_rows.append(self._locked_row(circuit, doc))
                continue

            spec = by_id.get(_elid_value(circuit.Id)) or {}
            set_rating = bool(spec.get('set_rating', True))
            set_frame = bool(spec.get('set_frame', True))
            if not (set_rating or set_frame):
                continue

            cid = _elid_value(circuit.Id)
            builtin_values = {}
            preview_values = {}
            did_change = False
            if set_rating:
                rating = self._numeric_or_none(spec.get('rating'))
                if rating is not None:
                    builtin_values['Rating'] = rating
                    preview_values['CKT_Rating_CED'] = rating
                    did_change = self._numeric_changed(circuit, 'Rating', rating) or did_change
            if set_frame:
                frame = self._numeric_or_none(spec.get('frame'))
                if frame is not None:
                    builtin_values['Frame'] = frame
                    preview_values['CKT_Frame_CED'] = frame
                    did_change = self._numeric_changed(circuit, 'Frame', frame) or did_change
            if did_change:
                changed_ids.append(cid)
                staged_builtin_values_by_id[cid] = builtin_values
                staged_preview_values_by_id[cid] = preview_values

        if not changed_ids:
            return {
                'status': 'cancelled',
                'reason': 'no_changes',
                'locked_rows': locked_rows,
                'runtime_alert_rows': [],
            }

        allow_15a = bool(request.options.get('allow_15a', False))
        calc_options = _calc_options_from_request(request)
        if allow_15a:
            calc_options['min_breaker_size_override'] = 15
        calc_options['staged_builtin_values_by_id'] = staged_builtin_values_by_id
        calc_options['staged_preview_values_by_id'] = staged_preview_values_by_id
        calc_options['transaction_name'] = 'Auto Size Breaker/Frame + Calculate Circuits'
        calc_options['write_transaction_name'] = 'Auto Size Breaker/Frame + Write Circuit Parameters'

        calc_request = OperationRequest(
            operation_key='calculate_circuits',
            circuit_ids=changed_ids,
            source=request.source,
            options=calc_options,
        )
        calc_result = self._calculate_operation.execute(calc_request, doc) or {}
        if locked_rows:
            existing = list(calc_result.get('locked_rows') or [])
            calc_result['locked_rows'] = existing + locked_rows
        return calc_result

    def _branch_type(self, circuit):
        try:
            settings = settings_manager.load_circuit_settings(circuit.Document)
            branch = CircuitBranch(circuit, settings=settings)
            return (branch.branch_type or '').upper()
        except Exception:
            return ''

    def _is_locked(self, doc, eid):
        if not getattr(doc, 'IsWorkshared', False):
            return False
        try:
            return DB.WorksharingUtils.GetCheckoutStatus(doc, eid) == DB.CheckoutStatus.OwnedByOtherUser
        except Exception:
            return False

    def _numeric_or_none(self, value):
        try:
            return float(value)
        except Exception:
            return None

    def _numeric_changed(self, circuit, prop_name, value):
        try:
            current = getattr(circuit, prop_name)
            return current is None or abs(float(current) - float(value)) > 0.0001
        except Exception:
            return False

    def _locked_row(self, circuit, doc):
        panel = ''
        try:
            panel = circuit.BaseEquipment.Name if circuit.BaseEquipment else ''
        except Exception:
            panel = ''
        number = ''
        try:
            number = circuit.CircuitNumber or ''
        except Exception:
            number = ''
        owner = ''
        try:
            owner = DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, circuit.Id).Owner or ''
        except Exception:
            owner = ''
        return {
            'circuit_id': _elid_value(circuit.Id),
            'circuit': '{}-{}'.format(panel, number),
            'load_name': getattr(circuit, 'LoadName', '') or '',
            'circuit_owner': owner,
            'device_owner': '',
            'sync_writeback': False,
        }


