# -*- coding: utf-8 -*-

"""Calculate-circuits application operation."""

from datetime import datetime
import time

from pyrevit import DB, forms, script
import Autodesk.Revit.DB.Electrical as DBE

from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.CircuitBranch import (
    CircuitBranch,
    get_native_circuit_type_label,
)
from CEDElectrical.Model.circuit_settings import CircuitSettings
from Snippets import _elecutils as eu
from Snippets import categories as category_utils
from Snippets import revit_helpers


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


SPECIAL_MODE_NORMAL = 'normal'
SPECIAL_MODE_LEGACY = 'legacy_special'
SPECIAL_MODE_REGULAR_COMPATIBLE = 'regular_compatible_special'


def get_special_processing_mode(doc, circuit):
    """Select the one version-gated path for SPARE/SPACE circuits."""
    native_type = get_native_circuit_type_label(circuit)
    if native_type not in ('SPARE', 'SPACE'):
        return SPECIAL_MODE_NORMAL
    revit_version = category_utils._revit_major_version(doc=doc)
    if revit_version and revit_version <= 2025:
        return SPECIAL_MODE_LEGACY
    return SPECIAL_MODE_REGULAR_COMPATIBLE


class CalculateCircuitsOperation(object):
    """Orchestrates calculation, writes, and alert persistence for circuits."""
    key = 'calculate_circuits'
    CIRCUIT_DATA_METADATA_KEYS = ('circuit_vd_method', 'last_calculation')
    PREVIEW_COMPARE_FIELDS = (
        'CKT_Number of Sets_CED',
        'CKT_Wire Hot Size_CEDT',
        'CKT_Wire Ground Size_CEDT',
        'Conduit Size_CEDT',
        'Conduit Type_CEDT',
    )

    def __init__(self, repository, writer, alert_store):
        self.repository = repository
        self.writer = writer
        self.alert_store = alert_store
        self.logger = script.get_logger()

    def execute(self, request, doc):
        """Run calculation workflow for target circuits in the active document."""
        staged_result = request.options.get('staged_calculation')
        if not isinstance(staged_result, dict):
            staged_result = request.options.get('staged_result')
        if isinstance(staged_result, dict):
            return self.apply_staged_result(request, doc, staged_result)

        timings = {}
        started = time.time()
        settings = self._load_effective_settings(doc, request)
        phase_start = time.time()
        circuits = list(
            self.repository.get_target_circuits(doc, request.circuit_ids) or []
        )
        timings['target/circuit collection'] = time.time() - phase_start
        supplied_existing_values_by_id = self._normalize_existing_values_by_id(
            request.options.get('calc_preview_existing_values_by_id')
        )
        force_auto_for_ids = self._normalize_id_set(request.options.get('calc_force_auto_for_ids'))
        staged_preview_values_by_id = self._normalize_values_by_id(
            request.options.get('staged_preview_values_by_id')
        )
        staged_builtin_values_by_id = self._normalize_values_by_id(
            request.options.get('staged_builtin_values_by_id')
        )
        ignored_preview_fields_by_id = self._normalize_preview_ignore_fields_by_id(
            request.options.get('calc_preview_ignore_fields_by_id')
        )

        special_modes_by_id = {}
        legacy_special_ids = set()
        phase_start = time.time()
        for circuit in circuits:
            circuit_id = _elid_value(circuit.Id)
            mode = get_special_processing_mode(doc, circuit)
            special_modes_by_id[circuit_id] = mode
            if mode == SPECIAL_MODE_LEGACY:
                legacy_special_ids.add(circuit_id)
        timings['special-circuit handling'] = time.time() - phase_start

        if len(legacy_special_ids) != len(circuits):
            param_bootstrap = settings_manager.ensure_electrical_parameters_for_calculate(
                doc,
                logger=self.logger,
            )
            status = str((param_bootstrap or {}).get('status') or '').lower()
            if status == 'loaded':
                self.logger.info(
                    'Auto-loaded electrical parameters for calculate. updated={} unchanged={} skipped={}'.format(
                        int((param_bootstrap or {}).get('updated') or 0),
                        int((param_bootstrap or {}).get('unchanged') or 0),
                        int((param_bootstrap or {}).get('skipped') or 0),
                    )
                )
            elif status == 'failed':
                self.logger.warning(
                    'Auto-load electrical parameters before calculate failed: {}'.format(
                        (param_bootstrap or {}).get('reason') or 'unknown'
                    )
                )
        else:
            self.logger.info(
                'Skipping electrical-parameter bootstrap for legacy SPARE/SPACE-only selection.'
            )

        connected_elements_by_id = {}
        phase_start = time.time()
        circuits, locked_ids, locked_rows = self.repository.partition_locked_elements(
            doc,
            circuits,
            settings,
            connected_elements_by_circuit=connected_elements_by_id,
            skip_device_traversal_ids=legacy_special_ids,
        )
        timings['worksharing/lock checks'] = time.time() - phase_start

        phase_start = time.time()
        for circuit in list(circuits or []):
            circuit_id = _elid_value(circuit.Id)
            if circuit_id in legacy_special_ids:
                continue
            if circuit_id in connected_elements_by_id:
                continue
            try:
                if not eu.is_circuit_eligible(circuit):
                    continue
            except Exception:
                continue
            try:
                connected_elements_by_id[circuit_id] = list(circuit.Elements)
            except Exception:
                connected_elements_by_id[circuit_id] = []
        timings['connected-element collection'] = time.time() - phase_start
        if locked_ids:
            summary = self.repository.summarize_locked(doc, locked_ids)
            self.logger.info(
                'Locked elements detected; proceeding with editable set only. circuits={} fixture_devices={} equipment={} other={}'.format(
                    int(summary.get('circuits') or 0),
                    int(summary.get('fixtures') or 0),
                    int(summary.get('equipment') or 0),
                    int(summary.get('other') or 0),
                )
            )

        if not circuits:
            forms.alert('No editable circuits found to process.')
            return {'status': 'cancelled', 'reason': 'no_circuits'}

        count = len(circuits)
        if count > 1000:
            proceed = forms.alert(
                '{} circuits selected.\n\nThis may take a while.\n\n'.format(count),
                title='Large Selection Warning',
                options=['Continue', 'Cancel']
            )
            if proceed != 'Continue':
                return {'status': 'cancelled', 'reason': 'large_selection_cancel'}

        branches = []
        calculation_branches = []
        existing_values_by_id = {}
        current_existing_values_by_id = {}
        existing_alert_payload_by_id = {}
        legacy_special_count = 0
        regular_compatible_special_count = 0
        branch_input_started = time.time()
        engineering_time = 0.0
        for circuit in circuits:
            cid = _elid_value(circuit.Id)
            special_mode = special_modes_by_id.get(cid, SPECIAL_MODE_NORMAL)
            if special_mode == SPECIAL_MODE_LEGACY:
                legacy_special_count += 1
                continue

            circuit_data_payload = self._read_circuit_data_payload(circuit)
            existing_alert_payload_by_id[cid] = circuit_data_payload
            current_existing_values = self._collect_existing_preview_values(circuit)
            current_existing_values_by_id[cid] = current_existing_values
            existing_values_by_id[cid] = dict(
                supplied_existing_values_by_id.get(cid)
                or current_existing_values
            )
            existing_values_by_id[cid]['_is_first_calculation'] = not bool(circuit_data_payload)
            preview_values = dict(staged_preview_values_by_id.get(cid) or {})
            if cid in force_auto_for_ids:
                preview_values['CKT_User Override_CED'] = 0
            branch = CircuitBranch(
                circuit,
                settings=settings,
                preview_values=preview_values,
                connected_elements=connected_elements_by_id.get(cid),
            )
            if cid in force_auto_for_ids:
                branch._calc_preview_force_auto = True
            if not branch.is_power_circuit:
                continue

            branches.append(branch)
            if branch.is_special:
                regular_compatible_special_count += 1
                continue

            calculation_started = time.time()
            branch.calculate_hot_wire_size()
            branch.calculate_neutral_wire_size()
            branch.calculate_ground_wire_size()
            branch.calculate_isolated_ground_wire_size()
            branch.calculate_conduit_size()
            engineering_time += time.time() - calculation_started
            calculation_branches.append(branch)
        timings['branch input construction'] = time.time() - branch_input_started - engineering_time
        timings['engineering calculations'] = engineering_time

        if not branches:
            forms.alert('No editable power circuits found to process.')
            return {'status': 'cancelled', 'reason': 'no_branches'}

        phase_start = time.time()
        preview_rows = self._collect_conduit_wire_preview_rows(
            calculation_branches,
            existing_values_by_id,
            ignored_preview_fields_by_id,
        )
        timings['preview preparation'] = time.time() - phase_start
        preview_changed_ids = self._preview_changed_ids(preview_rows)
        preview_enabled = bool(request.options.get('calc_preview_enabled', False))
        preview_decision = str(request.options.get('calc_preview_decision') or '').strip().lower()

        # Compound operations perform preparatory writes inside an outer
        # TransactionGroup.  Their preview path must roll that group back and
        # rerun the original operation so those writes are applied again.  A
        # staged calculation contains calculation outputs, not the caller's
        # mutation plan, so returning it here would allow the caller to bypass
        # and lose its original action.
        if (
                preview_enabled
                and not preview_decision
                and preview_rows
                and not self._preview_can_use_staged_apply(request)):
            result = {
                'status': 'preview_required',
                'reason': 'conduit_wire_changes',
                'preview_rows': preview_rows,
                'locked_rows': locked_rows,
                'runtime_alert_rows': [],
                'preview_contract': 'rerun_original_operation',
            }
            self._log_timing(
                timings,
                len(branches),
                legacy_special_count,
                regular_compatible_special_count,
                'preview_required',
                started,
            )
            return result

        keep_existing_branches_by_id = {}
        if preview_rows and preview_decision == 'keep_existing':
            rebuilt = self._rebuild_branches_with_existing_sizes(
                calculation_branches,
                existing_values_by_id,
                staged_preview_values_by_id,
                settings,
                changed_ids=preview_changed_ids,
            )
            special_branches = [branch for branch in branches if branch.is_special]
            branches = special_branches + rebuilt
        elif preview_enabled and not preview_decision and preview_rows:
            rebuilt = self._rebuild_branches_with_existing_sizes(
                calculation_branches,
                existing_values_by_id,
                staged_preview_values_by_id,
                settings,
                changed_ids=preview_changed_ids,
            )
            keep_existing_branches_by_id = dict(
                (_elid_value(branch.circuit.Id), branch)
                for branch in rebuilt
                if _elid_value(branch.circuit.Id) in preview_changed_ids
            )

        phase_start = time.time()
        stage_crosses_ui_boundary = bool(
            preview_enabled and not preview_decision and preview_rows
        )
        staged = self._build_staged_result(
            doc,
            branches,
            special_modes_by_id,
            current_existing_values_by_id,
            staged_builtin_values_by_id,
            existing_alert_payload_by_id,
            preview_rows,
            preview_changed_ids,
            locked_rows,
            settings,
            keep_existing_branches_by_id=keep_existing_branches_by_id,
            include_validation_snapshot=stage_crosses_ui_boundary,
        )
        timings['staged result preparation'] = time.time() - phase_start

        if preview_enabled and not preview_decision and preview_rows:
            result = {
                'status': 'preview_required',
                'reason': 'conduit_wire_changes',
                'preview_rows': preview_rows,
                'locked_rows': locked_rows,
                'runtime_alert_rows': [],
                'staged_calculation': staged,
            }
            self._log_timing(
                timings,
                len(branches),
                legacy_special_count,
                regular_compatible_special_count,
                'preview_required',
                started,
            )
            return result

        staged['preview_decision'] = preview_decision
        result = self.apply_staged_result(
            request,
            doc,
            staged,
            settings=settings,
            timings=timings,
            validate=False,
        )
        return result

    def apply_staged_result(
            self,
            request,
            doc,
            staged_result,
            settings=None,
            timings=None,
            validate=True,
    ):
        """Validate and commit a previously calculated plain-data result."""
        if not isinstance(staged_result, dict):
            return {'status': 'cancelled', 'reason': 'invalid_staged_result'}

        timings = dict(timings or {})
        apply_started = time.time()
        settings = settings or self._load_effective_settings(doc, request)
        staged_document = staged_result.get('document')
        if isinstance(staged_document, dict) and not self._document_stamps_match(
                doc,
                staged_document,
        ):
            self._log_timing(
                timings,
                len(list(staged_result.get('circuits') or [])),
                int(staged_result.get('legacy_special_count') or 0),
                int(staged_result.get('regular_compatible_special_count') or 0),
                'stale',
                apply_started,
            )
            return {
                'status': 'stale',
                'reason': 'staged_result_wrong_document',
                'stale_rows': [],
                'locked_rows': list(staged_result.get('locked_rows') or []),
                'calc_preview_rows': list(staged_result.get('preview_rows') or []),
                'calc_preview_decision': str(
                    request.options.get('calc_preview_decision') or ''
                ).strip().lower(),
            }
        expected_settings = staged_result.get('settings_fingerprint')
        if expected_settings:
            try:
                settings_match = str(settings.to_json()) == str(expected_settings)
            except Exception:
                settings_match = False
            if not settings_match:
                self._log_timing(
                    timings,
                    len(list(staged_result.get('circuits') or [])),
                    int(staged_result.get('legacy_special_count') or 0),
                    int(staged_result.get('regular_compatible_special_count') or 0),
                    'stale',
                    apply_started,
                )
                return {
                    'status': 'stale',
                    'reason': 'calculation_settings_changed',
                    'stale_rows': [],
                    'locked_rows': list(staged_result.get('locked_rows') or []),
                    'calc_preview_rows': list(staged_result.get('preview_rows') or []),
                    'calc_preview_decision': str(
                        request.options.get('calc_preview_decision') or ''
                    ).strip().lower(),
                }
        decision = str(
            request.options.get('calc_preview_decision')
            or staged_result.get('preview_decision')
            or ''
        ).strip().lower()
        preview_changed_ids = self._preview_changed_ids(staged_result.get('preview_rows') or [])
        items = list(staged_result.get('circuits') or [])
        approved_items = []
        stale_rows = []
        for item in items:
            if not isinstance(item, dict):
                continue
            circuit_id = self._coerce_boundary_id(item.get('circuit_id'))
            if decision == 'skip' and circuit_id in preview_changed_ids:
                continue
            if validate:
                valid, reason, circuit, connected_elements = self._validate_staged_item(
                    doc,
                    item,
                )
            else:
                valid, reason, circuit, connected_elements = self._rehydrate_staged_item(
                    doc,
                    item,
                )
            if not valid:
                stale_rows.append({
                    'circuit_id': circuit_id,
                    'reason': reason,
                })
                continue
            approved_items.append((item, circuit, connected_elements))

        timings['staged result validation'] = time.time() - apply_started
        if stale_rows:
            self._log_timing(
                timings,
                len(items),
                int(staged_result.get('legacy_special_count') or 0),
                int(staged_result.get('regular_compatible_special_count') or 0),
                'stale',
                apply_started,
            )
            return {
                'status': 'stale',
                'reason': 'staged_result_stale',
                'stale_rows': stale_rows,
                'locked_rows': list(staged_result.get('locked_rows') or []),
                'calc_preview_rows': list(staged_result.get('preview_rows') or []),
                'calc_preview_decision': decision,
            }

        if not approved_items:
            return {
                'status': 'cancelled',
                'reason': 'calc_preview_skipped' if decision == 'skip' else 'no_staged_circuits',
                'locked_rows': list(staged_result.get('locked_rows') or []),
                'runtime_alert_rows': [],
                'calc_preview_rows': list(staged_result.get('preview_rows') or []),
                'calc_preview_decision': decision,
            }

        use_existing_group = bool(request.options.get('use_existing_transaction_group', False))
        tg = None
        transaction_name = str(
            request.options.get('transaction_name') or 'Calculate Circuits'
        ).strip() or 'Calculate Circuits'
        write_transaction_name = str(
            request.options.get('write_transaction_name') or 'Write Shared Parameters'
        ).strip() or 'Write Shared Parameters'
        if not use_existing_group:
            tg = DB.TransactionGroup(doc, transaction_name)
            tg.Start()
        tx = DB.Transaction(doc, write_transaction_name)

        total_fixtures = 0
        total_equipment = 0
        circuit_write_started = time.time()
        device_write_time = 0.0
        alert_write_time = 0.0
        try:
            tx.Start()
            for item, circuit, connected_elements in approved_items:
                self._apply_staged_builtin_values(
                    circuit,
                    self._from_staged_map(item.get('builtin_values') or {}),
                )
                param_values = self._from_staged_map(item.get('parameter_values') or {})
                use_keep_existing = (
                    decision == 'keep_existing'
                    and bool(item.get('preview_changed', False))
                )
                has_keep_existing_variant = False
                if use_keep_existing:
                    staged_keep_values = item.get('keep_existing_parameter_values')
                    if isinstance(staged_keep_values, dict):
                        has_keep_existing_variant = True
                        param_values = self._from_staged_map(staged_keep_values)
                    elif int(staged_result.get('version') or 1) < 2:
                        # Compatibility with an in-memory version-1 stage.
                        param_values = self._apply_keep_existing_decision(
                            param_values,
                            item.get('source_snapshot') or {},
                        )
                self.writer.write_circuit_parameters(circuit, param_values)

                if not bool(item.get('is_special', False)):
                    device_started = time.time()
                    f_cnt, e_cnt = self.writer.write_connected_elements(
                        circuit,
                        param_values,
                        settings,
                        locked_ids=set(),
                        connected_elements=connected_elements,
                    )
                    device_write_time += time.time() - device_started
                    total_fixtures += f_cnt
                    total_equipment += e_cnt

                alert_started = time.time()
                alert_key = (
                    'keep_existing_alert_payload'
                    if use_keep_existing and has_keep_existing_variant
                    else 'alert_payload'
                )
                raw_alert_payload = item.get(alert_key)
                alert_payload = self._from_staged_value(raw_alert_payload)
                if alert_payload is None:
                    self.alert_store.clear_alert_payload(circuit)
                else:
                    self.alert_store.write_alert_payload(circuit, alert_payload)
                alert_write_time += time.time() - alert_started

            sync_started = time.time()
            self._write_locked_sync_payloads(doc, list(staged_result.get('locked_rows') or []))
            alert_write_time += time.time() - sync_started
            tx.Commit()
            if tg is not None:
                tg.Assimilate()
        except Exception as ex:
            try:
                tx.RollBack()
            except Exception:
                pass
            try:
                if tg is not None:
                    tg.RollBack()
            except Exception:
                pass
            self.logger.error('CalculateCircuitsOperation failed: {}'.format(ex))
            raise

        timings['circuit writes'] = time.time() - circuit_write_started - device_write_time - alert_write_time
        timings['device writes'] = device_write_time
        timings['alert processing'] = alert_write_time
        runtime_alert_rows = self._selected_runtime_alert_rows(
            approved_items,
            decision,
        )
        self._log_timing(
            timings,
            len(approved_items),
            int(staged_result.get('legacy_special_count') or 0),
            int(staged_result.get('regular_compatible_special_count') or 0),
            'ok',
            apply_started,
        )

        show_output = bool(request.options.get('show_output', True))
        if show_output:
            self._print_staged_report(
                staged_result,
                total_fixtures,
                total_equipment,
                approved_items=approved_items,
                decision=decision,
            )
        return {
            'status': 'ok',
            'updated_circuits': len([
                item for item, unused_circuit, unused_elements in approved_items
                if not bool(item.get('is_special', False))
            ]),
            'updated_special_circuits': len([
                item for item, unused_circuit, unused_elements in approved_items
                if bool(item.get('is_special', False))
            ]),
            'updated_fixtures': total_fixtures,
            'updated_equipment': total_equipment,
            'locked_rows': list(staged_result.get('locked_rows') or []),
            'runtime_alert_rows': runtime_alert_rows,
            'calc_preview_rows': list(staged_result.get('preview_rows') or []),
            'calc_preview_decision': decision,
        }

    def _coerce_boundary_id(self, value):
        return revit_helpers.coerce_elementid_value(value)

    def _preview_can_use_staged_apply(self, request):
        """Return False when a caller owns preparatory transaction-group edits."""
        return not bool(request.options.get('use_existing_transaction_group', False))

    def _to_staged_value(self, value):
        """Convert a calculation value to plain data for the UI boundary."""
        if isinstance(value, DB.ElementId):
            return {
                '__ced_type__': 'ElementId',
                'value': _elid_value(value),
            }
        if isinstance(value, dict):
            return dict(
                (str(key), self._to_staged_value(item))
                for key, item in value.items()
            )
        if isinstance(value, (list, tuple)):
            return [self._to_staged_value(item) for item in value]
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        try:
            return value.isoformat()
        except Exception:
            return str(value)

    def _from_staged_value(self, value):
        """Rehydrate staged ElementIds only inside the Revit apply context."""
        if isinstance(value, dict):
            if value.get('__ced_type__') == 'ElementId':
                return _elid_from_value(self._coerce_boundary_id(value.get('value')))
            return dict(
                (key, self._from_staged_value(item))
                for key, item in value.items()
            )
        if isinstance(value, list):
            return [self._from_staged_value(item) for item in value]
        return value

    def _from_staged_map(self, values):
        if not isinstance(values, dict):
            return {}
        return dict(
            (key, self._from_staged_value(value))
            for key, value in values.items()
        )

    def _connected_element_ids(self, branch):
        ids = []
        for element in list(getattr(branch, 'connected_elements', []) or []):
            try:
                element_id = _elid_value(element.Id)
            except Exception:
                element_id = 0
            if element_id > 0:
                ids.append(element_id)
        return ids

    def _safe_element_attribute(self, element, name, default=None):
        try:
            return getattr(element, name)
        except Exception:
            return default

    def _current_engineering_inputs(self, circuit):
        values = {}
        for name, descriptor in (
                ('apparent_load', getattr(DBE.ElectricalSystem, 'ApparentLoad', None)),
                ('apparent_current', getattr(DBE.ElectricalSystem, 'ApparentCurrent', None)),
                ('power_factor', getattr(DBE.ElectricalSystem, 'PowerFactor', None)),
                ('poles', getattr(DBE.ElectricalSystem, 'PolesNumber', None))):
            try:
                values[name] = descriptor.__get__(circuit) if descriptor is not None else None
            except Exception:
                values[name] = None
        voltage = None
        try:
            param = circuit.get_Parameter(DB.BuiltInParameter.RBS_ELEC_VOLTAGE)
            if param and param.HasValue:
                voltage = DB.UnitUtils.ConvertFromInternalUnits(
                    param.AsDouble(),
                    DB.UnitTypeId.Volts,
                )
        except Exception:
            voltage = None
        values['voltage'] = voltage
        try:
            values['base_equipment_id'] = _elid_value(circuit.BaseEquipment.Id)
        except Exception:
            values['base_equipment_id'] = 0
        return values

    def _element_version_token(self, element):
        if element is None:
            return ''
        try:
            value = getattr(element, 'VersionGuid', None)
            if value is not None:
                return str(value)
        except Exception:
            pass
        return ''

    def _dependency_version_rows(self, doc, branch):
        """Capture cheap change tokens for every known calculation dependency."""
        rows = []
        seen = set()

        def _add(element):
            if element is None:
                return
            try:
                element_id = _elid_value(element.Id)
            except Exception:
                element_id = 0
            if element_id <= 0 or element_id in seen:
                return
            seen.add(element_id)
            rows.append({
                'element_id': element_id,
                'version_guid': self._element_version_token(element),
            })

        def _add_type_dependencies(element):
            if element is None:
                return
            try:
                type_id = element.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    _add(doc.GetElement(type_id))
            except Exception:
                pass
            try:
                symbol = element.Symbol
                _add(symbol)
                _add(symbol.Family if symbol is not None else None)
            except Exception:
                pass

        circuit = branch.circuit
        _add(circuit)
        try:
            base_equipment = circuit.BaseEquipment
        except Exception:
            base_equipment = None
        _add(base_equipment)
        _add_type_dependencies(base_equipment)

        for element in list(getattr(branch, 'connected_elements', []) or []):
            _add(element)
            _add_type_dependencies(element)
            try:
                ds_param = element.get_Parameter(
                    DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM
                )
                if ds_param and ds_param.HasValue:
                    _add(doc.GetElement(ds_param.AsElementId()))
            except Exception:
                pass
        return rows

    def _staged_source_snapshot(self, doc, circuit, branch, current_preview_values, connected_ids):
        return self._to_staged_value(
            {
                'circuit_type': getattr(branch, '_native_circuit_type_label', ''),
                'circuit_number': self._safe_element_attribute(circuit, 'CircuitNumber', ''),
                'load_name': self._safe_element_attribute(circuit, 'LoadName', ''),
                'length': self._safe_element_attribute(circuit, 'Length', None),
                'rating': self._safe_element_attribute(circuit, 'Rating', None),
                'frame': self._safe_element_attribute(circuit, 'Frame', None),
                'apparent_load': getattr(branch, 'apparent_power', None),
                'apparent_current': getattr(branch, 'apparent_current', None),
                'power_factor': getattr(branch, 'power_factor', None),
                'poles': getattr(branch, 'poles', None),
                'voltage': getattr(branch, 'voltage', None),
                'base_equipment_id': _elid_value(
                    getattr(getattr(circuit, 'BaseEquipment', None), 'Id', None)
                ),
                'preview_values': dict(current_preview_values or {}),
                'connected_element_ids': list(connected_ids or []),
                'dependency_versions': self._dependency_version_rows(doc, branch),
            }
        )

    def _notice_rows(self, branch):
        rows = []
        notices = getattr(branch, 'notices', None)
        if not notices:
            return rows
        for definition, severity, group, message in list(notices.items or []):
            definition_id = ''
            persistent = True
            try:
                definition_id = definition.GetId() if definition else ''
                persistent = bool(getattr(definition, 'persistent', True))
            except Exception:
                pass
            rows.append(
                {
                    'definition_id': definition_id or '',
                    'severity': severity or '',
                    'group': group or 'Other',
                    'message': message or '',
                    'persistent': persistent,
                }
            )
        return rows

    def _runtime_alert_rows_from_branch(self, branch):
        rows = []
        for notice in self._notice_rows(branch):
            if notice.get('persistent', True):
                continue
            rows.append(
                {
                    'panel': branch.panel or '',
                    'number': branch.circuit_number or '',
                    'load_name': branch.load_name or '',
                    'group': notice.get('group') or 'Other',
                    'definition_id': notice.get('definition_id') or '-',
                    'message': notice.get('message') or '',
                }
            )
        return rows

    def _build_staged_result(
            self,
            doc,
            branches,
            special_modes_by_id,
            current_existing_values_by_id,
            staged_builtin_values_by_id,
            existing_alert_payload_by_id,
            preview_rows,
            preview_changed_ids,
            locked_rows,
            settings,
            keep_existing_branches_by_id=None,
            include_validation_snapshot=False,
    ):
        staged_circuits = []
        report_rows = []
        runtime_alert_rows = []
        preview_changed_id_set = set(preview_changed_ids or [])
        keep_existing_by_id = dict(keep_existing_branches_by_id or {})
        for branch in list(branches or []):
            circuit_id = _elid_value(branch.circuit.Id)
            connected_ids = []
            if not branch.is_special:
                connected_ids = self._connected_element_ids(branch)
            parameter_values = self._collect_shared_param_values(branch)
            alert_payload = self._build_alert_payload(
                branch,
                existing_payload=existing_alert_payload_by_id.get(circuit_id),
            )
            keep_branch = keep_existing_by_id.get(circuit_id)
            keep_parameter_values = None
            keep_alert_payload = None
            keep_notice_rows = None
            keep_runtime_rows = None
            if keep_branch is not None:
                keep_parameter_values = self._collect_shared_param_values(keep_branch)
                keep_alert_payload = self._build_alert_payload(
                    keep_branch,
                    existing_payload=existing_alert_payload_by_id.get(circuit_id),
                )
                keep_notice_rows = self._notice_rows(keep_branch)
                keep_runtime_rows = self._runtime_alert_rows_from_branch(keep_branch)
            current_preview_values = current_existing_values_by_id.get(circuit_id) or {}
            item = {
                'circuit_id': circuit_id,
                'circuit_name': branch.name or '',
                'is_special': bool(branch.is_special),
                'special_mode': special_modes_by_id.get(
                    circuit_id,
                    SPECIAL_MODE_NORMAL,
                ),
                'parameter_values': self._to_staged_value(parameter_values),
                'proposed_values': self._to_staged_value(parameter_values),
                'existing_values': self._to_staged_value(current_preview_values),
                'builtin_values': self._to_staged_value(
                    staged_builtin_values_by_id.get(circuit_id) or {}
                ),
                'connected_element_ids': list(connected_ids),
                'device_ids': list(connected_ids),
                'source_snapshot': (
                    self._staged_source_snapshot(
                        doc,
                        branch.circuit,
                        branch,
                        current_preview_values,
                        connected_ids,
                    )
                    if include_validation_snapshot
                    else {}
                ),
                'preview_changed': circuit_id in preview_changed_id_set,
                'alert_payload': self._to_staged_value(alert_payload),
                'keep_existing_parameter_values': self._to_staged_value(
                    keep_parameter_values
                ) if keep_parameter_values is not None else None,
                'keep_existing_alert_payload': self._to_staged_value(
                    keep_alert_payload
                ) if keep_branch is not None else None,
                'notice_rows': self._to_staged_value(self._notice_rows(branch)),
                'keep_existing_notice_rows': self._to_staged_value(
                    keep_notice_rows
                ) if keep_notice_rows is not None else None,
                'runtime_alert_rows': self._to_staged_value(
                    self._runtime_alert_rows_from_branch(branch)
                ),
                'keep_existing_runtime_alert_rows': self._to_staged_value(
                    keep_runtime_rows
                ) if keep_runtime_rows is not None else None,
            }
            staged_circuits.append(item)
            report_rows.append(
                {
                    'circuit_id': circuit_id,
                    'circuit': branch.name or '',
                    'notices': self._notice_rows(branch),
                }
            )
            runtime_alert_rows.extend(self._runtime_alert_rows_from_branch(branch))

        return {
            'version': 2,
            'document': self._document_stamp(doc),
            'settings_fingerprint': self._settings_fingerprint(settings),
            'circuit_ids': [
                item.get('circuit_id')
                for item in staged_circuits
            ],
            'circuits': staged_circuits,
            'preview_rows': self._to_staged_value(list(preview_rows or [])),
            'locked_rows': self._to_staged_value(list(locked_rows or [])),
            'report_rows': self._to_staged_value(report_rows),
            'runtime_alert_rows': self._to_staged_value(runtime_alert_rows),
            'legacy_special_count': len([
                value for value in special_modes_by_id.values()
                if value == SPECIAL_MODE_LEGACY
            ]),
            'regular_compatible_special_count': len([
                item for item in staged_circuits
                if item.get('special_mode') == SPECIAL_MODE_REGULAR_COMPATIBLE
            ]),
            'settings': {
                'write_fixture_results': bool(
                    getattr(settings, 'write_fixture_results', False)
                ),
                'write_equipment_results': bool(
                    getattr(settings, 'write_equipment_results', False)
                ),
            },
        }

    def _settings_fingerprint(self, settings):
        try:
            return str(settings.to_json())
        except Exception:
            return ''

    def _load_effective_settings(self, doc, request):
        settings = settings_manager.load_circuit_settings(doc)
        min_breaker_size_override = request.options.get('min_breaker_size_override')
        if min_breaker_size_override is not None:
            try:
                override_value = int(min_breaker_size_override)
                if override_value > 0:
                    settings = CircuitSettings.from_json(settings.to_json())
                    settings.set('min_breaker_size', override_value)
            except Exception:
                pass
        return settings

    def _document_stamp(self, doc):
        stamp = {
            'path': str(getattr(doc, 'PathName', '') or ''),
            'title': str(getattr(doc, 'Title', '') or ''),
            'version': str(
                getattr(getattr(doc, 'Application', None), 'VersionNumber', '')
                or ''
            ),
        }
        try:
            stamp['hash'] = int(doc.GetHashCode())
        except Exception:
            stamp['hash'] = None
        return stamp

    def _document_stamps_match(self, doc, staged_document):
        current = self._document_stamp(doc)
        for key in ('path', 'title', 'version'):
            expected = str(staged_document.get(key) or '')
            if expected and str(current.get(key) or '') != expected:
                return False
        expected_hash = staged_document.get('hash')
        current_hash = current.get('hash')
        if expected_hash is not None and current_hash is not None:
            try:
                if int(expected_hash) != int(current_hash):
                    return False
            except Exception:
                return False
        return True

    def _staged_values_match(self, current, expected):
        if current is None or expected is None:
            return current is expected or (
                current in ('', '-') and expected in (None, '', '-')
            )
        if isinstance(current, (int, float)) and isinstance(expected, (int, float)):
            difference = abs(float(current) - float(expected))
            scale = max(1.0, abs(float(current)), abs(float(expected)))
            return difference <= max(1e-9, 1e-9 * scale)
        if isinstance(current, str) or isinstance(expected, str):
            return self._normalize_preview_compare_value(current) == self._normalize_preview_compare_value(expected)
        return current == expected

    def _owned_by_other_user(self, doc, element_id):
        if not getattr(doc, 'IsWorkshared', False):
            return False
        try:
            return DB.WorksharingUtils.GetCheckoutStatus(
                doc,
                element_id,
            ) == DB.CheckoutStatus.OwnedByOtherUser
        except Exception:
            return False

    def _validate_dependency_versions(self, doc, snapshot):
        for dependency in list(snapshot.get('dependency_versions') or []):
            if not isinstance(dependency, dict):
                continue
            dependency_id = self._coerce_boundary_id(dependency.get('element_id'))
            if dependency_id <= 0:
                continue
            try:
                current_dependency = doc.GetElement(_elid_from_value(dependency_id))
            except Exception:
                current_dependency = None
            if current_dependency is None:
                return False, 'calculation_dependency_deleted'
            expected_version = str(dependency.get('version_guid') or '')
            current_version = self._element_version_token(current_dependency)
            if expected_version and current_version != expected_version:
                return False, 'calculation_dependency_changed'
        return True, ''

    def _rehydrate_staged_item(self, doc, item):
        """Rehydrate an in-process stage without rescanning circuit.Elements."""
        circuit_id = self._coerce_boundary_id(item.get('circuit_id'))
        if circuit_id <= 0:
            return False, 'invalid_circuit_id', None, []
        try:
            circuit = doc.GetElement(_elid_from_value(circuit_id))
        except Exception:
            circuit = None
        if circuit is None:
            return False, 'circuit_deleted', None, []
        connected_elements = []
        if not bool(item.get('is_special', False)):
            for raw_id in list(item.get('connected_element_ids') or []):
                element_id = self._coerce_boundary_id(raw_id)
                if element_id <= 0:
                    continue
                try:
                    element = doc.GetElement(_elid_from_value(element_id))
                except Exception:
                    element = None
                if element is None:
                    return False, 'connected_element_deleted', None, []
                connected_elements.append(element)
        return True, '', circuit, connected_elements

    def _validate_staged_item(self, doc, item):
        circuit_id = self._coerce_boundary_id(item.get('circuit_id'))
        if circuit_id <= 0:
            return False, 'invalid_circuit_id', None, []
        circuit_element_id = _elid_from_value(circuit_id)
        try:
            circuit = doc.GetElement(circuit_element_id)
        except Exception:
            circuit = None
        if circuit is None:
            return False, 'circuit_deleted', None, []
        try:
            if not eu.is_circuit_eligible(circuit):
                return False, 'circuit_no_longer_eligible', None, []
        except Exception:
            return False, 'circuit_no_longer_eligible', None, []
        if self._owned_by_other_user(doc, circuit.Id):
            return False, 'circuit_owned_by_other_user', None, []

        expected_mode = str(item.get('special_mode') or SPECIAL_MODE_NORMAL)
        if get_special_processing_mode(doc, circuit) != expected_mode:
            return False, 'circuit_type_or_revit_version_changed', None, []

        snapshot = self._from_staged_map(item.get('source_snapshot') or {})
        expected_type = str(snapshot.get('circuit_type') or '')
        current_type = get_native_circuit_type_label(circuit)
        if expected_type != current_type:
            return False, 'circuit_type_changed', None, []
        for attr_name, snapshot_name in (
                ('CircuitNumber', 'circuit_number'),
                ('LoadName', 'load_name')):
            current_value = self._safe_element_attribute(circuit, attr_name, '')
            if not self._staged_values_match(
                    current_value,
                    snapshot.get(snapshot_name, ''),
            ):
                return False, '{}_changed'.format(snapshot_name), None, []

        for attr_name, snapshot_name in (
                ('Length', 'length'),
                ('Rating', 'rating'),
                ('Frame', 'frame')):
            current_value = self._safe_element_attribute(circuit, attr_name, None)
            if not self._staged_values_match(
                    current_value,
                    snapshot.get(snapshot_name),
            ):
                return False, '{}_changed'.format(snapshot_name), None, []

        expected_preview = snapshot.get('preview_values') or {}
        current_preview = self._collect_existing_preview_values(circuit)
        for key, expected in expected_preview.items():
            if key == '_is_first_calculation':
                continue
            if not self._staged_values_match(current_preview.get(key), expected):
                return False, 'calculation_input_changed', None, []

        current_engineering = self._current_engineering_inputs(circuit)
        for input_name in (
                'apparent_load',
                'apparent_current',
                'power_factor',
                'poles',
                'voltage',
                'base_equipment_id'):
            if not self._staged_values_match(
                    current_engineering.get(input_name),
                    snapshot.get(input_name),
            ):
                return False, '{}_changed'.format(input_name), None, []

        connected_ids = [
            self._coerce_boundary_id(value)
            for value in list(item.get('connected_element_ids') or [])
        ]
        connected_ids = [value for value in connected_ids if value > 0]
        connected_elements = []
        if not bool(item.get('is_special', False)):
            try:
                connected_elements = list(circuit.Elements)
            except Exception:
                connected_elements = []
            current_ids = sorted([
                _elid_value(element.Id)
                for element in connected_elements
                if getattr(element, 'Id', None) is not None
            ])
            if current_ids != sorted(connected_ids):
                return False, 'connected_elements_changed', None, []
            rehydrated_elements = []
            for element_id in connected_ids:
                try:
                    element = doc.GetElement(_elid_from_value(element_id))
                except Exception:
                    element = None
                if element is None:
                    return False, 'connected_element_deleted', None, []
                rehydrated_elements.append(element)
            connected_elements = rehydrated_elements
            for element in connected_elements:
                if self._owned_by_other_user(doc, element.Id):
                    return False, 'connected_element_owned_by_other_user', None, []

        dependencies_valid, dependency_reason = self._validate_dependency_versions(
            doc,
            snapshot,
        )
        if not dependencies_valid:
            return False, dependency_reason, None, []

        return True, '', circuit, connected_elements

    def _apply_keep_existing_decision(self, parameter_values, source_snapshot):
        values = dict(parameter_values or {})
        existing_values = dict(source_snapshot.get('preview_values') or {})
        for field_name in self.PREVIEW_COMPARE_FIELDS:
            if field_name in existing_values:
                values[field_name] = existing_values.get(field_name)
        values['CKT_User Override_CED'] = 1
        return values

    def _selected_runtime_alert_rows(self, approved_items, decision):
        rows = []
        for item, unused_circuit, unused_elements in list(approved_items or []):
            use_keep = (
                decision == 'keep_existing'
                and bool(item.get('preview_changed', False))
            )
            raw_rows = item.get(
                'keep_existing_runtime_alert_rows' if use_keep else 'runtime_alert_rows'
            )
            if raw_rows is None:
                raw_rows = item.get('runtime_alert_rows')
            rows.extend(list(self._from_staged_value(raw_rows) or []))
        return rows

    def _log_timing(
            self,
            timings,
            circuit_count,
            legacy_special_count,
            regular_compatible_special_count,
            status,
            started,
    ):
        ordered_names = (
            'target/circuit collection',
            'special-circuit handling',
            'connected-element collection',
            'worksharing/lock checks',
            'branch input construction',
            'engineering calculations',
            'preview preparation',
            'staged result preparation',
            'staged result validation',
            'circuit writes',
            'device writes',
            'alert processing',
        )
        parts = []
        for name in ordered_names:
            if name in timings:
                parts.append(
                    '{}={:.3f}s'.format(name, float(timings.get(name) or 0.0))
                )
        self.logger.info(
            'Calculate Circuits timing status={} circuits={} legacy_special={} '
            'regular_compatible_special={} total={:.3f}s {}'.format(
                status,
                circuit_count,
                legacy_special_count,
                regular_compatible_special_count,
                time.time() - started,
                ', '.join(parts),
            )
        )

    def _print_staged_report(
            self,
            staged_result,
            total_fixtures,
            total_equipment,
            approved_items=None,
            decision='',
    ):
        output = script.get_output()
        try:
            output.show()
        except Exception:
            pass
        output.close_others()
        output.print_md('## Shared Parameters Updated')
        if approved_items is None:
            applied_ids = set(
                item.get('circuit_id')
                for item in list(staged_result.get('circuits') or [])
            )
        else:
            applied_ids = set(
                item.get('circuit_id')
                for item, unused_circuit, unused_elements in approved_items
            )
        applied_items = [
            item for item in list(staged_result.get('circuits') or [])
            if item.get('circuit_id') in applied_ids
        ]
        regular_count = len([
            item for item in applied_items
            if not bool(item.get('is_special', False))
        ])
        special_count = len([
            item for item in applied_items
            if bool(item.get('is_special', False))
        ])
        output.print_md('* Circuits updated: **{}**'.format(regular_count))
        if special_count:
            output.print_md('* SPARE/SPACE circuits refreshed: **{}**'.format(special_count))
        output.print_md('* Fixtures and Devices updated: **{}**'.format(total_fixtures))
        output.print_md('* Electrical Equipment updated: **{}**'.format(total_equipment))
        locked_rows = list(staged_result.get('locked_rows') or [])
        if locked_rows:
            output.print_md('\n## Skipped Elements')
            output.print_md('The following elements are owned by other users and could not be calculated.')
            table = []
            for row in locked_rows:
                table.append(
                    [
                        row.get('circuit', ''),
                        row.get('circuit_owner', '') or '-',
                        row.get('device_owner', '') or '-',
                    ]
                )
            output.print_table(
                table_data=table,
                columns=['Circuit', 'Circuit Owner', 'Device Owner'],
            )
        notice_lines = []
        selected_decision = str(
            decision or staged_result.get('preview_decision') or ''
        ).strip().lower()
        for item in list(staged_result.get('circuits') or []):
            if item.get('circuit_id') not in applied_ids:
                continue
            use_keep = (
                selected_decision == 'keep_existing'
                and bool(item.get('preview_changed', False))
            )
            raw_notices = item.get(
                'keep_existing_notice_rows' if use_keep else 'notice_rows'
            )
            if raw_notices is None:
                raw_notices = item.get('notice_rows')
            for notice in list(self._from_staged_value(raw_notices) or []):
                notice_lines.append(
                    '* **{}** {}: {}'.format(
                        notice.get('group') or 'Other',
                        item.get('circuit_name') or '-',
                        notice.get('message') or '',
                    )
                )
        if notice_lines:
            output.print_md('\n## Warnings / Errors')
            for line in notice_lines:
                output.print_md(line)

    def _lookup_param_text(self, element, param_name, default_value=''):
        try:
            param = element.LookupParameter(param_name)
        except Exception:
            param = None
        if not param:
            return default_value
        try:
            value = param.AsString()
            if value is None:
                value = param.AsValueString()
            return str(value or default_value)
        except Exception:
            return default_value

    def _lookup_param_value(self, element, param_name, default_value=None):
        try:
            param = element.LookupParameter(param_name)
        except Exception:
            param = None
        if not param:
            return default_value
        try:
            st = param.StorageType
            if st == DB.StorageType.Integer:
                return param.AsInteger()
            if st == DB.StorageType.Double:
                return param.AsDouble()
            if st == DB.StorageType.String:
                value = param.AsString()
                if value is None:
                    value = param.AsValueString()
                return value
        except Exception:
            return default_value
        return default_value

    def _read_circuit_data_payload(self, circuit):
        try:
            payload = self.alert_store.read_alert_payload(circuit)
        except Exception:
            payload = None
        return payload if isinstance(payload, dict) else {}

    def _collect_existing_preview_values(self, circuit):
        return {
            'current_summary': self._lookup_param_text(circuit, 'Conduit and Wire Size_CEDT', '-').strip() or '-',
            'CKT_User Override_CED': self._lookup_param_value(
                circuit,
                'CKT_User Override_CED',
                0,
            ),
            'CKT_Number of Sets_CED': self._lookup_param_value(circuit, 'CKT_Number of Sets_CED', 0),
            'CKT_Include Neutral_CED': self._lookup_param_value(circuit, 'CKT_Include Neutral_CED', 0),
            'CKT_Include Isolated Ground_CED': self._lookup_param_value(circuit, 'CKT_Include Isolated Ground_CED', 0),
            'CKT_Wire Hot Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Hot Size_CEDT', ''),
            'CKT_Wire Neutral Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Neutral Size_CEDT', ''),
            'CKT_Wire Ground Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Ground Size_CEDT', ''),
            'CKT_Wire Isolated Ground Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Isolated Ground Size_CEDT', ''),
            'CKT_Length Makeup_CED': self._lookup_param_value(
                circuit,
                'CKT_Length Makeup_CED',
                0.0,
            ),
            'Wire Material_CEDT': self._lookup_param_text(circuit, 'Wire Material_CEDT', ''),
            'Wire Temparature Rating_CEDT': self._lookup_param_text(circuit, 'Wire Temparature Rating_CEDT', ''),
            'Wire Insulation_CEDT': self._lookup_param_text(circuit, 'Wire Insulation_CEDT', ''),
            'Conduit Size_CEDT': self._lookup_param_text(circuit, 'Conduit Size_CEDT', ''),
            'Conduit Type_CEDT': self._lookup_param_text(circuit, 'Conduit Type_CEDT', ''),
        }

    def _normalize_id_set(self, raw):
        ids = set()
        if isinstance(raw, dict):
            iterable = raw.keys()
        else:
            iterable = list(raw or [])
        for item in iterable:
            value = self._coerce_boundary_id(item)
            if value > 0:
                ids.add(value)
        return ids

    def _normalize_existing_values_by_id(self, raw):
        normalized = {}
        if not isinstance(raw, dict):
            return normalized
        for key, value in list(raw.items()):
            cid = self._coerce_boundary_id(key)
            if cid <= 0 or not isinstance(value, dict):
                continue
            normalized[cid] = dict(value)
        return normalized

    def _normalize_values_by_id(self, raw):
        normalized = {}
        if not isinstance(raw, dict):
            return normalized
        for key, value in list(raw.items()):
            cid = self._coerce_boundary_id(key)
            if cid <= 0 or not isinstance(value, dict):
                continue
            normalized[cid] = dict(value)
        return normalized

    def _normalize_summary_text(self, value):
        text = str(value or '').strip().lower()
        return ' '.join(text.split()) or '-'

    def _normalize_preview_compare_value(self, value):
        if value is None:
            return '-'
        text = str(value or '').strip()
        if not text:
            return '-'
        return ' '.join(text.lower().split())

    def _collect_preview_compare_values(self, branch):
        return {
            'CKT_Number of Sets_CED': branch.number_of_sets,
            'CKT_Wire Hot Size_CEDT': branch.hot_wire_size,
            'CKT_Wire Ground Size_CEDT': branch.ground_wire_size,
            'Conduit Size_CEDT': branch.conduit_size,
            'Conduit Type_CEDT': branch.conduit_type,
        }

    def _normalize_preview_ignore_fields_by_id(self, raw):
        normalized = {}
        if not isinstance(raw, dict):
            return normalized
        allowed = set(self.PREVIEW_COMPARE_FIELDS)
        for key, value in list(raw.items()):
            cid = self._coerce_boundary_id(key)
            if cid <= 0:
                continue
            if value == '*' or value is True:
                normalized[cid] = set(allowed)
                continue
            fields = set()
            for field_name in list(value or []):
                field_text = str(field_name or '').strip()
                if field_text in allowed:
                    fields.add(field_text)
            if fields:
                normalized[cid] = fields
        return normalized

    def _has_preview_compare_change(self, previous_values, new_values, ignored_fields=None):
        previous = dict(previous_values or {})
        new = dict(new_values or {})
        ignored = set(ignored_fields or [])
        for field_name in self.PREVIEW_COMPARE_FIELDS:
            if field_name in ignored:
                continue
            if (
                self._normalize_preview_compare_value(previous.get(field_name))
                != self._normalize_preview_compare_value(new.get(field_name))
            ):
                return True
        return False

    def _collect_conduit_wire_preview_rows(self, branches, existing_values_by_id, ignored_preview_fields_by_id=None):
        rows = []
        ignored_by_id = dict(ignored_preview_fields_by_id or {})
        for branch in list(branches or []):
            cid = _elid_value(branch.circuit.Id)
            existing = dict(existing_values_by_id.get(cid) or {})
            if bool(existing.get('_is_first_calculation', False)):
                continue
            previous = str(existing.get('current_summary') or '-').strip() or '-'
            new_value = str(branch.get_conduit_and_wire_size() or '-').strip() or '-'
            compare_values = self._collect_preview_compare_values(branch)
            if not self._has_preview_compare_change(existing, compare_values, ignored_by_id.get(cid)):
                continue
            rows.append({
                'circuit_id': cid,
                'panel': branch.panel or '',
                'number': branch.circuit_number or '',
                'circuit': '{} / {}'.format(branch.panel or '-', branch.circuit_number or '-'),
                'load_name': branch.load_name or '',
                'previous_size': previous,
                'new_size': new_value,
                'compare_values': compare_values,
            })
        return rows

    def _preview_changed_ids(self, preview_rows):
        changed_ids = set()
        for row in list(preview_rows or []):
            cid = self._coerce_boundary_id(row.get('circuit_id'))
            if cid > 0:
                changed_ids.add(cid)
        return changed_ids

    def _rebuild_branches_with_existing_sizes(
            self,
            branches,
            existing_values_by_id,
            staged_preview_values_by_id,
            settings,
            changed_ids=None,
    ):
        changed_ids = set(changed_ids or self._preview_changed_ids(
            self._collect_conduit_wire_preview_rows(branches, existing_values_by_id)
        ))
        if not changed_ids:
            return branches

        rebuilt = []
        for branch in list(branches or []):
            cid = _elid_value(branch.circuit.Id)
            if cid not in changed_ids:
                rebuilt.append(branch)
                continue
            preview_values = dict(staged_preview_values_by_id.get(cid) or {})
            existing_values = dict(existing_values_by_id.get(cid) or {})
            keep_values = {'CKT_User Override_CED': 1}
            for field_name in self.PREVIEW_COMPARE_FIELDS:
                if field_name in existing_values:
                    keep_values[field_name] = existing_values.get(field_name)
            preview_values.update(keep_values)
            keep_branch = CircuitBranch(
                branch.circuit,
                settings=settings,
                preview_values=preview_values,
                connected_elements=getattr(branch, 'connected_elements', None),
            )
            keep_branch.calculate_hot_wire_size()
            keep_branch.calculate_neutral_wire_size()
            keep_branch.calculate_ground_wire_size()
            keep_branch.calculate_isolated_ground_wire_size()
            keep_branch.calculate_conduit_size()
            keep_branch._calc_preview_keep_existing = True
            rebuilt.append(keep_branch)
        return rebuilt

    def _apply_staged_builtin_values(self, circuit, values):
        if not isinstance(values, dict) or not values:
            return
        if 'Rating' in values:
            self._set_numeric_builtin(
                circuit,
                'Rating',
                getattr(DB.BuiltInParameter, 'RBS_ELEC_CIRCUIT_RATING_PARAM', None),
                values.get('Rating'),
            )
        if 'Frame' in values:
            self._set_numeric_builtin(
                circuit,
                'Frame',
                getattr(DB.BuiltInParameter, 'RBS_ELEC_CIRCUIT_FRAME_PARAM', None),
                values.get('Frame'),
            )

    def _set_numeric_builtin(self, circuit, prop_name, bip, value):
        try:
            numeric = float(value)
        except Exception:
            return False

        try:
            current = getattr(circuit, prop_name)
            if current is not None and abs(float(current) - numeric) <= 0.0001:
                return False
            if current is None or abs(float(current) - numeric) > 0.0001:
                setattr(circuit, prop_name, numeric)
                # Rating and Frame properties are backed by the same Revit
                # parameters used below.  A successful property assignment is
                # the write; do not immediately write the backing parameter a
                # second time.
                return True
        except Exception:
            pass

        param = None
        if bip is not None:
            try:
                param = circuit.get_Parameter(bip)
            except Exception:
                param = None
        if not param:
            return False
        if param.StorageType == DB.StorageType.Integer:
            numeric = int(round(numeric))
        return revit_helpers.set_parameter_if_changed(param, numeric)

    def _collect_shared_param_values(self, branch):
        """Map branch results into shared-parameter values."""
        if branch.is_special:
            return branch.get_special_parameter_reset_values(settings_manager.RESULT_PARAM_NAMES)

        neutral_qty = branch.neutral_wire_quantity or 0
        ig_qty = branch.isolated_ground_wire_quantity or 0
        include_neutral = 1 if neutral_qty > 0 else 0
        include_ig = 1 if ig_qty > 0 else 0

        values = {
            'CKT_Circuit Type_CEDT': branch.branch_type,
            'CKT_Panel_CEDT': branch.panel,
            'CKT_Circuit Number_CEDT': branch.circuit_number,
            'CKT_Load Name_CEDT': branch.load_name,
            'CKT_Rating_CED': branch.rating,
            'CKT_Frame_CED': branch.frame,
            'CKT_Length_CED': branch.length,
            'CKT_Schedule Notes_CEDT': branch.circuit_notes,
            'Voltage Drop Percentage_CED': branch.voltage_drop_percentage,
            'CKT_Wire Hot Size_CEDT': branch.hot_wire_size,
            'CKT_Number of Wires_CED': branch.number_of_wires,
            'CKT_Number of Sets_CED': branch.number_of_sets,
            'CKT_Wire Hot Quantity_CED': branch.hot_wire_quantity,
            'CKT_Wire Ground Size_CEDT': branch.ground_wire_size,
            'CKT_Wire Ground Quantity_CED': branch.ground_wire_quantity,
            'CKT_Wire Neutral Size_CEDT': branch.neutral_wire_size,
            'CKT_Wire Neutral Quantity_CED': neutral_qty,
            'CKT_Wire Isolated Ground Size_CEDT': branch.isolated_ground_wire_size,
            'CKT_Wire Isolated Ground Quantity_CED': ig_qty,
            'CKT_Include Neutral_CED': include_neutral,
            'CKT_Include Isolated Ground_CED': include_ig,
            'Wire Material_CEDT': branch.wire_material,
            'Wire Temparature Rating_CEDT': branch.wire_temp_rating,
            'Wire Insulation_CEDT': branch.wire_insulation,
            'Conduit Size_CEDT': branch.conduit_size,
            'Conduit Type_CEDT': branch.conduit_type,
            'Conduit Fill Percentage_CED': branch.conduit_fill_percentage,
            'Wire Size_CEDT': branch.get_wire_size_callout(),
            'Conduit and Wire Size_CEDT': branch.get_conduit_and_wire_size(),
            'Circuit Load Current_CED': branch.circuit_load_current,
            'Circuit Ampacity_CED': branch.circuit_base_ampacity,
            'CKT_Length Makeup_CED': branch.wire_length_makeup,
        }
        if bool(getattr(branch, '_calc_preview_keep_existing', False)):
            values['CKT_User Override_CED'] = 1
        elif bool(getattr(branch, '_calc_preview_force_auto', False)):
            values['CKT_User Override_CED'] = 0
        return values

    def _build_alert_payload(self, branch, existing_payload=None):
        """Build serializable alert payload for persistence."""
        notices = getattr(branch, 'notices', None)
        if existing_payload is None:
            existing = self.alert_store.read_alert_payload(branch.circuit) or {}
        else:
            existing = existing_payload or {}
        if not isinstance(existing, dict):
            existing = {}
        metadata = {}
        metadata_keys = self.CIRCUIT_DATA_METADATA_KEYS
        if branch.is_special:
            # Voltage-drop method is not meaningful for a native SPARE/SPACE
            # circuit.  Rebuild the payload without carrying that old
            # regular-circuit metadata forward.
            metadata_keys = ('last_calculation',)
        for key in metadata_keys:
            if key in existing:
                metadata[key] = existing.get(key)
        metadata['last_calculation'] = datetime.utcnow().isoformat() + 'Z'

        has_notices = bool(notices and notices.has_items())
        if not has_notices and not metadata:
            return None

        existing_hidden = existing.get('hidden_definition_ids') if isinstance(existing, dict) else []
        if not isinstance(existing_hidden, list):
            existing_hidden = []

        items = []
        present_ids = set()
        if has_notices:
            for definition, severity, group, message in notices.items:
                if definition is None:
                    continue
                if not getattr(definition, 'persistent', True):
                    continue
                definition_id = definition.GetId() if definition else None
                if definition_id:
                    present_ids.add(definition_id)
                items.append({
                    'definition_id': definition_id,
                    'severity': severity,
                    'group': group,
                    'message': message,
                })

        hidden_ids = sorted(list(set(existing_hidden).intersection(present_ids)))
        payload = {
            'version': 1,
            'generated_utc': datetime.utcnow().isoformat() + 'Z',
            'circuit': {
                'id': _elid_value(branch.circuit.Id),
                'name': branch.name,
                'panel': branch.panel,
                'number': branch.circuit_number,
            },
            'alerts': items,
            'hidden_definition_ids': hidden_ids,
        }
        payload.update(metadata)
        return payload

    def _print_report(self, branches, total_fixtures, total_equipment, locked_rows):
        """Print a post-run report to pyRevit output."""
        output = script.get_output()
        try:
            output.show()
        except Exception:
            pass
        output.close_others()
        output.print_md('## Shared Parameters Updated')
        regular_count = len([branch for branch in branches if not branch.is_special])
        special_count = len([branch for branch in branches if branch.is_special])
        output.print_md('* Circuits updated: **{}**'.format(regular_count))
        if special_count:
            output.print_md('* SPARE/SPACE circuits refreshed: **{}**'.format(special_count))
        output.print_md('* Fixtures and Devices updated: **{}**'.format(total_fixtures))
        output.print_md('* Electrical Equipment updated: **{}**'.format(total_equipment))

        if locked_rows:
            output.print_md('\n## Skipped Elements')
            output.print_md('The following elements are owned by other users and could not be calculated.')
            table = []
            for row in locked_rows:
                table.append([
                    row.get('circuit', ''),
                    row.get('circuit_owner', '') or '-',
                    row.get('device_owner', '') or '-',
                ])
            output.print_table(table_data=table, columns=['Circuit', 'Circuit Owner', 'Device Owner'])

        label_map = {
            'Overrides': 'Overrides',
            'Calculation': 'Calculation',
            'Data': 'Data',
            'Design': 'Design',
            'Error': 'Error',
            'Other': 'Other',
        }
        severity_colors = {
            'NONE': None,
            'MEDIUM': '#d9822b',
            'HIGH': '#d9534f',
            'CRITICAL': '#b20000',
        }

        notice_lines = []
        for branch in branches:
            if not getattr(branch, 'notices', None) or not branch.notices.has_items():
                continue
            notice_lines.extend(branch.notices.formatted_lines(label_map, severity_colors))

        if notice_lines:
            output.print_md('\n## Warnings / Errors')
            for line in notice_lines:
                output.print_md(line)
        try:
            output.show()
        except Exception:
            pass

    def _collect_runtime_alert_rows(self, branches):
        rows = []
        for branch in branches:
            notices = getattr(branch, 'notices', None)
            if not notices or not notices.has_items():
                continue
            for definition, severity, group, message in notices.items:
                if definition is not None and getattr(definition, 'persistent', True):
                    continue
                definition_id = ''
                try:
                    definition_id = definition.GetId() if definition else ''
                except Exception:
                    definition_id = ''
                rows.append({
                    'panel': branch.panel or '',
                    'number': branch.circuit_number or '',
                    'load_name': branch.load_name or '',
                    'group': group or 'Other',
                    'definition_id': definition_id or '-',
                    'message': message or '',
                })
        return rows

    def _write_locked_sync_payloads(self, doc, locked_rows):
        for row in list(locked_rows or []):
            try:
                if not bool(row.get('sync_writeback', False)):
                    continue
                circuit_id = self._coerce_boundary_id(row.get('circuit_id'))
                if circuit_id <= 0:
                    continue
                circuit = doc.GetElement(_elid_from_value(circuit_id))
                if circuit is None:
                    continue
                payload = self.alert_store.read_alert_payload(circuit) or {}
                if not isinstance(payload, dict):
                    payload = {}
                alerts = payload.get('alerts')
                payload['alerts'] = alerts if isinstance(alerts, list) else []
                hidden = payload.get('hidden_definition_ids')
                payload['hidden_definition_ids'] = hidden if isinstance(hidden, list) else []
                payload['version'] = payload.get('version') or 1
                payload['sync_lock'] = {
                    'blocked': True,
                    'generated_utc': datetime.utcnow().isoformat() + 'Z',
                    'circuit_owner': row.get('circuit_owner') or '',
                    'device_owner': row.get('device_owner') or '',
                }
                self.alert_store.write_alert_payload(circuit, payload)
            except Exception:
                continue


class ApplyCalculatedCircuitsOperation(object):
    """Explicit operation-key adapter for staged Calculate Circuits results."""

    key = 'apply_calculated_circuits'

    def __init__(self, calculate_operation):
        self._calculate_operation = calculate_operation

    def execute(self, request, doc):
        staged_result = request.options.get('staged_calculation')
        if not isinstance(staged_result, dict):
            staged_result = request.options.get('staged_result')
        if not isinstance(staged_result, dict):
            return {
                'status': 'cancelled',
                'reason': 'missing_staged_result',
            }
        return self._calculate_operation.apply_staged_result(
            request,
            doc,
            staged_result,
        )

