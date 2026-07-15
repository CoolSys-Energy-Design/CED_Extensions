# -*- coding: utf-8 -*-

"""Calculate-circuits application operation."""

from datetime import datetime

from pyrevit import DB, forms, script

from CEDElectrical.Domain import settings_manager
from CEDElectrical.Model.CircuitBranch import CircuitBranch
from CEDElectrical.Model.circuit_settings import CircuitSettings
from Snippets import revit_helpers


def _elid_value(item):
    return revit_helpers.get_elementid_value(item)


def _elid_from_value(value):
    return revit_helpers.elementid_from_value(value)


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
        param_bootstrap = settings_manager.ensure_electrical_parameters_for_calculate(doc, logger=self.logger)
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
        circuits = self.repository.get_target_circuits(doc, request.circuit_ids)
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

        circuits, locked_ids, locked_rows = self.repository.partition_locked_elements(doc, circuits, settings)
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
        existing_values_by_id = {}
        for circuit in circuits:
            cid = _elid_value(circuit.Id)
            circuit_data_payload = self._read_circuit_data_payload(circuit)
            existing_values_by_id[cid] = dict(
                supplied_existing_values_by_id.get(cid)
                or self._collect_existing_preview_values(circuit)
            )
            existing_values_by_id[cid]['_is_first_calculation'] = not bool(circuit_data_payload)
            preview_values = dict(staged_preview_values_by_id.get(cid) or {})
            if cid in force_auto_for_ids:
                preview_values['CKT_User Override_CED'] = 0
            branch = CircuitBranch(circuit, settings=settings, preview_values=preview_values)
            if cid in force_auto_for_ids:
                branch._calc_preview_force_auto = True
            if not branch.is_power_circuit or branch.is_space or branch.is_spare:
                continue

            branch.calculate_hot_wire_size()
            branch.calculate_neutral_wire_size()
            branch.calculate_ground_wire_size()
            branch.calculate_isolated_ground_wire_size()
            branch.calculate_conduit_size()
            branches.append(branch)

        if not branches:
            forms.alert('No editable branch circuits found to process.')
            return {'status': 'cancelled', 'reason': 'no_branches'}

        preview_rows = self._collect_conduit_wire_preview_rows(
            branches,
            existing_values_by_id,
            ignored_preview_fields_by_id,
        )
        preview_changed_ids = self._preview_changed_ids(preview_rows)
        preview_enabled = bool(request.options.get('calc_preview_enabled', False))
        preview_decision = str(request.options.get('calc_preview_decision') or '').strip().lower()
        if preview_enabled and not preview_decision and preview_rows:
            return {
                'status': 'preview_required',
                'reason': 'conduit_wire_changes',
                'preview_rows': preview_rows,
                'locked_rows': locked_rows,
                'runtime_alert_rows': [],
            }

        if preview_decision == 'skip' and preview_changed_ids:
            branches = [
                branch for branch in branches
                if _elid_value(branch.circuit.Id) not in preview_changed_ids
            ]
            if not branches:
                return {
                    'status': 'cancelled',
                    'reason': 'calc_preview_skipped',
                    'locked_rows': locked_rows,
                    'runtime_alert_rows': [],
                    'calc_preview_rows': preview_rows,
                    'calc_preview_decision': preview_decision,
                }

        if preview_decision == 'keep_existing' and preview_rows:
            branches = self._rebuild_branches_with_existing_sizes(
                branches,
                existing_values_by_id,
                staged_preview_values_by_id,
                settings,
            )

        total_fixtures = 0
        total_equipment = 0

        use_existing_group = bool(request.options.get('use_existing_transaction_group', False))
        tg = None
        transaction_name = str(request.options.get('transaction_name') or 'Calculate Circuits').strip() or 'Calculate Circuits'
        write_transaction_name = str(request.options.get('write_transaction_name') or 'Write Shared Parameters').strip() or 'Write Shared Parameters'
        if not use_existing_group:
            tg = DB.TransactionGroup(doc, transaction_name)
            tg.Start()
        tx = DB.Transaction(doc, write_transaction_name)

        try:
            tx.Start()
            for branch in branches:
                self._apply_staged_builtin_values(
                    branch.circuit,
                    staged_builtin_values_by_id.get(_elid_value(branch.circuit.Id)),
                )
                param_values = self._collect_shared_param_values(branch)
                self.writer.write_circuit_parameters(branch.circuit, param_values)
                f_cnt, e_cnt = self.writer.write_connected_elements(branch, param_values, settings, locked_ids)
                total_fixtures += f_cnt
                total_equipment += e_cnt

                alert_payload = self._build_alert_payload(branch)
                if alert_payload is None:
                    self.alert_store.clear_alert_payload(branch.circuit)
                else:
                    self.alert_store.write_alert_payload(branch.circuit, alert_payload)

            self._write_locked_sync_payloads(doc, locked_rows)

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

        show_output = bool(request.options.get('show_output', True))
        if show_output:
            self._print_report(branches, total_fixtures, total_equipment, locked_rows)
        runtime_alert_rows = self._collect_runtime_alert_rows(branches)
        return {
            'status': 'ok',
            'updated_circuits': len(branches),
            'updated_fixtures': total_fixtures,
            'updated_equipment': total_equipment,
            'locked_rows': locked_rows,
            'runtime_alert_rows': runtime_alert_rows,
            'calc_preview_rows': preview_rows,
            'calc_preview_decision': preview_decision,
        }

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
            'CKT_User Override_CED': 1,
            'CKT_Number of Sets_CED': self._lookup_param_value(circuit, 'CKT_Number of Sets_CED', 0),
            'CKT_Include Neutral_CED': self._lookup_param_value(circuit, 'CKT_Include Neutral_CED', 0),
            'CKT_Include Isolated Ground_CED': self._lookup_param_value(circuit, 'CKT_Include Isolated Ground_CED', 0),
            'CKT_Wire Hot Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Hot Size_CEDT', ''),
            'CKT_Wire Neutral Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Neutral Size_CEDT', ''),
            'CKT_Wire Ground Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Ground Size_CEDT', ''),
            'CKT_Wire Isolated Ground Size_CEDT': self._lookup_param_text(circuit, 'CKT_Wire Isolated Ground Size_CEDT', ''),
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
            try:
                value = int(item or 0)
            except Exception:
                value = 0
            if value > 0:
                ids.add(value)
        return ids

    def _normalize_existing_values_by_id(self, raw):
        normalized = {}
        if not isinstance(raw, dict):
            return normalized
        for key, value in list(raw.items()):
            try:
                cid = int(key or 0)
            except Exception:
                cid = 0
            if cid <= 0 or not isinstance(value, dict):
                continue
            normalized[cid] = dict(value)
        return normalized

    def _normalize_values_by_id(self, raw):
        normalized = {}
        if not isinstance(raw, dict):
            return normalized
        for key, value in list(raw.items()):
            try:
                cid = int(key or 0)
            except Exception:
                cid = 0
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
            try:
                cid = int(key or 0)
            except Exception:
                cid = 0
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
            try:
                cid = int(row.get('circuit_id') or 0)
            except Exception:
                cid = 0
            if cid > 0:
                changed_ids.add(cid)
        return changed_ids

    def _rebuild_branches_with_existing_sizes(
            self,
            branches,
            existing_values_by_id,
            staged_preview_values_by_id,
            settings,
    ):
        changed_ids = self._preview_changed_ids(
            self._collect_conduit_wire_preview_rows(branches, existing_values_by_id)
        )
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
            keep_branch = CircuitBranch(branch.circuit, settings=settings, preview_values=preview_values)
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

        changed = False
        try:
            current = getattr(circuit, prop_name)
            if current is None or abs(float(current) - numeric) > 0.0001:
                setattr(circuit, prop_name, numeric)
                changed = True
        except Exception:
            pass

        param = None
        if bip is not None:
            try:
                param = circuit.get_Parameter(bip)
            except Exception:
                param = None
        if not param:
            return changed
        try:
            if param.StorageType == DB.StorageType.Double:
                current = param.AsDouble()
                if current is None or abs(float(current) - numeric) > 0.0001:
                    param.Set(numeric)
                    changed = True
            elif param.StorageType == DB.StorageType.Integer:
                numeric_int = int(round(numeric))
                if param.AsInteger() != numeric_int:
                    param.Set(numeric_int)
                    changed = True
        except Exception:
            pass
        return changed

    def _collect_shared_param_values(self, branch):
        """Map branch results into shared-parameter values."""
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

    def _build_alert_payload(self, branch):
        """Build serializable alert payload for persistence."""
        notices = getattr(branch, 'notices', None)
        existing = self.alert_store.read_alert_payload(branch.circuit) or {}
        if not isinstance(existing, dict):
            existing = {}
        metadata = {}
        for key in self.CIRCUIT_DATA_METADATA_KEYS:
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
        output.print_md('* Circuits updated: **{}**'.format(len(branches)))
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
                circuit_id = int(row.get('circuit_id') or 0)
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

