# -*- coding: utf-8 -*-
"""Create Circuits by Device Parameter - apply layer (wraps the Revit transaction).

Creates one native Revit ElectricalSystem (power circuit) per plan, assigns the
chosen panel, sets the circuit rating + load name, and mirrors the panel and
rating onto the CKT_* shared parameters of each member so the CED data model
stays in sync. The load name is set on the circuit ONLY - each member's
CKT_Load Name_CEDT is deliberately left untouched.
"""

from pyrevit import DB

import cg_collect
import cg_core
from Snippets import revit_helpers


# ---------------------------------------------------------------------------
# Connector / circuit helpers
# ---------------------------------------------------------------------------
class SwallowCircuitGrouperWarnings(DB.IFailuresPreprocessor):
    """Delete Revit warnings raised by circuit creation, never errors."""

    def PreprocessFailures(self, failures_accessor):
        for failure in failures_accessor.GetFailureMessages():
            if failure.GetSeverity() == DB.FailureSeverity.Warning:
                failures_accessor.DeleteWarning(failure)
        return DB.FailureProcessingResult.Continue


def _configure_failure_handling(transaction):
    """Install the warning-only preprocessor on a started Revit transaction."""
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(SwallowCircuitGrouperWarnings())
    options.SetClearAfterRollback(True)
    forced_modal = getattr(options, "SetForcedModalHandling", None)
    if callable(forced_modal):
        forced_modal(False)
    transaction.SetFailureHandlingOptions(options)


class _CircuitTransaction(object):
    """Small context manager exposing a transaction with warning swallowing."""

    def __init__(self, doc, name):
        self._transaction = DB.Transaction(doc, name)
        self._started = False

    def __enter__(self):
        self._transaction.Start()
        self._started = True
        _configure_failure_handling(self._transaction)
        return self._transaction

    def __exit__(self, exc_type, exc_value, traceback):
        if not self._started:
            return False
        try:
            if exc_type is None:
                self._transaction.Commit()
            else:
                self._transaction.RollBack()
        except Exception:
            try:
                self._transaction.RollBack()
            except Exception:
                pass
            raise
        return False


def _set_text(elem, name, value):
    if value is None:
        return
    try:
        p = elem.LookupParameter(name)
        if p and not p.IsReadOnly:
            p.Set(str(value))
    except Exception:
        pass


def _set_double(elem, name, value):
    if value is None:
        return
    try:
        p = elem.LookupParameter(name)
        if p and not p.IsReadOnly:
            p.Set(float(value))
    except Exception:
        pass


def _get_element_from_token(doc, token):
    """Resolve an element without needlessly round-tripping its ElementId.

    Create Circuits by Device Parameter keeps Revit ElementId objects in its panel/circuit maps.
    The numeric fallback is retained only for compatibility with an older
    caller that may still provide an integer-like token.
    """
    if token is None:
        return None
    try:
        element = doc.GetElement(token)
        if element is not None:
            return element
    except Exception:
        pass
    try:
        value = revit_helpers.get_elementid_value(token)
        if value:
            return doc.GetElement(revit_helpers.elementid_from_value(value))
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Main entry
# ---------------------------------------------------------------------------
def run(doc, plans, name_to_id, logger=None):
    """Create circuits for each plan; panel assignment is a separate service step."""
    report = {
        "created": 0,
        "members_circuited": 0,
        "skipped_no_connector": [],       # element ids
        "skipped_unavailable_primary": [], # element ids
        "errors": [],                 # (group_key, message)
        "lines": [],                  # human-readable per-group summary
        "created_circuit_ids_by_panel": {},
        "circuited_element_ids": [],
        "circuit_results": [],
        "circuit_failures": [],
        "attempted_source_group_keys": [],
    }

    if not plans:
        return report

    for plan in list(plans or []):
        source_key = plan.get("source_group_key") or plan.get("group_key") or ""
        if source_key not in report["attempted_source_group_keys"]:
            report["attempted_source_group_keys"].append(source_key)

    with _CircuitTransaction(doc, "Create Circuits by Device Parameter - Create Circuits"):
        for plan in plans:
            key = plan.get("group_key", "")
            panel_name = plan.get("panel", "")
            load_name = plan.get("load_name", "") or key
            schedule_notes = plan.get("schedule_notes", "") or ""
            amps = plan.get("rating_amps")

            elems = []
            for eid in plan.get("element_ids", []):
                el = doc.GetElement(revit_helpers.elementid_from_value(eid))
                if el is not None:
                    elems.append(el)

            circuitable = []
            for el in elems:
                connector = cg_collect.primary_power_connector(el)
                if connector is None:
                    report["skipped_no_connector"].append(
                        cg_collect.element_id_value(el.Id)
                    )
                    continue
                if not cg_collect.primary_power_connector_is_unused(connector):
                    report["skipped_unavailable_primary"].append(
                        cg_collect.element_id_value(el.Id)
                    )
                    continue
                circuitable.append((el, connector))

            if not circuitable:
                failure_message = (
                    "No members have an unused primary power connector.")
                report["errors"].append((key, failure_message))
                report["circuit_failures"].append({
                    "group_key": key,
                    "source_group_key": plan.get("source_group_key") or key,
                    "member_element_ids": list(plan.get("element_ids") or []),
                    "detail": failure_message,
                })
                report["lines"].append(
                    "[{}] skipped - no unused primary power connectors on any member".format(
                        key)
                )
                continue

            system = None
            try:
                system = DB.Electrical.ElectricalSystem.Create(
                    circuitable[0][1],
                    DB.Electrical.ElectricalSystemType.PowerCircuit,
                )
                if system is None:
                    raise Exception("Revit did not create a circuit from the primary connector.")
                if len(circuitable) > 1:
                    connectors = DB.ConnectorSet()
                    for _, connector in circuitable[1:]:
                        connectors.Insert(connector)
                    system.Add(connectors)

                # Mirror panel + rating only after every selected primary
                # connector has been added successfully. The load name stays
                # on the native circuit; member CKT_Load Name_CEDT values are
                # deliberately preserved.
                for el, _ in circuitable:
                    _set_text(el, cg_collect.PARAM_PANEL, panel_name)
                    if amps is not None:
                        _set_double(el, cg_collect.PARAM_RATING, amps)

                # Preserve all requested circuit data before panel assignment.
                # This leaves a useful, fully described circuit for a manual
                # move when Revit cannot assign it to the requested panel.
                _set_text(system, cg_collect.PARAM_PANEL, panel_name)
                _set_double(system, cg_collect.PARAM_RATING, amps)
                _set_text(system, cg_collect.PARAM_LOAD_NAME, load_name)
                _set_text(
                    system,
                    cg_collect.PARAM_SCHEDULE_NOTES,
                    schedule_notes,
                )

                # rating + load name on the native circuit
                try:
                    if amps is not None:
                        rp = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM)
                        if rp and not rp.IsReadOnly:
                            rp.Set(float(amps))
                except Exception:
                    pass
                try:
                    np = system.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
                    if np and not np.IsReadOnly and load_name:
                        np.Set(str(load_name))
                except Exception:
                    pass
                try:
                    notes_param = system.get_Parameter(
                        DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NOTES_PARAM)
                    if notes_param and not notes_param.IsReadOnly and schedule_notes:
                        notes_param.Set(str(schedule_notes))
                except Exception:
                    pass
            except Exception as ex:
                # Connector-based creation can succeed before adding the
                # remaining connectors fails. Delete only that brand-new
                # circuit so the surrounding creation transaction can keep
                # processing the other groups without touching any existing
                # circuit or using a subtransaction per group.
                if system is not None:
                    try:
                        doc.Delete(system.Id)
                    except Exception as delete_ex:
                        # Do not commit a partial new circuit. A failure to
                        # delete the just-created system is exceptional enough
                        # to roll back the single creation transaction.
                        raise Exception(
                            "Could not clean up the partial new circuit: {}"
                            .format(delete_ex))
                message = "circuit creation failed: {}".format(ex)
                report["errors"].append((key, message))
                report["circuit_failures"].append({
                    "group_key": key,
                    "source_group_key": plan.get("source_group_key") or key,
                    "member_element_ids": list(plan.get("element_ids") or []),
                    "detail": message,
                })
                report["lines"].append("[{}] ERROR creating circuit: {}".format(
                    key, message))
                continue

            report["created"] += 1
            report["members_circuited"] += len(circuitable)
            report["circuited_element_ids"].extend([
                cg_collect.element_id_value(element.Id)
                for element, _connector in circuitable
            ])
            report["circuit_results"].append({
                "group_key": key,
                "source_group_key": plan.get("source_group_key") or key,
                "dedicated": bool(plan.get("dedicated", False)),
                "circuit_id": cg_collect.element_id_value(system.Id),
                "member_element_ids": [
                    cg_collect.element_id_value(element.Id)
                    for element, _connector in circuitable
                ],
                "target_panel": panel_name,
                "load_name": load_name,
            })

            panel_bucket = report["created_circuit_ids_by_panel"].setdefault(
                panel_name, [])
            # Preserve the original API object for the later Move Selected
            # Circuits assignment step.
            panel_bucket.append(system.Id)

            report["lines"].append(
                "[{}] circuit created: {} member(s), target panel '{}', {}".format(
                    key,
                    len(circuitable),
                    panel_name or "(unassigned)",
                    cg_core.format_amps(amps) or "no rating",
                )
            )

    return report


def _same_revit_element_id(left, right):
    """Compare native Revit ElementId objects without converting them."""
    if left is None or right is None:
        return False
    try:
        if left == right:
            return True
    except Exception:
        pass
    try:
        equals = getattr(left, "Equals", None)
        if callable(equals) and bool(equals(right)):
            return True
    except Exception:
        pass
    return False


def _safe_parameter_text(element, built_in_parameter):
    try:
        parameter = element.get_Parameter(built_in_parameter)
        if parameter is None or not parameter.HasValue:
            return ""
        value = parameter.AsString()
        if value:
            return str(value)
        value = parameter.AsValueString()
        return str(value) if value else ""
    except Exception:
        return ""


def _safe_element_name(element, fallback=""):
    try:
        value = getattr(element, "Name", None)
        if value:
            return str(value)
    except Exception:
        pass
    return str(fallback or "")


def _panel_display_name(panel, fallback=""):
    if panel is None:
        return str(fallback or "")
    name = _safe_parameter_text(
        panel, DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
    return name or _safe_element_name(panel, fallback)


def _circuit_display_label(circuit):
    number = _safe_parameter_text(
        circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
    name = _safe_parameter_text(
        circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
    if number and name:
        return "{} / {}".format(number, name)
    if number:
        return number
    if name:
        return name
    return "Created circuit"


def _circuit_id_display(circuit):
    """Return a display-only id; native ids remain native for all lookups."""
    try:
        value = revit_helpers.get_elementid_value(circuit.Id)
        if int(value) > 0:
            return str(value)
    except Exception:
        pass
    try:
        return str(circuit.Id)
    except Exception:
        return "-"


def _assignment_status_rows(doc, circuits, target_panel, target_panel_name):
    """Read the actual panel after assignment, without trusting the target name."""
    rows = []
    for original_circuit in list(circuits or []):
        circuit = original_circuit
        try:
            refreshed = doc.GetElement(original_circuit.Id)
            if refreshed is not None:
                circuit = refreshed
        except Exception:
            pass

        try:
            actual_panel = getattr(circuit, "BaseEquipment", None)
        except Exception:
            actual_panel = None
        actual_panel_name = _panel_display_name(actual_panel, "")
        circuit_number = _safe_parameter_text(
            circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        circuit_name = _safe_parameter_text(
            circuit, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME)
        if actual_panel is None:
            status = "UNASSIGNED"
            actual_panel_name = "(none)"
        elif target_panel is not None and _same_revit_element_id(
                getattr(actual_panel, "Id", None),
                getattr(target_panel, "Id", None)):
            status = "ASSIGNED"
        else:
            status = "ON OTHER PANEL"

        rows.append({
            "circuit": _circuit_display_label(circuit),
            "circuit_id": cg_collect.element_id_value(circuit.Id),
            "element_id": _circuit_id_display(circuit),
            "circuit_number": circuit_number or "-",
            "load_name": circuit_name or "",
            "target_panel": str(target_panel_name or "(none)"),
            "actual_panel": actual_panel_name or "(unnamed panel)",
            "status": status,
            "detail": (
                "Assigned to the intended target panel."
                if status == "ASSIGNED" else
                "Circuit was created but has no panel assignment."
                if status == "UNASSIGNED" else
                "Circuit was created but is assigned to a different panel."
            ),
        })
    return rows


def _is_expected_assignment_failure(error):
    """Identify capacity/user-decision failures that need a concise warning."""
    text = str(error or "").lower()
    expected_phrases = (
        "insufficient valid slot capacity",
        "target panel cannot fit",
        "no circuits could be moved",
        "user chose rollback",
        "move was not confirmed",
        "assignment canceled",
        "selected panel was not found",
        "no created circuits were found",
    )
    return any(phrase in text for phrase in expected_phrases)


def _record_assignment_status(result, status_rows):
    result["circuit_status"].extend(list(status_rows or []))
    for row in list(status_rows or []):
        if row.get("status") != "ASSIGNED":
            result["not_on_target"] += 1
        if row.get("status") == "UNASSIGNED":
            result["unassigned"] += 1


def _reconcile_assigned_panel_metadata(doc, status_rows):
    """Mirror an actual assignment without erasing an unassigned target.

    An unassigned circuit intentionally retains its requested panel metadata so
    a user can identify the correct manual destination. Once Revit assigns a
    circuit, however, the shared circuit/device panel value must describe the
    panel that actually owns it.
    """
    assigned_rows = [
        row for row in list(status_rows or [])
        if row.get("status") in ("ASSIGNED", "ON OTHER PANEL")
        and str(row.get("actual_panel") or "").strip()
    ]
    if not assigned_rows:
        return 0

    updated = 0
    with _CircuitTransaction(
            doc,
            "Create Circuits by Device Parameter - Reconcile Panel Metadata"):
        for row in assigned_rows:
            circuit_id = revit_helpers.elementid_from_value(
                row.get("circuit_id")
            )
            circuit = doc.GetElement(circuit_id)
            if circuit is None:
                continue
            actual_panel = str(row.get("actual_panel") or "").strip()
            _set_text(circuit, cg_collect.PARAM_PANEL, actual_panel)
            try:
                members = list(circuit.Elements)
            except Exception:
                members = []
            for member in members:
                _set_text(member, cg_collect.PARAM_PANEL, actual_panel)
            updated += 1
    return updated


def _assign_created_to_unscheduled_panel(doc, circuits, target_panel):
    """Assign new, unassigned circuits without schedule-capacity analysis.

    With no PanelScheduleView there are no schedule rows to preserve, replace,
    or backfill. Revit's native SelectPanel validation is the authority here;
    running the full Move Selected Circuits capacity repository adds a large
    scan without providing schedule-row protection.
    """
    moved = []
    failed = []
    transaction = DB.Transaction(
        doc, "Assign Created Circuits to Unscheduled Panel")
    transaction.Start()
    try:
        for circuit in list(circuits or []):
            subtransaction = DB.SubTransaction(doc)
            subtransaction.Start()
            try:
                selected = circuit.SelectPanel(target_panel)
                if isinstance(selected, bool) and not selected:
                    raise Exception("SelectPanel returned False.")
                actual_panel = getattr(circuit, "BaseEquipment", None)
                if not _same_revit_element_id(
                        getattr(actual_panel, "Id", None),
                        getattr(target_panel, "Id", None)):
                    raise Exception(
                        "Revit did not assign the circuit to the target panel.")
                subtransaction.Commit()
                moved.append(circuit)
            except Exception as ex:
                try:
                    subtransaction.RollBack()
                except Exception:
                    pass
                failed.append([
                    _circuit_display_label(circuit),
                    _safe_element_name(target_panel, "(unnamed panel)"),
                    str(ex).strip() or "Panel assignment failed.",
                ])
        transaction.Commit()
    except Exception:
        try:
            transaction.RollBack()
        except Exception:
            pass
        raise
    return {
        "moved": moved,
        "failed": failed,
        "skipped": [],
        "partial": bool(failed),
        "fallback_used": False,
    }


def assign_created_circuits_to_panels(doc, created_by_panel, name_to_id,
                                      logger=None):
    """Assign newly created circuits through Move Selected Circuits.

    Scheduled panels delegate capacity decisions to the existing
    ``move_circuits_to_panel_service`` path. That path remains the source of
    truth for fit-without-defaults, removable SPARE/SPACE handling, partial
    moves, and restoration/backfill. After explicit confirmation, a target
    without a PanelScheduleView uses Revit's direct SelectPanel operation;
    there are no schedule rows to analyze or protect in that case.
    """
    result = {
        "moved": 0,
        "failed": [],
        "errors": [],
        "fallback_used": False,
        "buffered_events": [],
        "circuit_status": [],
        "not_on_target": 0,
        "unassigned": 0,
        "panel_metadata_reconciled": 0,
    }
    if not created_by_panel:
        return result

    try:
        from pyrevit import forms
        from CEDElectrical.Application.operations.move_selected_circuits_operation import (
            BufferedMoveOutput,
        )
        from CEDElectrical.Application.services.move_circuits_to_panel_service import (
            move_circuits_to_panel,
        )
        from Snippets._elecutils import (
            MOVE_MISSING_PANEL_SCHEDULE_WARNING,
            move_target_requires_schedule_confirmation,
        )
    except Exception as ex:
        for panel_name, circuit_ids in dict(created_by_panel or {}).items():
            circuits = []
            for circuit_id in list(circuit_ids or []):
                circuit = _get_element_from_token(doc, circuit_id)
                if circuit is not None:
                    circuits.append(circuit)
            target_panel = _get_element_from_token(
                doc, name_to_id.get(str(panel_name or "").strip()))
            result["errors"].append((
                panel_name,
                "Move Selected Circuits service unavailable: {}".format(ex)))
            _record_assignment_status(
                result,
                _assignment_status_rows(doc, circuits, target_panel, panel_name),
            )
        if logger is not None:
            try:
                logger.exception("Create Circuits by Device Parameter could not load Move Selected Circuits service: %s", ex)
            except Exception:
                pass
        return result

    for panel_name, circuit_ids in dict(created_by_panel or {}).items():
        panel_name = str(panel_name or "").strip()
        if not panel_name:
            continue
        panel = None
        circuits = []
        try:
            for circuit_id in list(circuit_ids or []):
                circuit = _get_element_from_token(doc, circuit_id)
                if circuit is not None:
                    circuits.append(circuit)
            if not circuits:
                raise Exception("No created circuits were found for this panel.")

            panel = _get_element_from_token(doc, name_to_id.get(panel_name))
            if panel is None:
                raise Exception("Selected panel was not found in the active document.")

            # Match Move Selected Circuits' explicit confirmation for panels
            # that do not yet have a panel schedule view. The circuits have
            # already been created, so cancelling leaves them safely created
            # and unassigned rather than rolling back circuit creation.
            target_has_no_schedule = bool(
                move_target_requires_schedule_confirmation(doc, panel))
            if target_has_no_schedule:
                proceed = bool(forms.alert(
                    MOVE_MISSING_PANEL_SCHEDULE_WARNING,
                    title="Create Circuits by Device Parameter",
                    ok=True,
                    cancel=True,
                    warn_icon=True,
                ))
                if not proceed:
                    result["errors"].append((
                        panel_name,
                        "Panel assignment canceled because the target panel has no panel schedule view.",
                    ))
                    continue

            buffered = BufferedMoveOutput()
            if target_has_no_schedule:
                move_result = _assign_created_to_unscheduled_panel(
                    doc, circuits, panel)
            else:
                move_result = move_circuits_to_panel(
                    circuits,
                    panel,
                    doc,
                    buffered,
                    allow_unassigned_partial=True,
                    consolidate_fallback_transaction=True,
                    defer_schedule_analysis=True,
                )
            move_result = move_result if isinstance(move_result, dict) else {}
            moved = list(move_result.get("moved") or [])
            failed = list(move_result.get("failed") or [])
            result["moved"] += len(moved)
            result["failed"].extend(failed)
            result["fallback_used"] = bool(
                result["fallback_used"] or
                move_result.get("fallback_used", False))
            result["buffered_events"].extend(
                list(getattr(buffered, "_events", []) or []))
        except Exception as ex:
            message = str(ex).strip() or "Panel assignment failed."
            result["errors"].append((panel_name, message))
            if logger is not None:
                try:
                    if _is_expected_assignment_failure(ex):
                        logger.warning(
                            "Create Circuits by Device Parameter could not assign created circuits to %s: %s",
                            panel_name,
                            message,
                        )
                    else:
                        # Unexpected failures retain their traceback in pyRevit
                        # output; expected capacity failures stay readable.
                        logger.exception(
                            "Create Circuits by Device Parameter panel assignment failed for %s: %s",
                            panel_name,
                            ex,
                        )
                except Exception:
                    pass
        finally:
            status_rows = _assignment_status_rows(
                doc, circuits, panel, panel_name)
            _record_assignment_status(result, status_rows)
    try:
        result["panel_metadata_reconciled"] = _reconcile_assigned_panel_metadata(
            doc,
            result["circuit_status"],
        )
    except Exception as ex:
        # Panel metadata is secondary to keeping the fully described circuits
        # that were already created. Report reconciliation trouble without
        # causing the enclosing workflow group to roll the circuits back.
        result["errors"].append((
            "Panel metadata",
            "Could not reconcile actual panel metadata: {}".format(ex),
        ))
        if logger is not None:
            try:
                logger.exception(
                    "Create Circuits by Device Parameter panel metadata reconciliation failed: %s",
                    ex,
                )
            except Exception:
                pass
    return result
