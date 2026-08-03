# -*- coding: utf-8 -*-
"""Circuit Grouper - apply layer (wraps the Revit transaction).

Creates one native Revit ElectricalSystem (power circuit) per plan, assigns the
chosen panel, sets the circuit rating + load name, and mirrors the panel and
rating onto the CKT_* shared parameters of each member so the CED data model
stays in sync. The load name is set on the circuit ONLY - each member's
CKT_Load Name_CEDT is deliberately left untouched.
"""

from System.Collections.Generic import List
from pyrevit import DB, revit

import cg_core
import cg_collect


# ---------------------------------------------------------------------------
# Connector / circuit helpers
# ---------------------------------------------------------------------------
def _has_power_connector(elem):
    # single definition lives in cg_collect (used by both collection + apply)
    return cg_collect.has_power_connector(elem)


def _existing_systems(elem):
    mep = getattr(elem, "MEPModel", None)
    if mep is None:
        return []
    for attr in ("GetElectricalSystems", "GetAssignedElectricalSystems"):
        fn = getattr(mep, attr, None)
        if callable(fn):
            try:
                res = fn()
                if res:
                    return list(res)
            except Exception:
                pass
    return []


def _remove_from_existing(elem):
    removed = 0
    for system in _existing_systems(elem):
        try:
            ids = List[DB.ElementId]()
            ids.Add(elem.Id)
            system.RemoveFromCircuit(ids)
            removed += 1
        except Exception:
            pass
    return removed


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


def _assign_panel(system, panel_element):
    if panel_element is None:
        return False, "no panel selected"
    can = True
    check = getattr(system, "CanAssignToPanel", None)
    if callable(check):
        try:
            can = bool(check(panel_element))
        except Exception:
            can = True
    if not can:
        return False, "panel not compatible (poles/voltage)"
    try:
        system.SelectPanel(panel_element)
        return True, ""
    except Exception as ex:
        return False, "SelectPanel failed: {}".format(ex)


def _to_display_volts(internal):
    """Convert Revit's internal electrical-potential value to display volts."""
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertFromInternalUnits(internal, UnitTypeId.Volts)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertFromInternalUnits(internal, DisplayUnitType.DUT_VOLTS)
    except Exception:
        return internal


def _panel_distribution_label(panel_element):
    """Return the panel's distribution-system display string, e.g. '208Y/120V'.

    The distribution-system parameter's AsValueString() already gives the clean
    label, so no element/voltage parsing is needed."""
    for attr in ("RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM",
                 "RBS_FAMILY_CONTENT_SECONDARY_DISTRIBSYS"):
        bip = getattr(DB.BuiltInParameter, attr, None)
        if bip is None:
            continue
        try:
            p = panel_element.get_Parameter(bip)
        except Exception:
            p = None
        if p and p.HasValue:
            try:
                label = p.AsValueString()
            except Exception:
                label = None
            if label:
                return label
    return ""


def _panel_mismatch_detail(system, panel_element):
    parts = []
    # --- circuit side ---
    try:
        poles = getattr(system, "PolesNumber", None)
        if poles:
            parts.append("circuit {}P".format(poles))
    except Exception:
        pass
    try:
        v = getattr(system, "Voltage", None)
        if v:
            parts.append("circuit {:g}V".format(round(_to_display_volts(v))))
    except Exception:
        pass
    # --- panel side ---
    panel_bits = []
    ds_label = _panel_distribution_label(panel_element)
    if ds_label:
        panel_bits.append(ds_label)
    try:
        p = panel_element.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NUMPHASES_PARAM)
        if p and p.HasValue:
            panel_bits.append("{}-phase".format(p.AsInteger()))
    except Exception:
        pass
    if panel_bits:
        parts.append("panel {}".format(", ".join(panel_bits)))

    return ", ".join(parts) if parts else "distribution system / voltage / poles do not match"


def preflight_panel_check(doc, plans, name_to_id):
    """Find groups whose chosen panel is incompatible with the circuit BEFORE
    committing anything. Runs inside a transaction that is always rolled back,
    so the model is left untouched. Returns a list of dicts:
        {"group_key", "panel", "detail"}
    Only groups that (a) have a panel selected and (b) have at least one
    connectable member are evaluated.
    """
    issues = []
    if not plans:
        return issues

    tx = DB.Transaction(doc, "Circuit Grouper - Preflight (rolled back)")
    tx.Start()
    try:
        for plan in plans:
            panel_name = plan.get("panel", "")
            pid = name_to_id.get(panel_name)
            if not pid:
                continue  # no panel chosen -> nothing to mismatch against

            elems = []
            for eid in plan.get("element_ids", []):
                el = doc.GetElement(DB.ElementId(int(eid)))
                if el is not None and _has_power_connector(el):
                    elems.append(el)
            if not elems:
                continue

            for el in elems:
                _remove_from_existing(el)

            ids = List[DB.ElementId]()
            for el in elems:
                ids.Add(el.Id)
            try:
                system = DB.Electrical.ElectricalSystem.Create(
                    doc, ids, DB.Electrical.ElectricalSystemType.PowerCircuit
                )
            except Exception:
                continue  # creation failures are surfaced by run()

            panel_element = doc.GetElement(DB.ElementId(int(pid)))
            compatible = True
            reason = ""

            # CanAssignToPanel is a fast first filter...
            check = getattr(system, "CanAssignToPanel", None)
            if callable(check):
                try:
                    if not bool(check(panel_element)):
                        compatible = False
                except Exception:
                    pass

            # ...but it can return True while SelectPanel still throws, so we
            # actually attempt the assignment here (the whole transaction is
            # rolled back, so this never persists). This mirrors run() exactly.
            if compatible:
                try:
                    system.SelectPanel(panel_element)
                except Exception as ex:
                    compatible = False
                    reason = str(ex)

            if not compatible:
                detail = _panel_mismatch_detail(system, panel_element)
                if reason:
                    detail = "{} ({})".format(reason.strip().rstrip("."), detail)
                issues.append({
                    "group_key": plan.get("group_key", ""),
                    "panel": panel_name,
                    "detail": detail,
                })
    finally:
        tx.RollBack()

    return issues


# ---------------------------------------------------------------------------
# Main entry
# ---------------------------------------------------------------------------
def run(doc, plans, name_to_id, logger=None):
    """Create circuits for each plan. Returns a report dict."""
    report = {
        "created": 0,
        "members_circuited": 0,
        "removed_from_existing": 0,
        "skipped_no_connector": [],   # element ids
        "panel_warnings": [],         # (group_key, message)
        "errors": [],                 # (group_key, message)
        "lines": [],                  # human-readable per-group summary
    }

    if not plans:
        return report

    with revit.Transaction("Circuit Grouper - Create Circuits"):
        for plan in plans:
            key = plan.get("group_key", "")
            panel_name = plan.get("panel", "")
            load_name = plan.get("load_name", "") or key
            amps = plan.get("rating_amps")

            elems = []
            for eid in plan.get("element_ids", []):
                el = doc.GetElement(DB.ElementId(int(eid)))
                if el is not None:
                    elems.append(el)

            circuitable = []
            for el in elems:
                if _has_power_connector(el):
                    circuitable.append(el)
                else:
                    report["skipped_no_connector"].append(
                        cg_collect.element_id_value(el.Id)
                    )

            # mirror panel + rating CKT_* onto every member regardless of
            # connector status. The load name is NOT mirrored - members'
            # CKT_Load Name_CEDT stays whatever it already was; the chosen
            # name goes on the native circuit below.
            for el in elems:
                _set_text(el, cg_collect.PARAM_PANEL, panel_name)
                if amps is not None:
                    _set_double(el, cg_collect.PARAM_RATING, amps)

            if not circuitable:
                report["errors"].append((key, "no members have a power connector"))
                report["lines"].append(
                    "[{}] skipped - no power connectors on any member".format(key)
                )
                continue

            for el in circuitable:
                report["removed_from_existing"] += _remove_from_existing(el)

            ids = List[DB.ElementId]()
            for el in circuitable:
                ids.Add(el.Id)

            try:
                system = DB.Electrical.ElectricalSystem.Create(
                    doc, ids, DB.Electrical.ElectricalSystemType.PowerCircuit
                )
            except Exception as ex:
                report["errors"].append((key, "circuit creation failed: {}".format(ex)))
                report["lines"].append("[{}] ERROR creating circuit: {}".format(key, ex))
                continue

            report["created"] += 1
            report["members_circuited"] += len(circuitable)

            # panel
            panel_element = None
            pid = name_to_id.get(panel_name)
            if pid:
                panel_element = doc.GetElement(DB.ElementId(int(pid)))
            ok, msg = _assign_panel(system, panel_element)
            if not ok and msg:
                report["panel_warnings"].append((key, msg))

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

            report["lines"].append(
                "[{}] circuit created: {} member(s), panel '{}', {}".format(
                    key,
                    len(circuitable),
                    panel_name or "(unassigned)",
                    cg_core.format_amps(amps) or "no rating",
                )
            )

    return report
