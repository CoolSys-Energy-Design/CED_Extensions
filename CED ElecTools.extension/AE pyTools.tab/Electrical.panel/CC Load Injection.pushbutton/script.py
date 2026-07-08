# -*- coding: utf-8 -*-
"""CC Load Injection (IronPython 2.7).

Read a refrigeration-schedule workbook ("Electrical Info" format, e.g.
ExampleCarrolltonNewFormat.xlsx), match each circuit ID # to the case
controllers' "Identity Mark" parameter in the model, verify the family's
voltage / pole count against the sheet row, then write:

  * "Ampacity Rating_CED"      <- sum of the five AMPS columns
                                  (DEFROST / FANS / LIGHTS / ANTISWEAT / OTHER)
  * "Apparent Load Ph 1_CED"   <- per the distribution rules below
  * "Apparent Load Ph 2_CED"
  * "Apparent Load Ph 3_CED"

VA math: every component converts at ITS OWN volt column, except that 115 V
columns convert at 120 V for a safety margin (defrost x 208, fans x 120, ...).

Distribution (keyed on the DEFROST voltage config):
  * defrost 208/3 : per-phase defrost VA = V x I / sqrt(3) on Ph1/Ph2/Ph3;
                    lights+antisweat VA added to Ph1, fans+other VA to Ph2.
  * defrost 208/1 : half the defrost VA on Ph1 and half on Ph2;
                    lights+antisweat VA on Ph1, fans+other VA on Ph2.
  * no defrost    : the whole load VA goes on Ph1 (Ph2 / Ph3 = 0).

A VOLTAGE mismatch between sheet row and family skips the element (nothing
is written). A POLE-COUNT difference only gets noted in the preview -
controllers are routinely ganged across phases (e.g. three 2-pole
controllers wired as one 3-pole circuit) and designers review the phase
assignments afterward.
"""

import math
import re

import clr
import System
from pyrevit import DB, forms, revit, script

__title__ = "CC Load\nInjection"

output = script.get_output()
doc = revit.doc

TITLE = "CC Load Injection"

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
IDENTITY_PARAM = "Identity Mark"

PARAM_AMPACITY = "Ampacity Rating_CED"
PARAM_PH = {
    1: "Apparent Load Ph 1_CED",
    2: "Apparent Load Ph 2_CED",
    3: "Apparent Load Ph 3_CED",
}

# Same scope Circuit Grouper uses for case controllers.
DEVICE_CATEGORIES = [
    "OST_MechanicalControlDevices",
    "OST_ElectricalEquipment",
    "OST_ElectricalFixtures",
]

# Voltage / poles read-back for the mismatch check (tried in order).
VOLTAGE_PARAM_NAMES = ["Voltage", "Voltage_CED", "Voltage Nominal_CED"]
POLES_PARAM_NAMES = ["Number of Poles", "Number of Poles_CED"]

LOAD_GROUPS = ["DEFROST", "FANS", "LIGHTS", "ANTISWEAT", "OTHER"]

STANDARD_VOLTAGES = [24, 48, 120, 208, 240, 277, 347, 480, 600]

_VOLT_CONFIG_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*/\s*(\d+)\s*$")


# ---------------------------------------------------------------------------
# Small shared helpers (mirrors Circuit Grouper's cg_core / cg_collect)
# ---------------------------------------------------------------------------
def snap_voltage(volts, tolerance=8.0):
    """Snap to the nearest standard nominal (115 -> 120, 207.8 -> 208)."""
    if volts is None:
        return None
    try:
        v = float(volts)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    best = min(STANDARD_VOLTAGES, key=lambda s: abs(s - v))
    if abs(best - v) <= tolerance:
        return best
    return None


def _lookup(elem, name):
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def _first_param(elem, names):
    for name in names:
        p = _lookup(elem, name)
        if p is not None and p.HasValue:
            return p
    return None


def _spec_is_electrical_potential(param):
    try:
        defn = param.Definition
    except Exception:
        return False
    try:
        from Autodesk.Revit.DB import SpecTypeId
        return defn.GetDataType() == SpecTypeId.ElectricalPotential
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ParameterType
        return defn.ParameterType == ParameterType.ElectricalPotential
    except Exception:
        return False


def _volts_from_internal(raw):
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertFromInternalUnits(raw, UnitTypeId.Volts)
    except Exception:
        pass
    try:  # pre-ForgeTypeId Revit
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertFromInternalUnits(raw, DisplayUnitType.DUT_VOLTS)
    except Exception:
        return raw


def _parse_leading_number(text):
    if not text:
        return None
    num = ""
    for ch in str(text).strip():
        if ch.isdigit() or ch in ".-":
            num += ch
        elif num:
            break
    try:
        return float(num) if num not in ("", "-", ".", "-.") else None
    except ValueError:
        return None


def read_family_voltage(elem):
    """Snapped nominal voltage of the family instance, or None."""
    p = _first_param(elem, VOLTAGE_PARAM_NAMES)
    if p is None:
        return None
    try:
        st = p.StorageType
    except Exception:
        st = None
    volts = None
    if st == DB.StorageType.Double:
        try:
            raw = p.AsDouble()
        except Exception:
            raw = None
        if raw:
            # Only ElectricalPotential specs hold internal units; a plain
            # Number "Voltage" is already in volts.
            volts = _volts_from_internal(raw) if _spec_is_electrical_potential(p) else raw
    elif st == DB.StorageType.Integer:
        try:
            volts = float(p.AsInteger())
        except Exception:
            volts = None
    else:
        try:
            volts = _parse_leading_number(p.AsValueString() or p.AsString())
        except Exception:
            volts = None
    return snap_voltage(volts)


def read_family_poles(elem):
    """Pole count of the family instance, or None."""
    p = _first_param(elem, POLES_PARAM_NAMES)
    if p is None:
        bip = getattr(DB.BuiltInParameter, "RBS_ELEC_NUMBER_OF_POLES", None)
        if bip is not None:
            try:
                cand = elem.get_Parameter(bip)
                if cand is not None and cand.HasValue:
                    p = cand
            except Exception:
                p = None
    if p is None:
        return None
    try:
        if p.StorageType == DB.StorageType.Double:
            d = p.AsDouble()
            return int(round(d)) if d else None
        return p.AsInteger() or None
    except Exception:
        return None


def _to_internal_amps(amps):
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertToInternalUnits(amps, UnitTypeId.Amperes)
    except Exception:
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertToInternalUnits(amps, DisplayUnitType.DUT_AMPERES)


def _to_internal_va(va):
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return UnitUtils.ConvertToInternalUnits(va, UnitTypeId.VoltAmperes)
    except Exception:
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return UnitUtils.ConvertToInternalUnits(va, DisplayUnitType.DUT_VOLT_AMPERES)


# ---------------------------------------------------------------------------
# Excel reading (late-bound COM; no interop assembly / no openpyxl needed)
# ---------------------------------------------------------------------------
def read_workbook_grid(path):
    """Open the workbook read-only, return (grid, nrows, ncols) where grid is
    a dict {(row, col): value} with 1-based absolute worksheet coordinates.
    Prefers a sheet named like 'Electrical Info', else the first sheet."""
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise Exception("Excel is not installed (Excel.Application ProgID not found).")
    app = System.Activator.CreateInstance(excel_type)
    app.Visible = False
    app.DisplayAlerts = False
    wb = None
    try:
        wb = app.Workbooks.Open(path, False, True)  # no-update-links, read-only
        ws = None
        for sheet in wb.Worksheets:
            if "ELECTRICAL INFO" in str(sheet.Name).strip().upper():
                ws = sheet
                break
        if ws is None:
            ws = wb.Worksheets.Item(1)
        ur = ws.UsedRange
        first_row = ur.Row
        first_col = ur.Column
        values = ur.Value2  # one bulk COM call: object[,]
        grid = {}
        nrows = ur.Rows.Count
        ncols = ur.Columns.Count
        if not isinstance(values, System.Array):
            # single-cell UsedRange returns a scalar, not an array
            if values is not None:
                grid[(first_row, first_col)] = values
        else:
            # Excel's COM array is 1-based; iterate its declared bounds via
            # GetValue rather than [] indexing (IronPython's [] on a
            # non-zero-based array walks out of bounds).
            rlo, rhi = values.GetLowerBound(0), values.GetUpperBound(0)
            clo, chi = values.GetLowerBound(1), values.GetUpperBound(1)
            for r in range(rlo, rhi + 1):
                for c in range(clo, chi + 1):
                    v = values.GetValue(r, c)
                    if v is not None:
                        grid[(first_row + r - rlo, first_col + c - clo)] = v
        return grid, first_row + nrows - 1, first_col + ncols - 1, str(ws.Name)
    finally:
        if wb is not None:
            wb.Close(False)
        app.Quit()


def _cell_text(grid, r, c):
    v = grid.get((r, c))
    if v is None:
        return ""
    return str(v).strip()


def _norm_header(text):
    """Uppercase, collapse whitespace/newlines for header comparison."""
    return " ".join(str(text).split()).upper()


def _find_cells(grid, max_row, max_col, wanted):
    """All (row, col) whose normalized text equals ``wanted``."""
    hits = []
    for (r, c), v in grid.items():
        if isinstance(v, str) and _norm_header(v) == wanted:
            hits.append((r, c))
    return sorted(hits)


def locate_columns(grid, max_row, max_col):
    """Header-scan the sheet; returns a dict with the header row and every
    column index the parser needs. Raises with a readable message when the
    format doesn't match."""
    # The load-group labels (DEFROST / FANS / ...) share a row; the AMPS/VOLT
    # sub-headers sit a couple of rows below in the same columns.
    defrost_hits = _find_cells(grid, max_row, max_col, "DEFROST")
    groups_row = None
    group_cols = {}
    for (r, c) in defrost_hits:
        cols = {"DEFROST": c}
        for name in LOAD_GROUPS[1:]:
            for (r2, c2) in _find_cells(grid, max_row, max_col, name):
                if r2 == r:
                    cols[name] = c2
                    break
        if len(cols) >= 3:  # DEFROST + at least FANS/LIGHTS on one row
            groups_row = r
            group_cols = cols
            break
    if groups_row is None:
        raise Exception("Could not find the DEFROST/FANS/LIGHTS group header row.")

    # Sub-header row: the row (within a few rows below) holding 'AMPS' in the
    # DEFROST column.
    sub_row = None
    for r in range(groups_row + 1, groups_row + 5):
        if _norm_header(_cell_text(grid, r, group_cols["DEFROST"])) == "AMPS":
            sub_row = r
            break
    if sub_row is None:
        raise Exception("Could not find the AMPS/VOLT sub-header row.")

    # For each group: AMPS is at the group column; VOLT is the next 'VOLT'
    # cell to the right of it, before the next group's column.
    ordered = sorted(group_cols.items(), key=lambda kv: kv[1])
    bounds = {}
    for i, (name, col) in enumerate(ordered):
        hi = ordered[i + 1][1] if i + 1 < len(ordered) else col + 10
        bounds[name] = (col, hi)
    amp_volt_cols = {}
    for name, (lo, hi) in bounds.items():
        amps_col = None
        volt_col = None
        for c in range(lo, hi):
            h = _norm_header(_cell_text(grid, sub_row, c))
            if h == "AMPS" and amps_col is None:
                amps_col = c
            elif h == "VOLT" and volt_col is None:
                volt_col = c
        if amps_col is None or volt_col is None:
            raise Exception("Missing AMPS/VOLT sub-headers for group '{}'.".format(name))
        amp_volt_cols[name] = (amps_col, volt_col)

    def _single(wanted, required=True):
        hits = _find_cells(grid, max_row, max_col, wanted)
        if not hits:
            if required:
                raise Exception("Could not find the '{}' header.".format(wanted))
            return None
        return hits[0][1]

    id_col = _single("ID #")
    ckt_col = _single("CKT #")
    # Optional cross-check columns (present in this format's summary block).
    total_amps_col = _single("TOTAL AMPS", required=False)
    va1_col = _single("VA PH1", required=False)
    va2_col = _single("VA PH2", required=False)

    return {
        "data_start": sub_row + 1,
        "id_col": id_col,
        "ckt_col": ckt_col,
        "groups": amp_volt_cols,
        "total_amps_col": total_amps_col,
        "va1_col": va1_col,
        "va2_col": va2_col,
    }


def _parse_amps(value):
    """Float amps from a cell; '-' / blank / garbage -> None."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s or s == "-":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_volt_config(value):
    """'208/1' -> (208.0, 1); '-' / blank -> None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s or s == "-":
        return None
    m = _VOLT_CONFIG_RE.match(s)
    if not m:
        return None
    return float(m.group(1)), int(m.group(2))


def parse_rows(grid, max_row, cols):
    """One record per controller row (rows with a CKT # entry). Parent rollup
    rows (ID but no CKT #) carry no electrical data and are skipped."""
    rows = []
    for r in range(cols["data_start"], max_row + 1):
        ckt = _cell_text(grid, r, cols["ckt_col"])
        if not ckt or not re.search(r"[A-Za-z]", ckt):
            continue
        rec = {"row": r, "id": ckt.upper(), "id_display": ckt, "loads": {}}
        for name, (amps_col, volt_col) in cols["groups"].items():
            amps = _parse_amps(grid.get((r, amps_col)))
            volt = _parse_volt_config(grid.get((r, volt_col)))
            if amps and volt:
                rec["loads"][name] = (amps, volt[0], volt[1])
        if not rec["loads"]:
            continue
        for key, col_key in (("total_amps", "total_amps_col"),
                             ("va1", "va1_col"), ("va2", "va2_col")):
            col = cols.get(col_key)
            rec[key] = _parse_amps(grid.get((r, col))) if col else None
        rows.append(rec)
    return rows


# ---------------------------------------------------------------------------
# Load math
# ---------------------------------------------------------------------------
def _va_volts(volts):
    """Conversion voltage for the VA math: 115 V columns convert at 120 V
    for a safety margin; everything else at face value."""
    if volts is not None and 110.0 <= volts < 120.0:
        return 120.0
    return volts


def compute_loads(rec):
    """Returns (total_amps, ph_va{1,2,3}, expected_voltage, expected_poles,
    sheet_basis_va).

    Every component's VA uses its own volt column (115 uprated to 120).
    Distribution is keyed on the defrost voltage config; without defrost the
    whole load goes on Ph1 (the sheet puts these on its PH2 column instead -
    the phase choice is intentional here, the total always reconciles).
    ``sheet_basis_va`` is the total at the sheet's literal voltages, used
    only to cross-check parsing against the sheet's own VA columns."""
    loads = rec["loads"]
    total_amps = sum(amps for (amps, _, _) in loads.values())
    sheet_basis_va = sum(amps * volts for (amps, volts, _) in loads.values())

    def va_of(name):
        if name not in loads:
            return 0.0
        amps, volts, _ = loads[name]
        return amps * _va_volts(volts)

    lights_anti = va_of("LIGHTS") + va_of("ANTISWEAT")
    fans_other = va_of("FANS") + va_of("OTHER")

    defrost = loads.get("DEFROST")
    ph = {1: 0.0, 2: 0.0, 3: 0.0}
    if defrost is not None and defrost[2] == 3:
        amps, volts, _ = defrost
        per_phase = _va_volts(volts) * amps / math.sqrt(3.0)
        ph[1] = per_phase + lights_anti
        ph[2] = per_phase + fans_other
        ph[3] = per_phase
    elif defrost is not None:
        half = va_of("DEFROST") / 2.0
        ph[1] = half + lights_anti
        ph[2] = half + fans_other
    else:
        ph[1] = lights_anti + fans_other

    # Expected family electrical config for the mismatch check.
    max_volts = max(volts for (_, volts, _) in loads.values())
    expected_voltage = snap_voltage(max_volts)
    if defrost is not None and defrost[2] == 3:
        expected_poles = 3
    elif expected_voltage and expected_voltage >= 200:
        expected_poles = 2
    else:
        expected_poles = 1
    return total_amps, ph, expected_voltage, expected_poles, sheet_basis_va


# ---------------------------------------------------------------------------
# Revit collection
# ---------------------------------------------------------------------------
def collect_case_controllers():
    """{identity_mark_upper: [FamilyInstance, ...]} for all case-controller
    category elements with a non-empty Identity Mark."""
    by_mark = {}
    for cat_name in DEVICE_CATEGORIES:
        bic = getattr(DB.BuiltInCategory, cat_name, None)
        if bic is None:
            continue
        collector = (
            DB.FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
        )
        for elem in collector:
            p = _lookup(elem, IDENTITY_PARAM)
            if p is None or not p.HasValue:
                continue
            try:
                mark = (p.AsString() or "").strip()
            except Exception:
                mark = ""
            if not mark:
                continue
            by_mark.setdefault(mark.upper(), []).append(elem)
    return by_mark


def _fmt_va(v):
    return "{:,.2f}".format(v) if v else "0"


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    xl_path = forms.pick_file(
        files_filter="Excel Workbooks (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
        title="Select the refrigeration schedule workbook")
    if not xl_path:
        return

    try:
        grid, max_row, max_col, sheet_name = read_workbook_grid(xl_path)
    except Exception as ex:
        forms.alert("Could not read the workbook:\n\n{}".format(ex), title=TITLE)
        return

    try:
        cols = locate_columns(grid, max_row, max_col)
    except Exception as ex:
        forms.alert(
            "Workbook did not match the expected 'Electrical Info' format "
            "(sheet '{}'):\n\n{}".format(sheet_name, ex),
            title=TITLE,
        )
        return

    records = parse_rows(grid, max_row, cols)
    if not records:
        forms.alert("No controller rows with electrical data were found on "
                    "sheet '{}'.".format(sheet_name), title=TITLE)
        return

    by_mark = collect_case_controllers()
    if not by_mark:
        forms.alert(
            "No case controllers with an '{}' value were found in the model."
            .format(IDENTITY_PARAM),
            title=TITLE,
        )
        return

    # ---- plan ------------------------------------------------------------
    planned = []          # (rec, elems, total_amps, ph, note)
    mismatched = []       # (rec, elem, family_v, family_p, expected_v, expected_p)
    unmatched_sheet = []  # recs with no Revit element
    va_deltas = []        # (id, computed_sum, sheet_sum)

    for rec in records:
        total_amps, ph, exp_v, exp_p, sheet_basis_va = compute_loads(rec)

        # Cross-check parsing against the sheet's own summary columns (the
        # sum of VA PH1+PH2 is independent of how loads are split across
        # phases). Compared at the sheet's literal voltages so the
        # intentional 115->120 uprate doesn't false-flag.
        sheet_sum = None
        if rec.get("va1") is not None and rec.get("va2") is not None:
            sheet_sum = rec["va1"] + rec["va2"]
        if sheet_sum is not None and ph[3] == 0 and abs(sheet_basis_va - sheet_sum) > 1.0:
            va_deltas.append((rec["id_display"], sheet_basis_va, sheet_sum))

        elems = by_mark.get(rec["id"])
        if not elems:
            unmatched_sheet.append(rec)
            continue

        ok_elems = []
        notes = []
        for elem in elems:
            fam_v = read_family_voltage(elem)
            fam_p = read_family_poles(elem)
            # Voltage mismatch = wrong numbers -> skip. A pole-count
            # difference is only noted: controllers are routinely ganged
            # across phases (e.g. three 2-pole controllers on one 3-pole
            # circuit) and designers review the phase assignments.
            if fam_v is not None and exp_v is not None and fam_v != exp_v:
                mismatched.append((rec, elem, fam_v, fam_p, exp_v, exp_p))
                continue
            ok_elems.append(elem)
            if fam_v is None:
                notes.append("voltage unverified (no voltage param)")
            if fam_p is not None and fam_p != exp_p:
                notes.append("poles differ (family {}P, sheet {}P) - review phasing"
                             .format(fam_p, exp_p))
        if ok_elems:
            planned.append((rec, ok_elems, total_amps, ph, "; ".join(notes)))

    matched_marks = set(rec["id"] for rec, _, _, _, _ in planned)
    matched_marks.update(rec["id"] for rec, _, _, _, _, _ in mismatched)
    unmatched_revit = sorted(m for m in by_mark if m not in matched_marks)

    # ---- preview ----------------------------------------------------------
    output.print_md("# {} - Preview".format(TITLE))
    output.print_md("Workbook: `{}` | sheet: `{}`".format(xl_path, sheet_name))
    output.print_md(
        "Sheet rows: **{}** | matched: **{}** | voltage mismatch: **{}** | "
        "no Revit match: **{}** | Revit CCs without a sheet row: **{}**".format(
            len(records), len(planned), len(mismatched),
            len(unmatched_sheet), len(unmatched_revit)))

    if planned:
        table = []
        for rec, elems, total_amps, ph, note in planned:
            table.append([
                rec["id_display"],
                ", ".join(output.linkify(e.Id) for e in elems),
                "{:.2f}".format(total_amps),
                _fmt_va(ph[1]), _fmt_va(ph[2]), _fmt_va(ph[3]),
                note,
            ])
        output.print_table(
            table,
            columns=["ID #", "Element(s)", "Ampacity (A)",
                     "VA Ph1", "VA Ph2", "VA Ph3", "Notes"],
            title="Will write")

    if mismatched:
        table = [[
            rec["id_display"], output.linkify(elem.Id),
            "{} V / {}P".format(fam_v or "?", fam_p or "?"),
            "{} V / {}P".format(exp_v or "?", exp_p),
        ] for rec, elem, fam_v, fam_p, exp_v, exp_p in mismatched]
        output.print_table(
            table,
            columns=["ID #", "Element", "Family volt/poles", "Sheet expects"],
            title="SKIPPED - voltage mismatch")

    if unmatched_sheet:
        output.print_table(
            [[rec["id_display"]] for rec in unmatched_sheet],
            columns=["Sheet ID # with no Revit Identity Mark match"],
            title="Unmatched sheet rows")

    if unmatched_revit:
        output.print_table(
            [[m] for m in unmatched_revit],
            columns=["Revit Identity Mark with no sheet row"],
            title="Unmatched Revit case controllers")

    if va_deltas:
        output.print_table(
            [[i, _fmt_va(a), _fmt_va(b)] for i, a, b in va_deltas],
            columns=["ID #", "Computed VA total", "Sheet VA PH1+PH2"],
            title="Cross-check deltas vs sheet (>1 VA)")

    if not planned:
        forms.alert("Nothing to write - see the preview report.", title=TITLE)
        return

    if not forms.alert(
            "Write Ampacity Rating + Apparent Load Ph 1/2/3 to {} controller "
            "ID(s)?\n\n(Mismatched / unmatched elements are skipped - see the "
            "preview report.)".format(len(planned)),
            title=TITLE, yes=True, no=True):
        return

    # ---- apply -------------------------------------------------------------
    written = 0
    failures = []
    tx = DB.Transaction(doc, TITLE)
    tx.Start()
    try:
        for rec, elems, total_amps, ph, _ in planned:
            for elem in elems:
                try:
                    targets = [(PARAM_AMPACITY, _to_internal_amps(total_amps))]
                    for n in (1, 2, 3):
                        # write VA rounded to two decimals
                        targets.append((PARAM_PH[n], _to_internal_va(round(ph[n], 2))))
                    missing = []
                    for pname, value in targets:
                        p = _lookup(elem, pname)
                        if p is None or p.IsReadOnly:
                            missing.append(pname)
                            continue
                        p.Set(value)
                    if missing:
                        failures.append((rec["id_display"], elem.Id,
                                         "missing/read-only: " + ", ".join(missing)))
                    else:
                        written += 1
                except Exception as ex:
                    failures.append((rec["id_display"], elem.Id, str(ex)))
        tx.Commit()
    except Exception as ex:
        tx.RollBack()
        forms.alert("Transaction failed and was rolled back:\n\n{}".format(ex),
                    title=TITLE)
        return

    output.print_md("---")
    output.print_md("## Results")
    output.print_md("Elements updated: **{}**".format(written))
    if failures:
        output.print_table(
            [[i, output.linkify(eid), msg] for i, eid, msg in failures],
            columns=["ID #", "Element", "Problem"],
            title="Write failures")
    forms.alert(
        "Done. {} element(s) updated{}.".format(
            written,
            ", {} failure(s) - see report".format(len(failures)) if failures else ""),
        title=TITLE)


main()
