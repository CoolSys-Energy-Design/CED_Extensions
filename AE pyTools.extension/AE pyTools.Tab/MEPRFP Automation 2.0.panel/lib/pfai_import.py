# -*- coding: utf-8 -*-
"""Inject a PFAI power-plan workbook into the active document.

The plan is generated outside Revit (E101 Power Plan plug-in) and arrives as a
workbook with four sheets - Panels, Devices, Circuits, Keynotes. This module
turns that into model elements in four transactions wrapped in one transaction
group, so the whole import is a single undo step.

Why four and not one per concern: an instance's parameters are set on the object
you just created, so placement and parameters belong together - splitting them
means re-finding every device by Mark on a second pass. The seams below are the
places a Regenerate is actually required.

    1  types + panels    activate symbols, create missing types, name panels
    2  devices           place, rotate, elevate, stamp
    3  circuits          create, select panel, rating, load name
    4  annotations       keynotes + circuit tags

Everything created is stamped Comments = "PFAI". Keynote instances have no
Comments or Mark to stamp, and their one free-text field is a column in the
Keynotes Master schedule, so their element ids are recorded in extensible
storage instead. Either way sweep_previous finds the previous import, so
re-running replaces its own work rather than stacking a second copy on it.
"""

import math

from Autodesk.Revit import DB
from Autodesk.Revit.DB import Electrical as DBE
from Autodesk.Revit.DB import ExtensibleStorage as DBES
from System import Guid
from System.Collections.Generic import List

import pfai_xlsx

__all__ = ["PfaiPlan", "ImportResult", "read_plan", "sweep_previous",
           "run_import", "KEYNOTE_FAMILY"]

STAMP = "PFAI"
# Reed, explicitly: any other keynote family is wrong. Enforced rather than
# preferred - a workbook asking for something else is refused, not honoured.
KEYNOTE_FAMILY = "GA_Keynote Symbol_CED"
KEYNOTE_TYPE = "Standard"
# `Keynote Value_CEDT` is the input; `Keynote Value on Sheet_CEDT` is what
# actually PRINTS. Nothing derives one from the other, so both are written -
# setting only the first leaves the family default `E001` on the drawing.
KEYNOTE_VALUE_PARAM = "Keynote Value_CEDT"
KEYNOTE_SHEET_PARAM = "Keynote Value on Sheet_CEDT"
KEYNOTE_DESC_PARAM = "Keynote Description_CEDT"
# Retired as a stamp: it is a column in the Keynotes Master schedule, so writing
# "PFAI" there invented a keynote category on a real schedule. Still read, to
# find and clean up imports made before this changed.
LEGACY_STAMP_PARAM = "Keynote Category_CEDT"
PANEL_NAME_PARAM = "Panel Name_CEDT"      # built-in Panel Name is read-only
TAG_FAMILY = "EF-Tag_Electrical Fixtures_CED"
TAG_TYPE = "Panel & Circuit Number"
TAG_OFFSET_FT = 1.5


class PfaiPlan(object):
    def __init__(self, panels, devices, circuits, keynotes, source=""):
        self.panels = panels
        self.devices = devices
        self.circuits = circuits
        self.keynotes = keynotes
        self.source = source

    @property
    def counts(self):
        return {"panels": len(self.panels), "devices": len(self.devices),
                "circuits": len(self.circuits), "keynotes": len(self.keynotes)}


class ImportResult(object):
    def __init__(self):
        self.placed = 0
        self.circuits_ok = 0
        self.keynotes = 0
        # of those, how many verifiably carry their number in the parameter that
        # PRINTS. Counted separately because a keynote can place perfectly and
        # still plot as the family default.
        self.keynotes_printed = 0
        self.tags = 0
        self.panels_named = 0
        self.types_created = []
        self.swept = 0
        self.problems = []        # (stage, subject, reason)

    def fail(self, stage, subject, reason):
        self.problems.append((stage, str(subject), str(reason)[:200]))


def _num(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _int(v, default=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _txt(v):
    return ("" if v is None else str(v)).strip()


def _idx_list(v):
    """"12;13;14" or "12,13,14" or a lone number -> [12, 13, 14]."""
    s = _txt(v).replace(",", ";")
    return [_int(p) for p in s.split(";") if _txt(p)]


def read_plan(path):
    book = pfai_xlsx.read_workbook(path)
    panels = []
    for r in pfai_xlsx.sheet_dicts(book, "Panels"):
        panels.append({
            "name": _txt(r.get("name")),
            "type": _txt(r.get("type")),
            "distribution_system": _txt(r.get("distribution_system")),
            "x": _num(r.get("x")), "y": _num(r.get("y")),
        })
    devices = []
    for r in pfai_xlsx.sheet_dicts(book, "Devices"):
        devices.append({
            "index": _int(r.get("index"), -1),
            "type": _txt(r.get("type")),
            "x": _num(r.get("x")), "y": _num(r.get("y")),
            "rotation_deg": _num(r.get("rotation_deg")),
            "elevation_ft": _num(r.get("elevation_ft")),
            "load_name": _txt(r.get("load_name")),
            "panel": _txt(r.get("panel")),
            "va": _int(r.get("va")),
            "room": _txt(r.get("room")),
            "confidence": _txt(r.get("confidence")),
        })
    circuits = []
    for r in pfai_xlsx.sheet_dicts(book, "Circuits"):
        circuits.append({
            "panel": _txt(r.get("panel")),
            "load_name": _txt(r.get("load_name")),
            "rating_a": _int(r.get("rating_a"), 20),
            "devices": _idx_list(r.get("device_indices")),
        })
    keynotes = []
    for r in pfai_xlsx.sheet_dicts(book, "Keynotes"):
        keynotes.append({
            "number": _txt(r.get("number")),
            "x": _num(r.get("x")), "y": _num(r.get("y")),
            "description": _txt(r.get("description")),
            "family": _txt(r.get("family")) or KEYNOTE_FAMILY,
            "type": _txt(r.get("type")) or KEYNOTE_TYPE,
        })
    return PfaiPlan(panels, devices, circuits, keynotes, source=path)


def _family_name(symbol):
    """Family name, read the way that survives both Revit Python engines.

    SYMBOL_FAMILY_NAME_PARAM reads None on a FamilySymbol here, so go through
    the Family element and fall back to its Name property.
    """
    fam = symbol.Family
    p = fam.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    if p is not None and p.AsString():
        return p.AsString()
    return getattr(fam, "Name", "") or ""


def _type_name(symbol):
    p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    return (p.AsString() if p else "") or ""


def _symbols_in_category(doc, bic):
    cid = int(bic)
    out = {}
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements():
        if s.Category is not None and s.Category.Id.IntegerValue == cid:
            out[_type_name(s)] = s
    return out


def _activate(doc, symbol):
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    return symbol


def _set(elem, builtin, value):
    p = elem.get_Parameter(builtin)
    if p is not None and not p.IsReadOnly:
        p.Set(value)
        return True
    return False


def _set_named(elem, name, value):
    p = elem.LookupParameter(name)
    if p is not None and not p.IsReadOnly:
        p.Set(value)
        return True
    return False


_SCHEMA_GUID = Guid("b7c41e92-3a6d-4f18-9c25-7e0d5a1f8b34")
_SCHEMA_NAME = "PFAI_Import_Registry_v1"
_SCHEMA_FIELD = "KeynoteIds"


def _registry_schema():
    schema = DBES.Schema.Lookup(_SCHEMA_GUID)
    if schema is not None:
        return schema
    b = DBES.SchemaBuilder(_SCHEMA_GUID)
    b.SetSchemaName(_SCHEMA_NAME)
    b.SetReadAccessLevel(DBES.AccessLevel.Public)
    b.SetWriteAccessLevel(DBES.AccessLevel.Public)
    b.AddSimpleField(_SCHEMA_FIELD, str)
    return b.Finish()


def _registry_store(doc, create=False):
    """Where the keynote registry lives.

    Project Information rather than a DataStorage element: `DataStorage` is not
    reachable from the DB namespace in every Revit Python host, and Project
    Information always exists, needs nothing created, and is namespaced by the
    schema GUID so it cannot collide with anything else stored there.
    """
    return doc.ProjectInformation, _registry_schema()


def record_keynotes(doc, ids):
    """Remember which keynotes this import placed, so the next one can clear
    them. Their family has nowhere to hide a stamp that a schedule will not
    pick up."""
    store, schema = _registry_store(doc, create=True)
    entity = DBES.Entity(schema)
    entity.Set(_SCHEMA_FIELD, ",".join(str(i.IntegerValue) for i in ids))
    store.SetEntity(entity)


def recorded_keynotes(doc):
    store, schema = _registry_store(doc)
    if store is None:
        return []
    entity = store.GetEntity(schema)
    if not entity.IsValid():
        return []
    raw = entity.Get[str](_SCHEMA_FIELD) or ""
    out = []
    for part in raw.split(","):
        part = part.strip()
        if part:
            try:
                out.append(DB.ElementId(int(part)))
            except ValueError:
                pass
    return out


def _is_ours(elem):
    p = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if p is not None and p.AsString() == STAMP:
        return True
    # keynotes imported before the stamp moved to extensible storage
    p = elem.LookupParameter(LEGACY_STAMP_PARAM)
    return p is not None and p.AsString() == STAMP


SWEEP_CATEGORIES = (
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_ElectricalFixtureTags,
    DB.BuiltInCategory.OST_GenericAnnotation,
)


def sweep_previous(doc):
    """Delete a previous PFAI import. Panels are deliberately left alone."""
    ids = List[DB.ElementId]()
    seen = set()
    for bic in SWEEP_CATEGORIES:
        for e in DB.FilteredElementCollector(doc).OfCategory(
                bic).WhereElementIsNotElementType().ToElements():
            if _is_ours(e) and e.Id.IntegerValue not in seen:
                seen.add(e.Id.IntegerValue)
                ids.Add(e.Id)
    for eid in recorded_keynotes(doc):
        if eid.IntegerValue in seen:
            continue
        if doc.GetElement(eid) is not None:
            seen.add(eid.IntegerValue)
            ids.Add(eid)
    if ids.Count:
        doc.Delete(ids)
    return ids.Count


def _pass_types_and_panels(doc, plan, result, fallback_types):
    symbols = _symbols_in_category(doc, DB.BuiltInCategory.OST_ElectricalFixtures)
    wanted = set(d["type"] for d in plan.devices if d["type"])
    for tn in sorted(wanted - set(symbols)):
        src = fallback_types.get(tn)
        base = symbols.get(src) if src else None
        if base is None:
            result.fail("types", tn,
                        "not in the document and no fallback to duplicate from")
            continue
        try:
            symbols[tn] = base.Duplicate(tn)
            result.types_created.append((tn, src))
        except Exception as exc:
            result.fail("types", tn, exc)

    by_name = {}
    for p in DB.FilteredElementCollector(doc).OfCategory(
            DB.BuiltInCategory.OST_ElectricalEquipment
    ).WhereElementIsNotElementType().ToElements():
        nm = p.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        nm = (nm.AsString() if nm else "") or ""
        if nm:
            by_name.setdefault(nm, p)

    ds_by_name = {}
    for d in DB.FilteredElementCollector(doc).OfClass(
            DBE.DistributionSysType).ToElements():
        p = d.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        nm = (p.AsString() if p else "") or getattr(d, "Name", "")
        if nm:
            ds_by_name[nm] = d

    for spec in plan.panels:
        name = spec["name"]
        if not name:
            continue
        panel = by_name.get(name)
        if panel is None:
            result.fail("panels", name,
                        "no panel with this name in the model - place it first")
            continue
        if _set_named(panel, PANEL_NAME_PARAM, name):
            result.panels_named += 1
        want_ds = spec.get("distribution_system")
        if want_ds:
            ds = ds_by_name.get(want_ds)
            if ds is None:
                result.fail("panels", name,
                            "distribution system %r is not in the document"
                            % want_ds)
            else:
                pp = panel.get_Parameter(
                    DB.BuiltInParameter.RBS_FAMILY_CONTENT_DISTRIBUTION_SYSTEM)
                if pp is not None and not pp.IsReadOnly:
                    pp.Set(ds.Id)
    return symbols


def _pass_devices(doc, plan, level, symbols, result):
    placed = {}
    for d in plan.devices:
        tn = d["type"]
        sym = symbols.get(tn)
        if sym is None:
            result.fail("devices", "#%s %s" % (d["index"], tn), "unknown type")
            continue
        try:
            _activate(doc, sym)
            pt = DB.XYZ(d["x"], d["y"], 0.0)
            fi = doc.Create.NewFamilyInstance(
                pt, sym, level, DB.Structure.StructuralType.NonStructural)
            rot = d["rotation_deg"]
            if abs(rot) > 1e-6:
                axis = DB.Line.CreateBound(pt, DB.XYZ(d["x"], d["y"], 10.0))
                DB.ElementTransformUtils.RotateElement(
                    doc, fi.Id, axis, math.radians(rot))
            _set(fi, DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM,
                 d["elevation_ft"])
            _set(fi, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, STAMP)
            _set(fi, DB.BuiltInParameter.ALL_MODEL_MARK,
                 "PFAI-%03d|%s|%s" % (d["index"], d["panel"], d["load_name"]))
            placed[d["index"]] = fi
            result.placed += 1
        except Exception as exc:
            result.fail("devices", "#%s %s" % (d["index"], tn), exc)
    return placed


def _pass_circuits(doc, plan, placed, result):
    panels = {}
    for p in DB.FilteredElementCollector(doc).OfCategory(
            DB.BuiltInCategory.OST_ElectricalEquipment
    ).WhereElementIsNotElementType().ToElements():
        nm = p.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_NAME)
        nm = (nm.AsString() if nm else "") or ""
        if nm:
            panels.setdefault(nm, p)

    for c in plan.circuits:
        label = "%s / %s" % (c["panel"], c["load_name"])
        missing = [i for i in c["devices"] if i not in placed]
        if missing:
            result.fail("circuits", label, "devices not placed: %s" % missing)
            continue
        panel = panels.get(c["panel"])
        if panel is None:
            result.fail("circuits", label, "no panel named %r" % c["panel"])
            continue
        try:
            ids = List[DB.ElementId]()
            for i in c["devices"]:
                ids.Add(placed[i].Id)
            system = DBE.ElectricalSystem.Create(
                doc, ids, DBE.ElectricalSystemType.PowerCircuit)
        except Exception as exc:
            result.fail("circuits", label, "could not create: %s" % exc)
            continue
        try:
            system.SelectPanel(panel)
        except Exception as exc:
            # The two real causes, both design problems rather than import bugs:
            # the panel is out of slots, or the load's connector voltage and
            # poles do not match the panel's distribution system.
            result.fail("circuits", label, exc)
            continue
        _set(system, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_RATING_PARAM,
             float(c["rating_a"]))
        _set(system, DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NAME, c["load_name"])
        _set(system, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, STAMP)
        result.circuits_ok += 1


def _keynote_symbol(doc, result):
    best = None
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements():
        if _family_name(s) != KEYNOTE_FAMILY:
            continue
        if _type_name(s) == KEYNOTE_TYPE:
            return s
        best = best or s
    if best is None:
        result.fail("keynotes", KEYNOTE_FAMILY,
                    "family is not loaded in this document")
    return best


def _tag_symbol(doc):
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements():
        if _family_name(s) == TAG_FAMILY and _type_name(s) == TAG_TYPE:
            return s
    return None


def _pass_annotations(doc, view, plan, placed, result):
    placed_keynotes = []
    kn_printed = 0
    sym = _keynote_symbol(doc, result)
    if sym is not None:
        _activate(doc, sym)
        for k in plan.keynotes:
            asked = k.get("family") or KEYNOTE_FAMILY
            if asked != KEYNOTE_FAMILY:
                result.fail("keynotes", "#%s" % k["number"],
                            "workbook asked for %r; only %s is allowed"
                            % (asked, KEYNOTE_FAMILY))
                continue
            num = str(k["number"])
            if not num:
                result.fail("keynotes", "(blank)",
                            "the workbook row has no number, so it would print "
                            "the family default - skipped")
                continue
            try:
                fi = doc.Create.NewFamilyInstance(
                    DB.XYZ(k["x"], k["y"], 0.0), sym, view)
                _set_named(fi, KEYNOTE_VALUE_PARAM, num)
                wrote = _set_named(fi, KEYNOTE_SHEET_PARAM, num)
                if k.get("description"):
                    _set_named(fi, KEYNOTE_DESC_PARAM, k["description"])
                # `Keynote Value on Sheet_CEDT` is the one that PRINTS. Read it
                # back rather than trusting the write: if it is missing, renamed
                # or read-only in some other template, the silent result is the
                # family default `E001` on every keynote on the sheet - which is
                # invisible here and glaring on a plot.
                shown = fi.LookupParameter(KEYNOTE_SHEET_PARAM)
                shown = (shown.AsString() if shown else None) or ""
                if not wrote or shown != num:
                    result.fail("keynotes", "#%s" % num,
                                "the printing value could not be set - %s reads "
                                "%r, so this keynote will plot as the family "
                                "default instead of %s"
                                % (KEYNOTE_SHEET_PARAM, shown or "<empty>", num))
                else:
                    kn_printed += 1
                placed_keynotes.append(fi.Id)
                result.keynotes += 1
            except Exception as exc:
                result.fail("keynotes", "#%s" % num, exc)

    result.keynotes_printed = kn_printed

    if placed_keynotes:
        try:
            record_keynotes(doc, placed_keynotes)
        except Exception as exc:
            result.fail("keynotes", "registry",
                        "placed, but could not record ids for the next sweep "
                        "to clear: %s" % exc)

    tag = _tag_symbol(doc)
    if tag is None:
        result.fail("tags", TAG_TYPE, "tag family is not loaded - tags skipped")
        return
    _activate(doc, tag)
    for idx in sorted(placed):
        elem = placed[idx]
        num = elem.get_Parameter(DB.BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER)
        if not ((num.AsString() if num else "") or "").strip():
            continue
        try:
            pt = elem.Location.Point
            DB.IndependentTag.Create(
                doc, tag.Id, view.Id, DB.Reference(elem), False,
                DB.TagOrientation.Horizontal,
                DB.XYZ(pt.X, pt.Y + TAG_OFFSET_FT, 0.0))
            result.tags += 1
        except Exception as exc:
            result.fail("tags", "device #%s" % idx, exc)


def run_import(doc, view, plan, sweep=True, fallback_types=None):
    """Four transactions, one undo step. Returns an ImportResult."""
    result = ImportResult()
    level = getattr(view, "GenLevel", None)
    if level is None:
        result.fail("setup", getattr(view, "Name", "active view"),
                    "this view has no level to place on - open the power plan")
        return result

    # Do NOT call doc.Regenerate() between these transactions: it needs an
    # open transaction of its own, and Commit() has already regenerated. Inside
    # a transaction (see _activate) it is both legal and necessary.
    group = DB.TransactionGroup(doc, "Import PFAI Design")
    group.Start()
    try:
        if sweep:
            t = DB.Transaction(doc, "PFAI: clear previous import")
            t.Start()
            result.swept = sweep_previous(doc)
            t.Commit()

        t = DB.Transaction(doc, "PFAI: types and panels")
        t.Start()
        symbols = _pass_types_and_panels(doc, plan, result, fallback_types or {})
        t.Commit()

        t = DB.Transaction(doc, "PFAI: place devices")
        t.Start()
        placed = _pass_devices(doc, plan, level, symbols, result)
        t.Commit()

        t = DB.Transaction(doc, "PFAI: circuits")
        t.Start()
        _pass_circuits(doc, plan, placed, result)
        t.Commit()

        t = DB.Transaction(doc, "PFAI: keynotes and tags")
        t.Start()
        _pass_annotations(doc, view, plan, placed, result)
        t.Commit()

        group.Assimilate()
    except Exception:
        group.RollBack()
        raise
    return result
