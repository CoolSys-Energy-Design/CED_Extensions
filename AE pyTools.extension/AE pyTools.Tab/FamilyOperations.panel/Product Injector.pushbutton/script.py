# -*- coding: utf-8 -*-
# ASCII-ONLY on purpose: some IronPython 2.7 setups fail to compile source with
# stray non-ASCII characters, which shows up as a blank pyRevit window.
"""Product Injector  (PROTOTYPE - scoped to Reznor SC Duct Furnace).

Flow:
  1. Pick MasterProductList.xlsx.
  2. Pick a Product Type  (= a sheet, e.g. "Reznor SC Duct Furnace").
  3. Pick a Model         (= a row, by its "Model" column, e.g. "SC-150").
  4. Pick an EXISTING loaded family to add this model to.
  5. Add a new type named after the Model and inject the row's specs as TYPE
     parameters (creating any missing family parameters).
  6. Reload into the project, activate the type, enter click-to-place mode.

Adding the type to an existing family means the placed instance inherits that
family's geometry and category - it is visible, and there is no template /
geometry / category handling to get wrong. The type carries the parameters, so
every instance the user drops is pre-filled.
"""
from __future__ import print_function

import os
import re
import sys

from pyrevit import revit, DB, HOST_APP
from pyrevit import forms
from pyrevit import script

logger = script.get_logger()
output = script.get_output()

FAMILY_SUFFIX = "_CED"
INDEX_SHEET = "_Index"
DEFAULT_XLSX_DIR = r"C:\Users\reed.pinterich\OneDrive - CoolSys Inc\Desktop\AutomateFamilyBuildingPerCutsheet"

# Globals populated by bootstrap() (kept out of module load so nothing here can
# blank the button before the error handler at the bottom is in effect).
doc = None
uidoc = None
app = None
NEW_API = None
GRP_GEOMETRY = GRP_MECH = GRP_IDENT = None
SPEC_TEXT = SPEC_NUMBER = SPEC_LENGTH = None
# Revit stores length internally in decimal FEET regardless of display units.
INCHES_TO_INTERNAL = 1.0 / 12.0


def _first_attr(owner, names, fallback=None):
    for n in names:
        obj = owner
        ok = True
        for part in n.split("."):
            try:
                obj = getattr(obj, part)
            except Exception:
                ok = False
                break
        if ok:
            return obj
    return fallback


def bootstrap():
    """Resolve the active document and API tokens. Raises with a clear message."""
    global doc, uidoc, app, NEW_API
    global GRP_GEOMETRY, GRP_MECH, GRP_IDENT, SPEC_TEXT, SPEC_NUMBER, SPEC_LENGTH

    uidoc = revit.uidoc
    doc = revit.doc
    if doc is None:
        raise Exception("No active Revit document. Open a project and try again.")
    app = doc.Application

    NEW_API = HOST_APP.is_newer_than(2022)
    if NEW_API:
        g = DB.GroupTypeId
        GRP_IDENT = _first_attr(g, ["IdentityData"])
        GRP_GEOMETRY = _first_attr(g, ["Geometry", "Dimensions"], GRP_IDENT)
        GRP_MECH = _first_attr(g, ["Mechanical"], GRP_IDENT)
        SPEC_TEXT = _first_attr(DB.SpecTypeId, ["String.Text"])
        SPEC_NUMBER = _first_attr(DB.SpecTypeId, ["Number"])
        SPEC_LENGTH = _first_attr(DB.SpecTypeId, ["Length"])
    else:
        GRP_GEOMETRY = DB.BuiltInParameterGroup.PG_GEOMETRY
        GRP_MECH = DB.BuiltInParameterGroup.PG_MECHANICAL
        GRP_IDENT = DB.BuiltInParameterGroup.PG_IDENTITY_DATA
        SPEC_TEXT = DB.ParameterType.Text
        SPEC_NUMBER = DB.ParameterType.Number
        SPEC_LENGTH = DB.ParameterType.Length


# =============================================================================
# Column -> parameter mapping   (PROTOTYPE: tuned for Reznor SC Duct Furnace)
# =============================================================================
SKIP_COLS = {"Model"}  # becomes the family TYPE name, not a parameter
NUMERIC_COLS = {
    "Nominal Input (MBh)",
    "CFM @30F rise", "P.D. @30F (in wc)",
    "CFM @40F rise", "P.D. @40F (in wc)",
    "CFM @50F rise", "P.D. @50F (in wc)",
    "Max Static (in wc)",
}
DERIVED_NUMERIC_COLS = {"Nominal Output (MBh) [derived @80%]"}
DIM_LENGTH_COLS = {
    "Dim A (in)", "Dim B (in)", "Dim C (in)", "Dim D (in)", "Dim E (in)",
    "Dim G duct conn (in)",
}
LOWCONF_TEXT_COLS = {"Thermal Efficiency"}


def _clean_param_name(name):
    """Strip characters Revit disallows in parameter names ( [ ] { } | ; < > ? ~ ` \\ ),
    keeping legal ones like ( ) @ % / . -. Trim the result."""
    bad = "[]{}|;<>?~`\\"
    return "".join(ch for ch in name if ch not in bad).strip()


def build_param_plan(headers):
    """Return (plan, flagged, skipped)."""
    plan, flagged, skipped = [], [], []
    for col in headers:
        if not col:
            continue
        if col in SKIP_COLS:
            skipped.append((col, "used as the family Type name, not a parameter"))
            continue

        if col in NUMERIC_COLS:
            spec, group, kind = SPEC_NUMBER, GRP_MECH, "number"
        elif col in DERIVED_NUMERIC_COLS:
            spec, group, kind = SPEC_NUMBER, GRP_MECH, "number"
            flagged.append((col, "DERIVED @80% (not a nameplate value) - confirm before use"))
        elif col in DIM_LENGTH_COLS:
            spec, group, kind = SPEC_LENGTH, GRP_GEOMETRY, "length"
        elif col in LOWCONF_TEXT_COLS:
            spec, group, kind = SPEC_TEXT, GRP_IDENT, "text"
            flagged.append((col, "contains a unit/symbol (e.g. %); stored as TEXT"))
        else:
            spec, group, kind = SPEC_TEXT, GRP_IDENT, "text"

        plan.append({"col": col, "param": _clean_param_name(col),
                     "spec": spec, "group": group, "kind": kind})
    return plan, flagged, skipped


# =============================================================================
# Excel helpers
# =============================================================================
def pick_workbook():
    init_dir = DEFAULT_XLSX_DIR if os.path.isdir(DEFAULT_XLSX_DIR) else None
    return forms.pick_file(file_ext="xlsx", title="Select MasterProductList.xlsx",
                           init_dir=init_dir)


def load_sheets(xlsx_path):
    """Return {sheet_name: {'headers': [...], 'rows': [ {col: val}, ... ]}}."""
    from pyrevit.interop import xl as pyxl  # lazy import
    data = pyxl.load(xlsx_path, headers=False)
    result = {}
    for sheet_name, payload in data.items():
        rows = payload.get("rows") or []
        if not rows:
            continue
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        recs = []
        for raw in rows[1:]:
            if raw is None:
                continue
            recs.append(dict(zip(headers, raw)))
        result[sheet_name] = {"headers": headers, "rows": recs}
    return result


# =============================================================================
# Category selection
# =============================================================================
def category_options():
    # (label, BuiltInCategory member name). Resolved defensively so a category
    # missing in this Revit version is skipped rather than crashing the tool.
    wanted = [
        ("Mechanical Equipment", "OST_MechanicalEquipment"),
        ("Electrical Equipment", "OST_ElectricalEquipment"),
        ("Electrical Fixtures", "OST_ElectricalFixtures"),
        ("Specialty Equipment", "OST_SpecialityEquipment"),  # API spelling variant
        ("Specialty Equipment", "OST_SpecialtyEquipment"),
        ("Plumbing Fixtures", "OST_PlumbingFixtures"),
        ("Plumbing Equipment", "OST_PlumbingEquipment"),
        ("Air Terminals", "OST_DuctTerminal"),
        ("Generic Models", "OST_GenericModel"),
    ]
    opts = {}
    for label, member in wanted:
        cat = getattr(DB.BuiltInCategory, member, None)
        if cat is not None and label not in opts:
            opts[label] = cat
    return opts


def pick_category():
    opts = category_options()
    choice = forms.SelectFromList.show(
        sorted(opts.keys()), title="Create family in which Category?",
        button_name="Use Category", multiselect=False)
    if not choice:
        return None, None
    return choice, opts[choice]


def _fam_cat_name(fam):
    try:
        return fam.FamilyCategory.Name
    except Exception:
        return "?"


def pick_existing_family():
    """Let the user pick a loaded, editable MODEL family to add a type to.

    Adding the type to an existing family means the placed instance inherits
    that family's geometry and category - so it is visible and needs no
    template/geometry/category handling.
    """
    fams = []
    for f in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        try:
            if not f.IsEditable:
                continue
            fc = f.FamilyCategory
            if fc is None or fc.CategoryType != DB.CategoryType.Model:
                continue
            fams.append((f.Name, f))
        except Exception:
            continue
    if not fams:
        forms.alert("No editable model families found in this project to add a type to.",
                    warn_icon=True)
        return None
    fams.sort(key=lambda nf: nf[0].lower())
    label_map = {}
    labels = []
    for name, f in fams:
        label = "{}   [{}]".format(name, _fam_cat_name(f))
        label_map[label] = f
        labels.append(label)
    choice = forms.SelectFromList.show(
        labels, title="Add the type to which existing family?",
        button_name="Use Family", multiselect=False)
    if not choice:
        return None
    return label_map.get(choice)


# =============================================================================
# Family building
# =============================================================================
class _OverwriteLoader(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = True
        return True


CATEGORY_TEMPLATE_HINTS = {
    "Mechanical Equipment": ["mechanical equipment"],
    "Electrical Equipment": ["electrical equipment"],
    "Electrical Fixtures": ["electrical fixture"],
    "Specialty Equipment": ["specialty equipment", "speciality equipment"],
    "Plumbing Fixtures": ["plumbing fixture"],
    "Plumbing Equipment": ["plumbing equipment"],
    "Air Terminals": ["air terminal"],
    "Generic Models": ["generic model"],
}
# Exclude template variants that don't allow a plain solid extrusion and/or
# can't be re-categorized (adaptive, hosted/"based", annotation, mass, etc.).
_TEMPLATE_EXCLUDE = ["adaptive", "based", "two level", "curtain", "pattern",
                     "annotation", "title", "profile", "detail", "mass",
                     "baluster", "structural", "railing", "rebar", "hosted"]


def find_template(cat_label):
    """Pick the plainest Imperial .rft that matches the chosen category so the
    family is that category NATIVELY (no reassignment) and allows geometry."""
    base = None
    try:
        base = app.FamilyTemplatePath
    except Exception:
        base = None
    rfts = []
    if base and os.path.isdir(base):
        for root, _dirs, files in os.walk(base):
            for f in files:
                if f.lower().endswith(".rft"):
                    rfts.append(os.path.join(root, f))

    hints = list(CATEGORY_TEMPLATE_HINTS.get(cat_label, [])) + ["generic model"]

    def _match(sub, allow_metric):
        out = []
        for p in rfts:
            name = os.path.basename(p).lower()
            if sub not in name:
                continue
            if not allow_metric and "metric" in name:
                continue
            if any(bad in name for bad in _TEMPLATE_EXCLUDE):
                continue
            out.append(p)
        out.sort(key=lambda p: len(os.path.basename(p)))  # plainest name first
        return out

    for allow_metric in (False, True):
        for sub in hints:
            hit = _match(sub, allow_metric)
            if hit:
                return hit[0]
    return forms.pick_file(file_ext="rft",
                           title="Pick a family template (.rft) for {}".format(cat_label))


def find_family_by_name(project_doc, name):
    for f in DB.FilteredElementCollector(project_doc).OfClass(DB.Family):
        try:
            if f.Name.strip() == name.strip():
                return f
        except Exception:
            pass
    return None


def add_placeholder_box(fdoc):
    """Best-effort visible box so the family isn't an invisible point."""
    try:
        w, d, h = 2.0, 1.5, 1.0  # feet
        pts = [DB.XYZ(-w / 2, -d / 2, 0), DB.XYZ(w / 2, -d / 2, 0),
               DB.XYZ(w / 2, d / 2, 0), DB.XYZ(-w / 2, d / 2, 0)]
        ca = DB.CurveArray()
        for i in range(4):
            ca.Append(DB.Line.CreateBound(pts[i], pts[(i + 1) % 4]))
        caa = DB.CurveArrArray()
        caa.Append(ca)
        plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ.Zero)
        sp = DB.SketchPlane.Create(fdoc, plane)
        fdoc.FamilyCreate.NewExtrusion(True, caa, sp, h)
        return True
    except Exception as ex:
        logger.warning("Placeholder geometry skipped: %s", ex)
        return False


def parse_inches(raw):
    """Parse an architectural inch string to decimal inches.

    Handles '30-23/32', '30 23/32', '12-1/2', '45-1/2', '40', '23/32', '32.25'.
    Whole/fraction separator may be a hyphen or a space. Returns float or None.
    """
    if raw is None:
        return None
    s = str(raw).strip().replace('"', '')  # strip inch marks
    if not s:
        return None
    s = s.replace('-', ' ')  # whole-fraction separator -> space
    total = 0.0
    try:
        for part in s.split():
            if '/' in part:
                num, den = part.split('/')
                total += float(num) / float(den)
            else:
                total += float(part)
        return total
    except Exception:
        return None


def coerce_value(kind, raw):
    if raw is None:
        return None
    if kind == "number":
        try:
            return float(raw)
        except Exception:
            return None
    if kind == "length":
        return parse_inches(raw)  # decimal inches
    text = str(raw).strip()
    return text if text else None


# =============================================================================
# Column -> existing family parameter matching (never creates parameters)
# =============================================================================
# Groups of names that mean the same thing. The first entry is the canonical
# token; any member on the sheet OR on the family collapses to it, so e.g. an
# "MCA" column matches a "Minimum Circuit Ampacity" family parameter.
SYNONYM_GROUPS = [
    ["mca", "minimum circuit ampacity", "min circuit ampacity"],
    ["mocp", "maximum overcurrent protection", "max overcurrent protection",
     "maximum overcurrent protective device"],
    ["maximum fuse", "max fuse", "maximum fuse size", "max fuse size", "fuse size"],
    ["fla", "full load amps", "motor full load amps"],
    ["rla", "rated load amps"],
    ["lra", "locked rotor amps"],
    ["airflow", "cfm", "air flow", "design airflow"],
    ["voltage", "volts"],
    ["input", "nominal input", "heating input"],
    ["output", "nominal output", "heating output"],
]
_SYN_LOOKUP = {}
for _grp in SYNONYM_GROUPS:
    for _s in _grp:
        _SYN_LOOKUP[_s] = _grp[0]


def _normalize(name):
    """Lowercase, drop parenthetical/bracket unit hints and punctuation."""
    s = (name or "").lower()
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"\[[^\]]*\]", " ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())


def _canonical(normname):
    return _SYN_LOOKUP.get(normname, normname)


def build_instance_param_index(inst):
    """Index a PLACED INSTANCE's writable parameters for exact/normalized/alias
    matching. Read-only params (incl. type params seen via the instance) are
    skipped, so we only ever match instance-level values. First one wins."""
    exact, norm, canon = {}, {}, {}
    for p in inst.Parameters:
        try:
            if p.IsReadOnly:
                continue
            nm = p.Definition.Name
        except Exception:
            continue
        exact.setdefault(nm.strip().lower(), p)
        n = _normalize(nm)
        if n:
            norm.setdefault(n, p)
            canon.setdefault(_canonical(n), p)
    return exact, norm, canon


def match_family_param(col_name, exact, norm, canon):
    """Return (fparam, method) for the obviously-corresponding family parameter,
    else (None, None). method is 'exact' | 'normalized' | 'alias'."""
    lc = col_name.strip().lower()
    if lc in exact:
        return exact[lc], "exact"
    n = _normalize(col_name)
    if n and n in norm:
        return norm[n], "normalized"
    c = _canonical(n)
    if c and c in canon:
        return canon[c], "alias"
    return None, None


def set_param_object(param, raw):
    """Set a writable Parameter on a PLACED INSTANCE, respecting its storage
    type. Read-only params (which include type params reached via an instance)
    are refused, so this only ever changes the single instance. Returns
    (ok, display)."""
    if raw is None or param is None:
        return False, None
    s = str(raw).strip()
    if not s:
        return False, None
    if param.IsReadOnly:
        return False, None
    st = param.StorageType
    try:
        if st == DB.StorageType.String:
            param.Set(s)
            return True, s
        if st == DB.StorageType.Integer:
            param.Set(int(round(float(s))))
            return True, s
        if st == DB.StorageType.Double:
            # Prefer a unit-aware parse. Normalize architectural fractions
            # "36-7/32" -> "36 7/32\"" so Revit reads them as inches.
            norm = s
            if "/" in s:
                norm = s.replace("-", " ")
                if '"' not in norm and "'" not in norm:
                    norm = norm + '"'
            try:
                param.SetValueString(norm)
                return True, s
            except Exception:
                param.Set(float(s))  # last resort: raw internal double
                return True, s
        return False, None  # ElementId / other - skip
    except Exception as ex:
        logger.error("Set failed for '%s': %s", param.Definition.Name, ex)
        return False, None


def _symbol_name(sym):
    try:
        nm = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if nm:
            return nm
    except Exception:
        pass
    return getattr(sym, "Name", "?")


def pick_family_symbol(family):
    """Pick which existing TYPE of the family to place (auto if only one)."""
    symbols = []
    for sid in family.GetFamilySymbolIds():
        s = doc.GetElement(sid)
        if s is not None:
            symbols.append(s)
    if not symbols:
        forms.alert("Family '{}' has no types to place.".format(family.Name), warn_icon=True)
        return None
    if len(symbols) == 1:
        return symbols[0]
    name_map = {}
    for s in symbols:
        name_map[_symbol_name(s)] = s
    choice = forms.SelectFromList.show(
        sorted(name_map.keys()),
        title="Which type of '{}' to place?".format(family.Name),
        button_name="Use Type", multiselect=False)
    if not choice:
        return None
    return name_map.get(choice)


def pick_family_in_category(bic):
    """Pick an editable family whose category matches the chosen BuiltInCategory."""
    target_id = None
    try:
        target_id = doc.Settings.Categories.get_Item(bic).Id
    except Exception:
        target_id = None
    fams = []
    for f in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        try:
            if not f.IsEditable:
                continue
            fc = f.FamilyCategory
            if fc is None:
                continue
            if target_id is not None and fc.Id != target_id:
                continue
            fams.append((f.Name, f))
        except Exception:
            continue
    if not fams:
        forms.alert("No editable families found in that category.", warn_icon=True)
        return None
    fams.sort(key=lambda nf: nf[0].lower())
    name_map = {}
    for nm, f in fams:
        name_map.setdefault(nm, f)
    choice = forms.SelectFromList.show(
        [nm for nm, _ in fams], title="Pick a Family",
        button_name="Use Family", multiselect=False)
    if not choice:
        return None
    return name_map.get(choice)


def get_or_create_type(family, type_name):
    """Return (symbol, created). Reuse a type named type_name if it exists on the
    family, otherwise duplicate an existing type to create it. Duplicating a type
    is additive - existing types and their instances are not affected."""
    ids = list(family.GetFamilySymbolIds())
    for sid in ids:
        s = doc.GetElement(sid)
        if s and _symbol_name(s).strip().lower() == type_name.strip().lower():
            return s, False
    if not ids:
        return None, False
    base = doc.GetElement(ids[0])
    new_sym = None
    with revit.Transaction("Create Type '{}'".format(type_name), doc):
        try:
            res = base.Duplicate(type_name)
            new_sym = doc.GetElement(res) if isinstance(res, DB.ElementId) else res
        except Exception as ex:
            logger.error("Duplicate type failed: %s", ex)
    return new_sym, True


def instance_ids_of_symbol(symbol):
    out = set()
    for fi in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance):
        try:
            if fi.Symbol and fi.Symbol.Id == symbol.Id:
                out.add(fi.Id.IntegerValue)
        except Exception:
            pass
    return out


def new_instances_of_symbol(symbol, before_ids):
    out = []
    for fi in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance):
        try:
            if fi.Symbol and fi.Symbol.Id == symbol.Id and fi.Id.IntegerValue not in before_ids:
                out.append(fi)
        except Exception:
            pass
    return out


def _spec_for_kind(kind):
    if kind == "number":
        return SPEC_NUMBER
    if kind == "length":
        return SPEC_LENGTH
    return SPEC_TEXT


def _group_for_kind(kind):
    if kind == "number":
        return GRP_MECH
    if kind == "length":
        return GRP_GEOMETRY
    return GRP_IDENT


def _shared_param_file():
    """Path to a private (temp) shared-parameter file used to define the new
    project parameters. Created with a valid empty header if missing."""
    tmp = os.path.join(os.environ.get("TEMP") or os.path.expanduser("~"),
                       "ProductInjector_SharedParams.txt")
    if not os.path.exists(tmp) or os.path.getsize(tmp) == 0:
        with open(tmp, "w") as fh:
            fh.write("# This is a Revit shared parameter file.\n")
            fh.write("# Do not edit manually.\n")
            fh.write("*META\tVERSION\tMINVERSION\n")
            fh.write("META\t2\t1\n")
            fh.write("*GROUP\tID\tNAME\n")
            fh.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP"
                     "\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n")
    return tmp


def _insert_binding(bindings, definition, binding, kind):
    """Insert a binding, tolerating the 2025+ group-token change."""
    for group in (_group_for_kind(kind), DB.BuiltInParameterGroup.PG_DATA
                  if not NEW_API else _group_for_kind(kind)):
        try:
            bindings.Insert(definition, binding, group)
            return True
        except Exception:
            continue
    try:
        bindings.Insert(definition, binding, GRP_IDENT)
        return True
    except Exception:
        return False


def create_and_bind_instance_params(category, items):
    """Create INSTANCE-bound PROJECT parameters (never family parameters) for the
    given items. Must run inside a transaction. Returns set of created names
    (lowercased)."""
    created = set()
    if not items or category is None:
        return created
    prior = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = _shared_param_file()
        spf = app.OpenSharedParameterFile()
        if spf is None:
            return created
        grp = None
        for g in spf.Groups:
            if g.Name == "ProductInjector":
                grp = g
                break
        if grp is None:
            grp = spf.Groups.Create("ProductInjector")
        bindings = doc.ParameterBindings
        cats = DB.CategorySet()
        cats.Insert(category)
        for it in items:
            name = it["param"]
            try:
                definition = None
                for d in grp.Definitions:
                    if d.Name == name:
                        definition = d
                        break
                if definition is None:
                    opts = DB.ExternalDefinitionCreationOptions(name, _spec_for_kind(it["kind"]))
                    definition = grp.Definitions.Create(opts)
                if not bindings.Contains(definition):
                    _insert_binding(bindings, definition, DB.InstanceBinding(cats), it["kind"])
                created.add(name.strip().lower())
            except Exception as ex:
                logger.error("Create/bind '%s' failed: %s", name, ex)
    finally:
        if prior:
            app.SharedParametersFilename = prior
    return created


def inject_all(placed, symbol, plan, row, category):
    """Three tiers, family never edited:
      1. column matches a writable INSTANCE param -> set on each placed instance
      2. else matches a writable TYPE param       -> set once on the type
      3. else create an INSTANCE project param     -> set on each placed instance
    Returns a report dict."""
    inst0 = placed[0]
    i_ex, i_no, i_ca = build_instance_param_index(inst0)   # writable instance params
    t_ex, t_no, t_ca = build_instance_param_index(symbol)  # writable type params

    set_instance, set_type, to_create = [], [], []
    for item in plan:
        p, method = match_family_param(item["param"], i_ex, i_no, i_ca)
        if p is not None:
            set_instance.append((item, p.Definition.Name, method))
            continue
        p, method = match_family_param(item["param"], t_ex, t_no, t_ca)
        if p is not None:
            set_type.append((item, p.Definition.Name, method))
            continue
        to_create.append(item)

    report = {"instance": [], "type": [], "created": [], "unsettable": []}

    # 1) type-level matches -> set once on the symbol
    for item, pname, method in set_type:
        ok, disp = set_param_object(symbol.LookupParameter(pname), row.get(item["col"]))
        if ok:
            report["type"].append((item["col"], pname, method, disp))
        else:
            report["unsettable"].append((item["col"], "type param '{}' not set".format(pname)))

    # 2) create instance project params for the unmatched, then let Revit apply
    created = create_and_bind_instance_params(category, to_create)
    doc.Regenerate()

    # 3) instance-level writes (matched-instance + newly-created) on every placed
    for inst in placed:
        for item, pname, method in set_instance:
            set_param_object(inst.LookupParameter(pname), row.get(item["col"]))
        for item in to_create:
            if item["param"].strip().lower() in created:
                set_param_object(inst.LookupParameter(item["param"]), row.get(item["col"]))

    for item, pname, method in set_instance:
        report["instance"].append((item["col"], pname, method, str(row.get(item["col"]))))
    for item in to_create:
        if item["param"].strip().lower() in created:
            report["created"].append((item["col"], item["param"], str(row.get(item["col"]))))
        else:
            report["unsettable"].append((item["col"], "could not create instance parameter"))
    return report


# =============================================================================
# Main
# =============================================================================
def main():
    output.close_others()
    output.print_md("# Product Injector")
    output.print_md("_Prototype - Reznor SC Duct Furnace._  Engine: Python {}"
                    .format(sys.version.split()[0]))
    output.print_md("**Step 1/6:** select the MasterProductList workbook...")

    xlsx_path = pick_workbook()
    if not xlsx_path:
        output.print_md("Cancelled - no workbook selected.")
        return
    output.print_md("- Workbook: `{}`".format(xlsx_path))
    if not os.path.isfile(xlsx_path):
        forms.alert("File not found:\n{}".format(xlsx_path), warn_icon=True)
        return

    try:
        sheets = load_sheets(xlsx_path)
    except Exception as ex:
        forms.alert("Could not read workbook:\n{}".format(ex), warn_icon=True)
        return

    product_types = sorted([s for s in sheets.keys() if s != INDEX_SHEET])
    if not product_types:
        forms.alert("No product sheets found in the workbook.", warn_icon=True)
        return

    output.print_md("**Step 2/6:** pick a Product Type ({} sheets found).".format(len(product_types)))
    product_type = forms.SelectFromList.show(
        product_types, title="Select a Product Type (sheet)",
        button_name="Use Product Type", multiselect=False)
    if not product_type:
        output.print_md("Cancelled - no product type selected.")
        return

    sheet = sheets[product_type]
    headers = sheet["headers"]
    if "Model" not in headers:
        forms.alert("Sheet '{}' has no 'Model' column.".format(product_type), warn_icon=True)
        return

    output.print_md("**Step 3/6:** pick a Model (its specs get injected per instance).")
    models = [str(r.get("Model")).strip() for r in sheet["rows"] if r.get("Model")]
    model = forms.SelectFromList.show(
        models, title="Select a Model - {}".format(product_type),
        button_name="Use Model", multiselect=False)
    if not model:
        return

    row = next((r for r in sheet["rows"] if str(r.get("Model")).strip() == model), None)
    if not row:
        forms.alert("Could not find row for model '{}'.".format(model), warn_icon=True)
        return

    plan, flagged, skipped = build_param_plan(headers)

    # --- Pick Category + Family; create a per-SHEET type to place -----------
    output.print_md("**Step 4/6:** pick the Category.")
    cat_name, bic = pick_category()
    if not bic:
        output.print_md("Cancelled - no category selected.")
        return

    output.print_md("**Step 5/6:** pick the Family in {}.".format(cat_name))
    target_family = pick_family_in_category(bic)
    if not target_family:
        output.print_md("Cancelled - no family selected.")
        return
    family_name = target_family.Name

    # The new type is named after the SHEET (product type) - the product line.
    type_name = product_type
    symbol, created = get_or_create_type(target_family, type_name)
    if symbol is None:
        forms.alert("Family '{}' has no types to duplicate.".format(family_name), warn_icon=True)
        return
    output.print_md("- Type `{}` {} in family `{}`."
                    .format(type_name, "created" if created else "reused", family_name))

    with revit.Transaction("Activate Type", doc):
        if not symbol.IsActive:
            symbol.Activate()
        doc.Regenerate()

    # --- Interactive placement, then write values onto the placed instances --
    before = instance_ids_of_symbol(symbol)
    forms.alert(
        "Click to place one or more '{}' instances, then press Esc / Modify to "
        "finish.\n\nThe values for {} will be written to the instances you place. "
        "Existing family/type params get the value; anything with no match gets a "
        "new INSTANCE parameter. The family definition is never edited."
        .format(type_name, model),
        title="Product Injector")
    output.print_md("**Step 6/6:** placing - click in the model, Esc to finish...")
    try:
        uidoc.PromptForFamilyInstancePlacement(symbol)
    except Exception:
        pass  # normal finish or user cancel raises here

    placed = new_instances_of_symbol(symbol, before)
    if not placed:
        output.print_md("No instances placed - nothing injected.")
        forms.alert("No instances were placed, so nothing was injected.",
                    title="Product Injector")
        return

    category = None
    try:
        category = doc.Settings.Categories.get_Item(bic)
    except Exception:
        category = None

    with revit.Transaction("Inject Product Values", doc):
        report = inject_all(placed, symbol, plan, row, category)

    # --- Report -------------------------------------------------------------
    n_total = len(report["instance"]) + len(report["type"]) + len(report["created"])
    output.print_md("---")
    output.print_md("**Placed {} instance(s)** of `{} : {}`   (Category: {})"
                    .format(len(placed), family_name, type_name, cat_name))

    if report["instance"]:
        output.print_md("### Set on each instance - matched existing ({})".format(len(report["instance"])))
        for col, pname, method, val in report["instance"]:
            arrow = "= `{}`".format(pname) if method == "exact" else "-> `{}` ({})".format(pname, method)
            output.print_md("- **{}** {} = `{}`".format(col, arrow, val))
    if report["type"]:
        output.print_md("### Set on the TYPE - matched existing ({})".format(len(report["type"])))
        for col, pname, method, val in report["type"]:
            arrow = "= `{}`".format(pname) if method == "exact" else "-> `{}` ({})".format(pname, method)
            output.print_md("- **{}** {} = `{}`  _(shared by this type's instances)_".format(col, arrow, val))
    if report["created"]:
        output.print_md("### Set on each instance - NEW instance parameter ({})".format(len(report["created"])))
        for col, pname, val in report["created"]:
            output.print_md("- **{}** -> new `{}` = `{}`".format(col, pname, val))
    if report["unsettable"]:
        output.print_md("### [!] Could not set ({})".format(len(report["unsettable"])))
        for col, reason in report["unsettable"]:
            output.print_md("- **{}** - {}".format(col, reason))
    if skipped:
        output.print_md("### Not mapped ({})".format(len(skipped)))
        for col, reason in skipped:
            output.print_md("- **{}** - {}".format(col, reason))

    forms.alert("Injected {} value(s) across {} placed instance(s)."
                .format(n_total, len(placed)), title="Product Injector - done")


def _surface_error(tb):
    try:
        output.print_md("## Product Injector - ERROR")
        output.print_md("```\n" + tb + "\n```")
    except Exception:
        pass
    try:
        forms.alert("Product Injector hit an error:\n\n" + tb[-1400:],
                    title="Product Injector - ERROR", warn_icon=True)
    except Exception:
        pass


# pyRevit executes this file directly. Run bootstrap + main under one guard so
# ANY failure (module init or main) surfaces as a dialog, never a blank window.
try:
    bootstrap()
    main()
except Exception:
    import traceback
    _surface_error(traceback.format_exc())
    raise
