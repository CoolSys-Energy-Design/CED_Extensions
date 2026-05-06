# -*- coding: utf-8 -*-
__title__ = "Place all Coils"
__doc__ = "Place coil families in spaces based on an Excel description list."

import re
from difflib import SequenceMatcher

import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System
from System import Type, Activator
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
from System.Windows.Forms import (
    Form,
    DataGridView,
    DataGridViewTextBoxColumn,
    DataGridViewComboBoxColumn,
    DockStyle,
    FormStartPosition,
    DataGridViewAutoSizeColumnsMode,
    DialogResult,
    Button,
)
from System.Drawing import Size, Point

from pyrevit import revit, DB, forms, script
from pyrevit.revit import query


logger = script.get_logger()
doc = revit.doc

VERTICAL_OFFSET_FT = 2.0
SPACE_MATCH_THRESHOLD = 0.6
MODEL_MATCH_THRESHOLD = 0.45
SHEET_NAME = "Circuit Schedule"
HEADER_SCAN_ROWS = 60
REQUIRED_MANUFACTURER = "KRACK"
SKIP_SPACE_LABEL = "<Do not place in any space>"

DESC_KEYS = ("description", "desc", "space", "spacename")
COUNT_KEYS = ("coilcount", "coils", "coil", "count")
MODEL_KEYS = ("model", "modelnumber", "modelno")
MFR_KEYS = ("manufacturer", "mfr", "mfg")


try:
    basestring
except NameError:
    basestring = str


def _args_array(*args):
    return System.Array[System.Object](list(args))


def _set(obj, prop, val):
    obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, None, obj, _args_array(val))


def _get(obj, prop):
    try:
        return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, None, obj, None)
    except Exception:
        return None


def _call(obj, name, *args):
    t = obj.GetType()
    try:
        return t.InvokeMember(name, BindingFlags.InvokeMethod, None, obj, _args_array(*args))
    except Exception:
        try:
            return t.InvokeMember(name, BindingFlags.GetProperty, None, obj, _args_array(*args) if args else None)
        except Exception:
            return None


def _cell(cells, r, c):
    it = _call(cells, "Item", r, c)
    v = _get(it, "Value2")
    return ("" if v is None else str(v)).strip()


def _norm_key(value):
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"[^0-9a-zA-Z]+", "", text).lower()
    return text


def _norm_text(value):
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"[^0-9a-zA-Z]+", " ", text).strip().lower()
    return re.sub(r"\s+", " ", text)


def _text_similarity(a, b):
    na = _norm_text(a)
    nb = _norm_text(b)
    if not na or not nb:
        return 0.0
    if na in nb or nb in na:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def _find_header_columns(cells, nrows, ncols):
    desc_col = coil_col = model_col = mfr_col = None
    desc_row = coil_row = model_row = mfr_row = 0
    desc_score = coil_score = model_score = mfr_score = 0
    max_rows = min(nrows, HEADER_SCAN_ROWS)

    for r in range(1, max_rows + 1):
        for c in range(1, ncols + 1):
            raw = _cell(cells, r, c)
            if not raw:
                continue
            raw_l = raw.strip().lower()
            compact = re.sub(r"\s+", "", raw_l)
            if "description" in raw_l:
                score = 2 if raw_l == "description" else 1
                if score > desc_score:
                    desc_col, desc_row, desc_score = c, r, score
            if "coil count" in raw_l:
                score = 2 if raw_l == "coil count" else 1
                if score > coil_score:
                    coil_col, coil_row, coil_score = c, r, score
            if "model#" in compact or compact == "model":
                score = 2 if "model#" in compact or compact == "model#" else 1
                if score > model_score:
                    model_col, model_row, model_score = c, r, score
            if "manufacturer" in raw_l:
                score = 2 if raw_l == "manufacturer" else 1
                if score > mfr_score:
                    mfr_col, mfr_row, mfr_score = c, r, score
        if desc_score == 2 and coil_score == 2 and model_score == 2 and mfr_score == 2:
            break

    return desc_col, desc_row, coil_col, coil_row, model_col, model_row, mfr_col, mfr_row


def _load_circuit_schedule_rows(path):
    xl = wb = ws = used = cells = rows_prop = cols_prop = None
    rows = []
    try:
        t = Type.GetTypeFromProgID("Excel.Application")
        if t is None:
            raise Exception("Excel is not registered on this machine.")
        xl = Activator.CreateInstance(t)
        _set(xl, "Visible", False)
        _set(xl, "DisplayAlerts", False)
        wb = _call(_get(xl, "Workbooks"), "Open", path)
        ws = _call(_get(wb, "Worksheets"), "Item", SHEET_NAME)
        if ws is None:
            raise Exception("Sheet not found: {}".format(SHEET_NAME))

        used = _get(ws, "UsedRange")
        cells = _get(used, "Cells")
        rows_prop = _get(used, "Rows")
        cols_prop = _get(used, "Columns")
        nrows = int(_get(rows_prop, "Count") or 0)
        ncols = int(_get(cols_prop, "Count") or 0)

        desc_col, desc_row, coil_col, coil_row, model_col, model_row, mfr_col, mfr_row = _find_header_columns(
            cells, nrows, ncols
        )
        if not (desc_col and coil_col and model_col and mfr_col):
            raise Exception(
                "Could not locate Description / Coil Count / Model # / Manufacturer columns on '{}'.".format(
                    SHEET_NAME
                )
            )

        start_row = max(desc_row, coil_row, model_row, mfr_row) + 1
        for r in range(start_row, nrows + 1):
            desc = _cell(cells, r, desc_col)
            coil = _cell(cells, r, coil_col)
            model = _cell(cells, r, model_col)
            mfr = _cell(cells, r, mfr_col)
            if not desc and not coil and not model:
                continue
            rows.append({
                "Description": desc,
                "Coil Count": coil,
                "Model #": model,
                "Manufacturer": mfr,
                "_row": r,
            })
    finally:
        try:
            if wb:
                _call(wb, "Close", False)
            if xl:
                _call(xl, "Quit")
        except Exception:
            pass
        try:
            if ws:
                Marshal.ReleaseComObject(ws)
            if wb:
                Marshal.ReleaseComObject(wb)
            if xl:
                Marshal.ReleaseComObject(xl)
        except Exception:
            pass
    return rows


def _get_location_point(elem):
    loc = getattr(elem, "Location", None)
    if loc and hasattr(loc, "Point"):
        return loc.Point
    return None


def _get_bbox_center(elem):
    bbox = None
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if not bbox:
        try:
            bbox = elem.get_BoundingBox(revit.active_view)
        except Exception:
            bbox = None
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) * 0.5


def _resolve_space_point(elem, target_doc=None):
    """Try every reasonable way to get a placement XYZ for a Space/Room.

    For host spaces, leave target_doc as None (uses the active host view as a
    secondary bbox source). For linked spaces, pass the link's document; this
    skips the active-view fallback to avoid mixing link/host coordinate frames.
    The returned point is in target_doc coordinates; callers placing into the
    host must transform with the link's transform.

    Order:
      1. bbox(None)                     - view-independent, own-doc coords.
      2. (host only) bbox(activeView)   - secondary attempt for some host
         spaces. Skipped when target_doc is set because the active view belongs
         to the host doc and would return host coords that callers would then
         double-transform.
      3. Location.Point                 - own-doc coords for any placed Space/Room.

    Returns None when all three fail. In practice that means the space is
    genuinely unplaced (or in a state where Revit can't materialize geometry),
    in which case no per-view fallback would help either - so we don't waste
    time iterating views. The picker surfaces these as [no location].
    """
    # Step 1: bbox(None) - safe for both host and linked.
    bbox = None
    try:
        bbox = elem.get_BoundingBox(None)
    except Exception:
        bbox = None
    if bbox is not None:
        try:
            return (bbox.Min + bbox.Max) * 0.5
        except Exception:
            pass

    # Step 2: bbox(activeView) - host only, to avoid host/link coord mixing.
    if target_doc is None:
        try:
            bbox = elem.get_BoundingBox(revit.active_view)
        except Exception:
            bbox = None
        if bbox is not None:
            try:
                return (bbox.Min + bbox.Max) * 0.5
            except Exception:
                pass

    # Step 3: Location.Point - own-doc coords, safe for both.
    return _get_location_point(elem)


def _find_spatial_element(doc, point):
    if not point:
        return None
    getter = getattr(doc, "GetSpaceAtPoint", None)
    if callable(getter):
        try:
            space = getter(point)
            if space:
                return space
        except Exception:
            pass
    getter = getattr(doc, "GetRoomAtPoint", None)
    if callable(getter):
        try:
            room = getter(point)
            if room:
                return room
        except Exception:
            pass
    return None


def _find_linked_spatial_element(link_doc, point):
    if not point:
        return None
    getter = getattr(link_doc, "GetSpaceAtPoint", None)
    if callable(getter):
        try:
            space = getter(point)
            if space:
                return space
        except Exception:
            pass
    getter = getattr(link_doc, "GetRoomAtPoint", None)
    if callable(getter):
        try:
            room = getter(point)
            if room:
                return room
        except Exception:
            pass
    return None


def _get_param_string(elem, bip):
    try:
        param = elem.get_Parameter(bip)
    except Exception:
        param = None
    if not param:
        return ""
    try:
        val = param.AsString()
        if val:
            return val
    except Exception:
        pass
    try:
        val = param.AsValueString()
        if val:
            return val
    except Exception:
        pass
    return ""


def _space_name(space):
    name = getattr(space, "Name", None)
    if not name:
        name = _get_param_string(space, DB.BuiltInParameter.ROOM_NAME)
    if not name:
        name = _get_param_string(space, DB.BuiltInParameter.SPACE_NAME)
    return (name or "").strip()


def _space_number(space):
    number = getattr(space, "Number", None)
    if not number:
        number = _get_param_string(space, DB.BuiltInParameter.ROOM_NUMBER)
    if not number:
        number = _get_param_string(space, DB.BuiltInParameter.SPACE_NUMBER)
    return (number or "").strip()


def _collect_host_spaces():
    spaces = []
    seen = set()
    categories = [DB.BuiltInCategory.OST_MEPSpaces, DB.BuiltInCategory.OST_Rooms]
    for cat in categories:
        elements = DB.FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
        for space in elements:
            sid = space.Id.IntegerValue
            if sid in seen:
                continue
            seen.add(sid)
            name = _space_name(space)
            number = _space_number(space)
            if not name and not number:
                continue
            point = _resolve_space_point(space)
            spaces.append({
                "element": space,
                "name": name,
                "number": number,
                "point": point,
                "unplaced": point is None,
                "source": "Host",
                "link": None,
                "link_doc": None,
                "link_name": None,
                "key": "host:{}".format(sid),
                "level_name": None,
            })
    return spaces


def _collect_linked_spaces():
    spaces = []
    categories = [DB.BuiltInCategory.OST_MEPSpaces, DB.BuiltInCategory.OST_Rooms]
    link_instances = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance)
    for link in link_instances:
        link_doc = link.GetLinkDocument()
        if link_doc is None:
            continue
        transform = link.GetTransform()
        link_name = getattr(link, "Name", None) or "Link {}".format(link.Id.IntegerValue)
        for cat in categories:
            elements = DB.FilteredElementCollector(link_doc).OfCategory(cat).WhereElementIsNotElementType()
            for space in elements:
                name = _space_name(space)
                number = _space_number(space)
                if not name and not number:
                    continue
                point = _resolve_space_point(space, target_doc=link_doc)
                host_point = transform.OfPoint(point) if point else None
                level_name = None
                try:
                    level = getattr(space, "Level", None)
                    if level is not None:
                        level_name = level.Name
                    elif space.LevelId and space.LevelId != DB.ElementId.InvalidElementId:
                        level_elem = link_doc.GetElement(space.LevelId)
                        level_name = level_elem.Name if level_elem else None
                except Exception:
                    level_name = None
                space_key = "{}:{}".format(link.Id.IntegerValue, space.Id.IntegerValue)
                spaces.append({
                    "element": space,
                    "name": name,
                    "number": number,
                    "point": host_point,
                    "unplaced": host_point is None,
                    "source": "Linked",
                    "link": link,
                    "link_doc": link_doc,
                    "link_name": link_name,
                    "key": space_key,
                    "level_name": level_name,
                })
    return spaces


def _collect_all_spaces():
    return _collect_linked_spaces() + _collect_host_spaces()


def _collect_linked_spaces_from_instance(link):
    """Collect Spaces / Rooms from a single RevitLinkInstance. Mirrors
    _collect_linked_spaces() but scoped to one link, which is what we want
    when the user has explicitly chosen a link as the mapping source."""
    spaces = []
    if link is None:
        return spaces
    link_doc = link.GetLinkDocument()
    if link_doc is None:
        return spaces
    transform = link.GetTransform()
    link_name = getattr(link, "Name", None) or "Link {}".format(link.Id.IntegerValue)
    categories = [DB.BuiltInCategory.OST_MEPSpaces, DB.BuiltInCategory.OST_Rooms]
    for cat in categories:
        elements = DB.FilteredElementCollector(link_doc).OfCategory(cat).WhereElementIsNotElementType()
        for space in elements:
            name = _space_name(space)
            number = _space_number(space)
            if not name and not number:
                continue
            point = _resolve_space_point(space, target_doc=link_doc)
            host_point = transform.OfPoint(point) if point else None
            level_name = None
            try:
                level = getattr(space, "Level", None)
                if level is not None:
                    level_name = level.Name
                elif space.LevelId and space.LevelId != DB.ElementId.InvalidElementId:
                    level_elem = link_doc.GetElement(space.LevelId)
                    level_name = level_elem.Name if level_elem else None
            except Exception:
                level_name = None
            space_key = "{}:{}".format(link.Id.IntegerValue, space.Id.IntegerValue)
            spaces.append({
                "element": space,
                "name": name,
                "number": number,
                "point": host_point,
                "unplaced": host_point is None,
                "source": "Linked",
                "link": link,
                "link_doc": link_doc,
                "link_name": link_name,
                "key": space_key,
                "level_name": level_name,
            })
    return spaces


def _pick_space_source():
    """Prompt the user to pick where the mapping picker should pull spaces from:
    either the host model or one specific RevitLinkInstance.

    Returns:
      ("host", None)                - user picked the host model
      ("linked", link_instance)     - user picked a specific link
      None                          - user cancelled
    """
    label_to_choice = {}
    options = []

    host_label = "Host model (this project)"
    label_to_choice[host_label] = ("host", None)
    options.append(host_label)

    link_entries = []
    for link in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
        link_doc = link.GetLinkDocument()
        if link_doc is None:
            continue
        link_name = getattr(link, "Name", None) or "Link {}".format(link.Id.IntegerValue)
        try:
            doc_title = link_doc.Title or ""
        except Exception:
            doc_title = ""
        label = link_name
        if doc_title and doc_title.lower() not in label.lower():
            label = "{}  ({})".format(link_name, doc_title)
        link_entries.append((label, link))

    link_entries.sort(key=lambda e: e[0].lower())

    # Prefer a link whose name suggests it's the architectural model.
    default_link_label = None
    for label, link in link_entries:
        if "arch" in label.lower():
            default_link_label = label
            break

    for label, link in link_entries:
        display = label
        if default_link_label and label == default_link_label:
            display = "(default) " + label
        label_to_choice[display] = ("linked", link)
        options.append(display)

    # Pin the default to the top.
    if default_link_label:
        marked = "(default) " + default_link_label
        if marked in options:
            options.remove(marked)
            options.insert(0, marked)

    if len(options) == 1:
        # Only host model is available; auto-pick to save a click.
        return label_to_choice[options[0]]

    picked = forms.SelectFromList.show(
        options,
        title="Select Source for Spaces / Rooms",
        button_name="Use this source",
        multiselect=False,
    )
    if not picked:
        return None
    return label_to_choice.get(picked)


def _collect_spaces_from_source(source_kind, link_instance):
    if source_kind == "host":
        return _collect_host_spaces()
    if source_kind == "linked":
        return _collect_linked_spaces_from_instance(link_instance)
    return []


def _space_keys(space):
    keys = []
    if space.get("name"):
        keys.append(space["name"])
    if space.get("number"):
        keys.append(space["number"])
    if space.get("name") and space.get("number"):
        keys.append("{} {}".format(space["number"], space["name"]))
    return keys


def _space_display(space):
    name = space.get("name") or ""
    number = space.get("number") or ""
    label = name or number or "Unnamed Space"
    if name and number:
        label = "{} ({})".format(name, number)
    source = space.get("source") or "Host"
    suffix = " [no location]" if space.get("unplaced") else ""
    if source == "Linked":
        link_name = space.get("link_name") or "Link"
        return "{} [Linked: {}]{}".format(label, link_name, suffix)
    return "{} [Host]{}".format(label, suffix)


def _prompt_space_mapping(descriptions, spaces):
    if not descriptions:
        return {}

    label_map = {}
    labels = []
    for space in spaces:
        label = _space_display(space)
        if label in label_map:
            label = "{} ({})".format(label, space.get("key"))
        label_map[label] = space
        labels.append(label)
    labels = [SKIP_SPACE_LABEL] + sorted(labels, key=lambda s: s.lower())

    form = Form()
    form.Text = "Map Descriptions to Spaces"
    form.Size = Size(900, 600)
    form.StartPosition = FormStartPosition.CenterScreen

    grid = DataGridView()
    grid.Dock = DockStyle.Top
    grid.Height = 520
    grid.AllowUserToAddRows = False
    grid.AllowUserToDeleteRows = False
    grid.ReadOnly = False
    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    col_desc = DataGridViewTextBoxColumn()
    col_desc.HeaderText = "Description"
    col_desc.ReadOnly = True
    grid.Columns.Add(col_desc)

    col_space = DataGridViewComboBoxColumn()
    col_space.HeaderText = "Space (Linked/Host)"
    col_space.DataSource = labels
    grid.Columns.Add(col_space)

    for desc in descriptions:
        idx = grid.Rows.Add()
        grid.Rows[idx].Cells[0].Value = desc
        grid.Rows[idx].Cells[1].Value = SKIP_SPACE_LABEL

    ok_btn = Button()
    ok_btn.Text = "OK"
    ok_btn.Size = Size(100, 30)
    ok_btn.Location = Point(680, 530)
    ok_btn.DialogResult = DialogResult.OK

    cancel_btn = Button()
    cancel_btn.Text = "Cancel"
    cancel_btn.Size = Size(100, 30)
    cancel_btn.Location = Point(790, 530)
    cancel_btn.DialogResult = DialogResult.Cancel

    form.Controls.Add(grid)
    form.Controls.Add(ok_btn)
    form.Controls.Add(cancel_btn)
    form.AcceptButton = ok_btn
    form.CancelButton = cancel_btn

    if form.ShowDialog() != DialogResult.OK:
        return None

    mapping = {}
    for row in grid.Rows:
        desc_val = row.Cells[0].Value
        space_val = row.Cells[1].Value
        if desc_val is None:
            continue
        desc_text = str(desc_val)
        if space_val is None:
            mapping[desc_text] = None
            continue
        space_label = str(space_val)
        if space_label == SKIP_SPACE_LABEL:
            mapping[desc_text] = None
            continue
        mapping[desc_text] = label_map.get(space_label)

    return mapping


def _best_space_match(description, spaces):
    best_space = None
    best_score = 0.0
    for space in spaces:
        for key in _space_keys(space):
            score = _text_similarity(description, key)
            if score > best_score:
                best_score = score
                best_space = space
    return best_space, best_score


def _collect_mech_symbols():
    symbols = []
    cat_id = int(DB.BuiltInCategory.OST_MechanicalEquipment)
    for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        try:
            if symbol.Category and symbol.Category.Id.IntegerValue == cat_id:
                symbols.append(symbol)
        except Exception:
            continue
    return symbols


def _matches_required_mfr(symbol):
    if not REQUIRED_MANUFACTURER:
        return True
    mfr_u = REQUIRED_MANUFACTURER.upper()
    try:
        fam_name = query.get_name(symbol.Family) or ""
    except Exception:
        fam_name = ""
    try:
        type_name = query.get_name(symbol) or ""
    except Exception:
        type_name = ""
    return mfr_u in fam_name.upper() or mfr_u in type_name.upper()


def _filter_required_mfr_symbols(symbols):
    return [s for s in symbols if _matches_required_mfr(s)]


def _collect_symbols_by_space(spaces):
    symbol_map = {}
    space_key_lookup = {}
    link_instances = {}
    for space in spaces:
        link = space.get("link")
        link_doc = space.get("link_doc")
        if not link or not link_doc:
            continue
        link_id = link.Id.IntegerValue
        link_instances[link_id] = link
        space_key_lookup[(link_id, space["element"].Id.IntegerValue)] = space.get("key")

    if not link_instances:
        return symbol_map

    cat = DB.BuiltInCategory.OST_MechanicalEquipment
    elems = DB.FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
    for elem in elems:
        point = _get_location_point(elem)
        if not point:
            continue
        symbol = getattr(elem, "Symbol", None)
        if not symbol:
            continue
        for link_id, link in link_instances.items():
            link_doc = link.GetLinkDocument()
            if link_doc is None:
                continue
            transform = link.GetTransform()
            inv = transform.Inverse
            link_point = inv.OfPoint(point)
            spatial = _find_linked_spatial_element(link_doc, link_point)
            if not spatial:
                continue
            key = space_key_lookup.get((link_id, spatial.Id.IntegerValue))
            if not key:
                continue
            symbol_map.setdefault(key, set()).add(symbol)
            break
    return symbol_map


def _best_symbol_match(model, symbols):
    """Return (best_symbol, score, top_candidates).

    best_symbol is the symbol with the highest display_score among candidates
    that passed the manufacturer filter, or None if the candidate pool is
    empty. The score returned is the display_score (max of match_score and
    text_similarity) so the caller can compare against MODEL_MATCH_THRESHOLD
    to label the placement as a confident `match` (>= threshold) versus a
    best-effort `fallback` (< threshold) - but the symbol is always returned
    when at least one KRACK candidate exists, never gated by the threshold.

    top_candidates is the same list of (display_score, label) sorted
    descending, capped at 5 entries.
    """
    if not model:
        return None, 0.0, []
    keys = _model_keys(model)
    best_symbol = None
    best_score = -1.0  # so the first valid candidate always wins, even at 0.0
    candidates = []
    for symbol in symbols:
        try:
            fam_name = query.get_name(symbol.Family)
        except Exception:
            fam_name = ""
        try:
            type_name = query.get_name(symbol)
        except Exception:
            type_name = ""
        if REQUIRED_MANUFACTURER:
            mfr_u = REQUIRED_MANUFACTURER.upper()
            if mfr_u not in (fam_name or "").upper() and mfr_u not in (type_name or "").upper():
                continue
        label = "{} : {}".format(fam_name, type_name)
        label_key = _norm_model_key(label)
        match_score = 0.0
        if keys:
            if any(k and k in label_key for k in keys):
                match_score = 1.0
        else:
            match_score = max(
                _text_similarity(model, label),
                _text_similarity(model, type_name),
            )
        text_score = max(
            _text_similarity(model, label),
            _text_similarity(model, type_name),
        )
        display_score = match_score if match_score > text_score else text_score
        candidates.append((display_score, label))
        if display_score > best_score:
            best_score = display_score
            best_symbol = symbol
    candidates.sort(key=lambda c: -c[0])
    if best_symbol is None:
        return None, 0.0, candidates[:5]
    return best_symbol, max(best_score, 0.0), candidates[:5]


def _level_belongs_to_host(level):
    if level is None:
        return False
    try:
        lev_doc = level.Document
    except Exception:
        return False
    if lev_doc is None:
        return False
    try:
        if lev_doc.Equals(doc):
            return True
    except Exception:
        pass
    try:
        return lev_doc.PathName == doc.PathName
    except Exception:
        return False


def _resolve_level(space):
    """Return a Level element that belongs to the host document.

    Critical: doc.Create.NewFamilyInstance refuses a level from any other
    document (including a linked one), so this function must never return a
    link-doc level.

    Strategy:
      1. Match a host-doc level by name (works for both host and linked
         spaces - the linked space's level_name is captured during collection).
      2. For host spaces: trust the space's own .Level / .LevelId (already
         host-doc). Skip this for linked spaces.
      3. Match a host-doc level by elevation closest to the placement point's
         Z (linked spaces have already had their point transformed to host
         coords, so this is apples-to-apples).
      4. Active view's GenLevel.
      5. First host-doc level as a last resort.
    """
    if isinstance(space, dict):
        level_name = space.get("level_name")
        space_elem = space.get("element")
        is_linked = (space.get("source") or "Host") == "Linked"
        host_point = space.get("point")
    else:
        level_name = None
        space_elem = space
        is_linked = False
        host_point = None

    host_levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))

    # 1. Name match against host levels.
    if level_name:
        for lvl in host_levels:
            if lvl.Name == level_name:
                return lvl

    # 2. Host space's own level (linked spaces would return a link-doc level
    #    here, which would crash NewFamilyInstance - so skip for linked).
    if not is_linked and space_elem is not None:
        level = getattr(space_elem, "Level", None)
        if _level_belongs_to_host(level):
            return level
        try:
            lid = getattr(space_elem, "LevelId", None)
            if lid is not None and lid != DB.ElementId.InvalidElementId:
                lev = doc.GetElement(lid)
                if _level_belongs_to_host(lev):
                    return lev
        except Exception:
            pass

    # 3. Closest host level by elevation to the placement point Z.
    if host_point is not None and host_levels:
        try:
            target_z = float(host_point.Z)
            return min(host_levels, key=lambda l: abs(float(l.Elevation) - target_z))
        except Exception:
            pass

    # 4. Active view's GenLevel.
    try:
        view = revit.active_view
        gen_level = getattr(view, "GenLevel", None)
        if _level_belongs_to_host(gen_level):
            return gen_level
    except Exception:
        pass

    # 5. First host level.
    return host_levels[0] if host_levels else None


def _format_level_elevation(level):
    try:
        elev = float(level.Elevation)
    except Exception:
        return "?"
    sign = "-" if elev < 0 else ""
    elev_abs = abs(elev)
    ft = int(elev_abs)
    inches = int(round((elev_abs - ft) * 12.0))
    if inches == 12:
        ft += 1
        inches = 0
    return "{}{}'-{}\"".format(sign, ft, inches)


_DEFAULT_LEVEL_TIERS = [
    # Tier 1: an exact-ish "Level 1" / "L1".
    [
        r"^\s*level\s*0*1\s*$",
        r"^\s*l\s*0*1\s*$",
    ],
    # Tier 2: "Level 1" appearing as a token, or names that start with "01 ...".
    [
        r"\blevel\s*0*1\b",
        r"^\s*0*1\b",
    ],
    # Tier 3: ground/first-floor terms.
    [
        r"\bground\s*(floor)?\b",
        r"\bfirst\s*floor\b",
        r"\b1st\s*floor\b",
    ],
]


def _find_default_level(levels):
    """Best-effort guess at which host level is 'Level 1' by name pattern.

    Falls back to the lowest-elevation level when no name matches."""
    if not levels:
        return None
    sorted_levels = sorted(levels, key=lambda l: l.Elevation)
    for tier in _DEFAULT_LEVEL_TIERS:
        compiled = [re.compile(p, re.IGNORECASE) for p in tier]
        for lvl in sorted_levels:
            name = lvl.Name or ""
            for pat in compiled:
                if pat.search(name):
                    return lvl
    return sorted_levels[0]


def _pick_host_level():
    """Prompt the user to pick a host-doc level. Defaults the selection to a
    'Level 1'-ish level if any exists, otherwise the lowest-elevation level."""
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    if not levels:
        return None

    sorted_levels = sorted(levels, key=lambda l: l.Elevation)
    default = _find_default_level(sorted_levels)

    label_to_level = {}
    options = []
    default_label = None
    for lvl in sorted_levels:
        label = "{}  ({})".format(lvl.Name or "<Unnamed>", _format_level_elevation(lvl))
        if default is not None and lvl.Id.IntegerValue == default.Id.IntegerValue:
            label = "(default) " + label
            default_label = label
        label_to_level[label] = lvl
        options.append(label)

    # Pin the default to the top of the list.
    if default_label and default_label in options:
        options.remove(default_label)
        options.insert(0, default_label)

    picked = forms.SelectFromList.show(
        options,
        title="Select Level for Coil Placement",
        button_name="Use this level",
        multiselect=False,
    )
    if not picked:
        return None
    return label_to_level.get(picked)


def _parse_models(raw):
    if raw is None:
        return []
    text = str(raw).strip()
    if not text:
        return []
    return [p.strip() for p in text.split(",") if p.strip()]


def _norm_model_key(value):
    if value is None:
        return ""
    return re.sub(r"[^0-9A-Za-z]+", "", str(value)).upper()


def _model_keys(model):
    base = _norm_model_key(model)
    if not base:
        return []
    keys = [base]
    trimmed = re.sub(r"[A-Z]+$", "", base)
    if trimmed and trimmed != base:
        keys.append(trimmed)
    return keys


def _expand_models(models, count):
    if count <= 0:
        return []
    if not models:
        return [None] * count
    if count <= len(models):
        return models[:count]
    expanded = []
    idx = 0
    while len(expanded) < count:
        expanded.append(models[idx % len(models)])
        idx += 1
    return expanded


def _coerce_int(value):
    if value is None:
        return 0
    if isinstance(value, basestring):
        text = value.strip()
        if not text:
            return 0
        try:
            return int(float(text))
        except Exception:
            return 0
    try:
        return int(value)
    except Exception:
        try:
            return int(float(value))
        except Exception:
            return 0


def _row_value(row, keyset):
    if not row:
        return None
    norm = {}
    for key, val in row.items():
        if key is None:
            continue
        norm[_norm_key(key)] = val
    for key in keyset:
        if key in norm:
            return norm[key]
    return None


def _load_excel_rows(path):
    return _load_circuit_schedule_rows(path)


def _build_placements(rows, spaces, symbol_by_space, all_symbols, chosen_level=None):
    placements = []
    warnings = []
    match_info = []
    model_attempts = []
    stats = {
        "rows_total": len(rows or []),
        "skipped_non_mfr": 0,
        "skipped_missing_mfr": 0,
        "skipped_no_count": 0,
        "skipped_no_desc": 0,
        "skipped_no_space": 0,
        "skipped_no_point": 0,
        "skipped_no_symbols": 0,
        "skipped_no_candidates": 0,
        "placements_match": 0,
        "placements_fallback": 0,
        "rows_processed": 0,
    }
    desc_order = []
    desc_seen = set()
    for row in rows:
        desc = _row_value(row, DESC_KEYS)
        count_val = _row_value(row, COUNT_KEYS)
        count = _coerce_int(count_val)
        if not desc:
            continue
        if count <= 0:
            continue
        if desc not in desc_seen:
            desc_seen.add(desc)
            desc_order.append(desc)
    mapping = _prompt_space_mapping(desc_order, spaces)
    if mapping is None:
        script.exit()
    for idx, row in enumerate(rows, start=2):
        row_idx = row.get("_row") if isinstance(row, dict) else None
        idx_label = row_idx if row_idx is not None else idx
        desc = _row_value(row, DESC_KEYS)
        count = _coerce_int(_row_value(row, COUNT_KEYS))
        models_raw = _row_value(row, MODEL_KEYS)
        manufacturer = _row_value(row, MFR_KEYS)
        if manufacturer:
            if REQUIRED_MANUFACTURER.lower() not in _norm_text(manufacturer):
                stats["skipped_non_mfr"] += 1
                continue
        else:
            stats["skipped_missing_mfr"] += 1
            continue

        if count <= 0:
            stats["skipped_no_count"] += 1
            continue
        if not desc:
            stats["skipped_no_desc"] += 1
            warnings.append("Row {}: missing description.".format(idx_label))
            continue

        space = mapping.get(desc)
        if space is None:
            stats["skipped_no_space"] += 1
            warnings.append("Row {}: no space selected for '{}'.".format(idx_label, desc))
            continue
        score = 0.0
        for key in _space_keys(space):
            score = max(score, _text_similarity(desc, key))
        match_info.append({
            "row": idx_label,
            "desc": desc,
            "space": (space.get("name") or space.get("number")) if space else "",
            "score": score,
            "passed": True,
            "source": space.get("source") or "Host",
        })

        models = _expand_models(_parse_models(models_raw), count)
        base_point = space.get("point")
        if not base_point:
            stats["skipped_no_point"] += 1
            warnings.append(
                "Row {}: chosen space '{}' has no usable 3D location after trying "
                "bbox(None), bbox(activeView), Location.Point, and bbox in every "
                "plan/3D view. Likely an unplaced Room, a phase mismatch, or the "
                "space lives in a linked file. Spaces in this state are now flagged "
                "'[no location]' in the picker so you can pick a different one.".format(
                    idx_label, space.get("name") or space.get("number")
                )
            )
            continue

        space_key = space.get("key") or space["element"].Id.IntegerValue
        symbols = list(symbol_by_space.get(space_key, [])) or list(all_symbols)
        if not symbols:
            stats["skipped_no_symbols"] += 1
            warnings.append(
                "Row {}: no mechanical equipment types available for '{}'".format(
                    idx_label, space.get("name") or space.get("number")
                )
            )
            continue

        stats["rows_processed"] += 1
        # Use the user-picked host level if provided. Falls back to the legacy
        # heuristic resolver for callers that don't pass one.
        level = chosen_level if chosen_level is not None else _resolve_level(space)
        for offset_idx, model in enumerate(models):
            symbol, sym_score, top_candidates = _best_symbol_match(model, symbols)
            attempt = {
                "row": idx_label,
                "desc": desc,
                "space": space.get("name") or space.get("number"),
                "model": model or "",
                "best_label": "",
                "score": sym_score,
                "candidates": list(top_candidates or []),
                "placed": False,
                "reason": "",
            }
            if symbol is None:
                # No KRACK candidates at all - nothing we can place.
                stats["skipped_no_candidates"] += 1
                attempt["reason"] = "no candidate symbols available"
                model_attempts.append(attempt)
                warnings.append(
                    "Row {}: no candidate family types available for model '{}' in '{}'.".format(
                        idx_label, model or "", space.get("name") or space.get("number")
                    )
                )
                continue
            fam_name = ""
            typ_name = ""
            try:
                fam_name = query.get_name(symbol.Family)
            except Exception:
                fam_name = ""
            try:
                typ_name = query.get_name(symbol)
            except Exception:
                typ_name = ""
            attempt["best_label"] = "{} : {}".format(fam_name, typ_name)
            attempt["placed"] = True
            if sym_score >= MODEL_MATCH_THRESHOLD:
                attempt["reason"] = "match"
                stats["placements_match"] += 1
            else:
                attempt["reason"] = "fallback (top candidate, score {:.2f} < {:.2f})".format(
                    sym_score, MODEL_MATCH_THRESHOLD
                )
                stats["placements_fallback"] += 1
                warnings.append(
                    "Row {}: model '{}' had no confident match in '{}' - placing top candidate '{} : {}' (score {:.2f}). Verify the type is correct.".format(
                        idx_label,
                        model or "",
                        space.get("name") or space.get("number"),
                        fam_name,
                        typ_name,
                        sym_score,
                    )
                )
            model_attempts.append(attempt)
            placements.append({
                "symbol": symbol,
                "point": base_point,
                "level": level,
                "offset": offset_idx * VERTICAL_OFFSET_FT,
                "space": space.get("name") or space.get("number"),
                "model": model,
                "manufacturer": manufacturer,
                "family": fam_name,
                "type": typ_name,
                "score": sym_score,
                "desc": desc,
                "space_score": score,
                "fallback": sym_score < MODEL_MATCH_THRESHOLD,
            })
    return placements, warnings, stats, match_info, model_attempts


def _place_instances(placements):
    placed_ids = []
    failures = []
    with revit.Transaction("Place all Coils"):
        for item in placements:
            symbol = item["symbol"]
            if not symbol:
                continue
            try:
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()
            except Exception:
                pass

            # Defensive: NewFamilyInstance refuses a level from any other doc
            # (a linked-doc level slipping through here is what produced the
            # "level does not exist in the given document" error). If the
            # resolved level isn't from the host doc, drop it.
            level = item.get("level")
            if level is not None and not _level_belongs_to_host(level):
                logger.warning(
                    "Discarding non-host-doc level for {}; placing without explicit level.".format(
                        item.get("model") or symbol.Name
                    )
                )
                level = None

            inst = None
            try:
                if level is not None:
                    inst = doc.Create.NewFamilyInstance(
                        item["point"],
                        symbol,
                        level,
                        DB.Structure.StructuralType.NonStructural,
                    )
                else:
                    inst = doc.Create.NewFamilyInstance(
                        item["point"],
                        symbol,
                        DB.Structure.StructuralType.NonStructural,
                    )
            except Exception as ex:
                failures.append("Failed to place {}: {}".format(item.get("model") or symbol.Name, ex))
                continue

            if inst is None:
                failures.append("Failed to place {} (unknown error).".format(item.get("model") or symbol.Name))
                continue

            try:
                if item["offset"]:
                    DB.ElementTransformUtils.MoveElement(
                        doc,
                        inst.Id,
                        DB.XYZ(0, float(item["offset"]), 0),
                    )
            except Exception as ex:
                failures.append("Failed to offset {}: {}".format(inst.Id.IntegerValue, ex))

            placed_ids.append(inst.Id)

    return placed_ids, failures


def main():
    path = forms.pick_file(file_ext="xlsx", title="Select Coil Placement Excel File")
    if not path:
        return

    try:
        rows = _load_excel_rows(path)
    except Exception as ex:
        forms.alert(
            "Failed to read '{}' sheet: {}".format(SHEET_NAME, ex),
            exitscript=True,
        )
    if not rows:
        forms.alert("No readable rows found on '{}'.".format(SHEET_NAME), exitscript=True)

    space_source = _pick_space_source()
    if space_source is None:
        forms.alert("No source selected. Cancelled.", exitscript=True)
    source_kind, source_link = space_source

    spaces = _collect_spaces_from_source(source_kind, source_link)
    if not spaces:
        if source_kind == "linked":
            link_label = (
                getattr(source_link, "Name", None)
                or "Link {}".format(source_link.Id.IntegerValue)
            )
            forms.alert(
                "No Spaces or Rooms found in the selected link '{}'.".format(link_label),
                exitscript=True,
            )
        else:
            forms.alert(
                "No Spaces or Rooms found in the host model.",
                exitscript=True,
            )

    symbol_by_space = _collect_symbols_by_space(spaces)
    all_symbols = _collect_mech_symbols()
    if not all_symbols:
        forms.alert("No Mechanical Equipment family types found in this model.", exitscript=True)

    mfr_symbols = _filter_required_mfr_symbols(all_symbols)

    chosen_level = _pick_host_level()
    if chosen_level is None:
        forms.alert("No level selected. Cancelled.", exitscript=True)

    placements, warnings, stats, match_info, model_attempts = _build_placements(
        rows, spaces, symbol_by_space, all_symbols, chosen_level=chosen_level
    )

    placed_ids = []
    failures = []
    if placements:
        placed_ids, failures = _place_instances(placements)
        if placed_ids:
            revit.get_selection().set_to(placed_ids)

    output = script.get_output()
    output.close_others()
    output.print_md("### Place all Coils")

    if placements:
        output.print_md("Placed **{}** coil(s).".format(len(placed_ids)))
    else:
        output.print_md(
            "**No coils were placed.** See diagnostic details below to determine why."
        )

    # Excel + symbol summary so failures have hard numbers up front.
    output.print_md("#### Inputs")
    output.print_md("- Excel file: `{}`".format(path))
    output.print_md(
        "- Excel rows read from sheet '{}': **{}**".format(SHEET_NAME, stats.get("rows_total", 0))
    )
    if source_kind == "linked":
        link_label = (
            getattr(source_link, "Name", None)
            or "Link {}".format(source_link.Id.IntegerValue)
        )
        source_label = "linked model '{}'".format(link_label)
    else:
        source_label = "host model"
    unplaced_count = sum(1 for s in spaces if s.get("unplaced"))
    output.print_md(
        "- Spaces / Rooms source: **{}** ({} space(s); {} flagged `[no location]`)".format(
            source_label, len(spaces), unplaced_count
        )
    )
    output.print_md(
        "- Mechanical Equipment family types in project: **{}** total, **{}** matching '{}' (case-insensitive)".format(
            len(all_symbols),
            len(mfr_symbols),
            REQUIRED_MANUFACTURER,
        )
    )
    if not mfr_symbols:
        output.print_md(
            "  - **No '{0}' families are loaded in this project.** Load at least one '{0}' "
            "Mechanical Equipment family before running this tool.".format(REQUIRED_MANUFACTURER)
        )
    output.print_md(
        "- Placement level: **{}** ({})".format(
            chosen_level.Name if chosen_level is not None else "<none>",
            _format_level_elevation(chosen_level) if chosen_level is not None else "?",
        )
    )

    output.print_md("#### Row Skip Counts")
    output.print_md("- Rows processed (passed all filters): **{}**".format(stats.get("rows_processed", 0)))
    output.print_md(
        "- Skipped because manufacturer != '{}': {}".format(
            REQUIRED_MANUFACTURER, stats.get("skipped_non_mfr", 0)
        )
    )
    output.print_md("- Skipped because manufacturer column was blank: {}".format(stats.get("skipped_missing_mfr", 0)))
    output.print_md("- Skipped because Coil Count <= 0: {}".format(stats.get("skipped_no_count", 0)))
    output.print_md("- Skipped because Description was blank: {}".format(stats.get("skipped_no_desc", 0)))
    output.print_md("- Skipped because no space was selected for the Description: {}".format(stats.get("skipped_no_space", 0)))
    output.print_md("- Skipped because the chosen space had no placement point: {}".format(stats.get("skipped_no_point", 0)))
    output.print_md("- Skipped because no Mechanical Equipment types were available: {}".format(stats.get("skipped_no_symbols", 0)))
    output.print_md("- Skipped because no candidate family types existed (no '{}' families loaded): {}".format(
        REQUIRED_MANUFACTURER, stats.get("skipped_no_candidates", 0)
    ))

    output.print_md("#### Placement Quality")
    output.print_md(
        "- Confident matches (score >= {:.2f}): **{}**".format(
            MODEL_MATCH_THRESHOLD, stats.get("placements_match", 0)
        )
    )
    output.print_md(
        "- Fallback placements (top candidate, score < {:.2f}): **{}** - review the Model # Match Attempts list to verify correctness.".format(
            MODEL_MATCH_THRESHOLD, stats.get("placements_fallback", 0)
        )
    )

    space_map = {}
    for p in placements:
        space_name = p.get("space") or "Unknown"
        fam_label = "{} : {}".format(p.get("family") or "", p.get("type") or "").strip()
        if fam_label == ":" or fam_label == "":
            fam_label = "Unknown Family"
        space_map.setdefault(space_name, {})
        space_map[space_name][fam_label] = space_map[space_name].get(fam_label, 0) + 1

    if space_map:
        output.print_md("#### Spaces and Coils Placed")
        for space_name in sorted(space_map.keys()):
            fam_counts = space_map[space_name]
            fam_list = ["{} x{}".format(name, fam_counts[name]) for name in sorted(fam_counts.keys())]
            output.print_md("- {}: {}".format(space_name, ", ".join(fam_list)))

    if match_info:
        output.print_md("#### Description → Space Matches")
        for entry in sorted(match_info, key=lambda e: e.get("row", 0)):
            desc = entry.get("desc") or ""
            space_name = entry.get("space") or "(no match)"
            score = entry.get("score", 0.0)
            status = "OK" if entry.get("passed") else "LOW"
            source = entry.get("source") or "Host"
            output.print_md(
                "- Row {}: '{}' → '{}' [{}] (score {:.2f}, {})".format(
                    entry.get("row"),
                    desc,
                    space_name,
                    source,
                    score,
                    status,
                )
            )

    if model_attempts:
        output.print_md("#### Model # Match Attempts")
        for attempt in model_attempts:
            reason = (attempt.get("reason") or "").lower()
            if not attempt.get("placed"):
                status = "MISS"
            elif "fallback" in reason:
                status = "FALLBACK"
            else:
                status = "PLACED"
            best_label = attempt.get("best_label") or "(none)"
            output.print_md(
                "- Row {} [{}] '{}': model `{}` → `{}` (score {:.2f}) — **{}** ({})".format(
                    attempt.get("row"),
                    attempt.get("space") or "?",
                    attempt.get("desc") or "",
                    attempt.get("model") or "",
                    best_label,
                    attempt.get("score", 0.0),
                    status,
                    attempt.get("reason") or "",
                )
            )
            if not attempt.get("placed"):
                cands = attempt.get("candidates") or []
                if cands:
                    output.print_md("    Top candidates considered:")
                    for cand_score, cand_label in cands:
                        output.print_md(
                            "      - `{}` (similarity {:.2f})".format(cand_label, cand_score)
                        )
                elif not mfr_symbols:
                    output.print_md(
                        "    No '{}' family types loaded in the project, so no candidates exist.".format(
                            REQUIRED_MANUFACTURER
                        )
                    )
                else:
                    output.print_md("    No candidates were considered for this model.")

    if warnings:
        output.print_md("#### Warnings")
        for message in warnings:
            output.print_md("- {}".format(message))

    if failures:
        output.print_md("#### Placement Failures")
        for message in failures:
            output.print_md("- {}".format(message))

    if not placements:
        forms.alert(
            "No valid coil placements were generated.\n\n"
            "Open the script output panel for a full diagnostic report "
            "(rows read, manufacturer matches, symbols available, model match attempts).",
            title="Place all Coils",
        )


if __name__ == "__main__":
    main()
