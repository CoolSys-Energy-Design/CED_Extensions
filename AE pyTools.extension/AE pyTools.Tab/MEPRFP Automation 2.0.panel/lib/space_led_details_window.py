# -*- coding: utf-8 -*-
"""
Modal sub-dialog for per-LED detail editing on a Space profile.

Three tabs in one dialog:

  * **Parameters** — flat key/value DataGrid bound to ``led.parameters``.
  * **Offsets**    — DataGrid (X / Y / Z / rotation) bound to
                     ``led.offsets``.
  * **Annotations** — DataGrid of annotation rows (kind, label, family,
                      type, offset). A per-row "Params..." button pops
                      a generic key/value editor for the annotation's
                      own ``parameters`` dict.

All edits land on the in-memory dicts of the calling Manage Space
Profiles editor; the parent window is responsible for persisting them
to the active YAML payload on save. OK commits in-flight cell edits
and closes; Cancel discards the current cell-in-progress.
"""

import copy
import os
import uuid

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import (  # noqa: E402
    Object as _NetObject,
    String as _NetString,
)
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402

import wpf as _wpf  # noqa: E402
import revit_symbol_index as _sym_index  # noqa: E402


def _make_clr_string_list(py_iter):
    """Wrap a Python iterable of strings in an ``ObservableCollection``
    so WPF ItemsSource bindings have something to enumerate.

    Plain CLR ``List<String>`` returned through Python reflection has
    been observed to surface as empty in the binding target;
    ObservableCollection is the codebase-proven choice for ItemsSource.
    """
    out = ObservableCollection[_NetObject]()
    for s in py_iter or ():
        out.Add(_NetString(str(s)))
    return out


_RESOURCES = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources",
)
_DETAILS_XAML = os.path.join(_RESOURCES, "SpaceLedDetailsDialog.xaml")
_KV_XAML = os.path.join(_RESOURCES, "KeyValueDialog.xaml")


_ANNOTATION_KINDS = ("tag", "keynote", "text_note")
_ANNOTATION_KIND_OPTIONS_NET = _make_clr_string_list(_ANNOTATION_KINDS)


# ---------------------------------------------------------------------
# Shared row classes
# ---------------------------------------------------------------------

class _ParamRow(object):
    """Two-column key/value row.

    Plain instance attributes don't survive pythonnet 3's TwoWay
    binding write-back — the cell looks editable but the user's
    typed value is silently dropped because TypeDescriptor exposes
    the attribute as read-only. Explicit ``@property`` + ``@setter``
    pairs make the property writable through reflection.
    """

    def __init__(self, name="", value=""):
        self._name = "" if name is None else str(name)
        self._value = "" if value is None else _coerce_to_text(value)

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, value):
        self._name = "" if value is None else str(value)

    @property
    def Value(self):
        return self._value

    @Value.setter
    def Value(self, value):
        self._value = "" if value is None else str(value)


def _coerce_to_text(value):
    """Stringify a parameter value for display.

    Dict / list values come from directives or structured params and
    can't be edited inline; display them but mark them protected via
    repr() so the user understands they're not plain text.
    """
    if isinstance(value, (dict, list)):
        return repr(value)
    return str(value)


def _coerce_from_text(name, text, original_value):
    """Best-effort parse of edited text back into the YAML value type.

    Numeric-looking strings stay text — Revit's parameter writer
    handles unit conversion. We only special-case directive dicts
    (``{...}`` literal) and lists by parsing them with ``ast.literal_eval``;
    plain strings are returned as-is.
    """
    if text is None:
        return None
    s = str(text)
    if not s.strip():
        return ""
    # If the original was a dict / list, try to round-trip through repr.
    if isinstance(original_value, (dict, list)) and (s.startswith("{") or s.startswith("[")):
        try:
            import ast
            return ast.literal_eval(s)
        except Exception:
            return s
    return s


def _to_float(text, default=0.0):
    if text is None:
        return default
    s = str(text).strip()
    if not s:
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def _fmt_float(value):
    if value is None:
        return ""
    try:
        v = float(value)
    except (ValueError, TypeError):
        return ""
    if v == int(v):
        return str(int(v))
    return "{:.4f}".format(v).rstrip("0").rstrip(".")


# ---------------------------------------------------------------------
# Offset row
# ---------------------------------------------------------------------

class _OffsetRow(object):
    """One ``led.offsets[*]`` entry."""

    def __init__(self, data):
        self._data = data

    @property
    def XText(self):
        return _fmt_float(self._data.get("x_inches"))

    @XText.setter
    def XText(self, value):
        self._data["x_inches"] = _to_float(value, 0.0)

    @property
    def YText(self):
        return _fmt_float(self._data.get("y_inches"))

    @YText.setter
    def YText(self, value):
        self._data["y_inches"] = _to_float(value, 0.0)

    @property
    def ZText(self):
        return _fmt_float(self._data.get("z_inches"))

    @ZText.setter
    def ZText(self, value):
        self._data["z_inches"] = _to_float(value, 0.0)

    @property
    def RotText(self):
        return _fmt_float(self._data.get("rotation_deg"))

    @RotText.setter
    def RotText(self, value):
        self._data["rotation_deg"] = _to_float(value, 0.0)


# ---------------------------------------------------------------------
# Annotation row
# ---------------------------------------------------------------------

class _AnnotationRow(object):
    """One ``led.annotations[*]`` entry.

    The Family:Type dropdown is *kind-aware* — the option list shifts
    based on the row's current ``Kind``:

      * ``tag``       — annotation FamilySymbols (excluding the
                        keynote family).
      * ``keynote``   — only types of family ``GA_Keynote Symbol_CED``.
      * ``text_note`` — TextNoteType names (no Family prefix; the
                        ``Label`` getter renders just the type name).

    The per-row dynamic list is plumbed via ``options_by_kind``
    (a ``{kind: CLR_list}`` map) and ``lookups_by_kind`` (the
    parallel ``{kind: {label: info_dict}}`` for free-text resolution).
    The dialog controller builds these once and shares them across
    every row.
    """

    def __init__(self, data, options_by_kind=None, lookups_by_kind=None,
                 doc=None):
        self._data = data
        self.KindOptions = _ANNOTATION_KIND_OPTIONS_NET
        self._options_by_kind = options_by_kind or {}
        self._lookups_by_kind = lookups_by_kind or {}
        # Doc reference so the Label setter can pull symbol-default
        # parameter values when the user picks a known Family:Type.
        # Without this, annotations created in the editor start with
        # an empty ``parameters`` dict and the placed instances show
        # default content (keynotes display 'XX' etc.). When set, the
        # setter mirrors what ``Auto-fill from family`` does for LED
        # parameters but inline on Family:Type selection.
        self._doc = doc

    @property
    def Kind(self):
        return self._data.get("kind") or "tag"

    @Kind.setter
    def Kind(self, value):
        # Coerce to Python str so the membership check against
        # _ANNOTATION_KINDS (Python tuple of str) works even when
        # WPF writes back a CLR System.String.
        new_kind = str(value or "").strip()
        if new_kind in _ANNOTATION_KINDS:
            self._data["kind"] = new_kind

    def options_for_current_kind(self):
        """Return the CLR list of valid Family:Type options for this
        row's current ``Kind``. Used by the per-row resolver passed
        to ``attach_per_row_handlers``."""
        return self._options_by_kind.get(self.Kind) or _make_clr_string_list(())

    @property
    def Label(self):
        # text_note rows show the TextNoteType name in the Family:Type
        # cell — text notes have no FamilySymbol, so a "family : type"
        # composite would be misleading. The actual text *content*
        # stays in the schema's ``label`` field (set during capture)
        # and is editable via the per-row Params... button alongside
        # the rest of the annotation parameters.
        kind = self._data.get("kind") or "tag"
        if kind == "text_note":
            return self._data.get("type_name") or ""
        family = self._data.get("family_name") or ""
        type_name = self._data.get("type_name") or ""
        if family and type_name:
            return "{} : {}".format(family, type_name)
        # Fall back to a free-text label that's been hand-entered for
        # data that pre-dates the Family:Type pattern.
        return self._data.get("label") or ""

    @Label.setter
    def Label(self, value):
        # Coerce to Python str so the lookup dict (keyed by Python
        # strings) hashes correctly.
        new_label = str(value or "").strip()
        kind = self._data.get("kind") or "tag"
        lookup = self._lookups_by_kind.get(kind) or {}

        # Match against the kind's own index first.
        info = lookup.get(new_label)
        if info:
            old_label = self._data.get("label") or ""
            self._data["family_name"] = info.get("family_name") or ""
            self._data["type_name"] = info.get("type_name") or ""
            # For text_note rows, keep ``label`` as the captured text
            # content; only family/type change with this combo. For
            # tag / keynote, ``label`` mirrors the Family:Type display.
            if kind != "text_note":
                self._data["label"] = new_label
            # Auto-populate the annotation's parameters from the
            # chosen family — only when (a) we have a doc to query,
            # (b) the parameters dict is currently empty (don't blow
            # away user-typed values), and (c) the label actually
            # changed (not a no-op set). Skipped for text_note since
            # TextNoteType has no FamilySymbol parameter set.
            if (kind in ("tag", "keynote") and self._doc is not None
                    and new_label != old_label):
                params = self._data.setdefault("parameters", {})
                if not params:
                    self._auto_populate_parameters(new_label)
            return

        # Free-text fallback paths.
        if kind == "text_note":
            # Whole input is the TextNoteType name.
            self._data["family_name"] = ""
            self._data["type_name"] = new_label
            return
        if " : " in new_label:
            family, type_name = new_label.split(" : ", 1)
            self._data["family_name"] = family.strip()
            self._data["type_name"] = type_name.strip()
            self._data["label"] = new_label
        else:
            self._data["family_name"] = ""
            self._data["type_name"] = ""
            self._data["label"] = new_label

    def _auto_populate_parameters(self, label):
        """Pull defaults from the chosen Family:Type's symbol and seed
        the annotation's ``parameters`` dict.

        Mirrors what the LED Parameters tab does on ``Auto-fill from
        family`` — symbol + any-instance values, written into the
        in-memory dict so downstream YAML save + Place Element
        Annotations see real values instead of an empty mapping
        (which is the root cause of keynotes placing as 'XX').
        """
        if self._doc is None or not label:
            return
        try:
            symbol = _sym_index.find_symbol_by_label(self._doc, label)
        except Exception:
            symbol = None
        if symbol is None:
            return
        try:
            defaults = _sym_index.symbol_parameter_defaults(self._doc, symbol)
        except Exception:
            defaults = {}
        params = self._data.setdefault("parameters", {})
        for name, value in defaults.items():
            # Don't overwrite anything already set; ``params`` is
            # empty at first call but defensive in case of re-entry.
            if name in params:
                continue
            params[name] = value

    def _offset_dict(self):
        # Annotation offsets are a single dict, not a list.
        d = self._data.setdefault("offsets", {})
        if isinstance(d, list):
            d = d[0] if d else {}
            self._data["offsets"] = d
        if not isinstance(d, dict):
            d = {}
            self._data["offsets"] = d
        return d

    @property
    def OffsetXText(self):
        return _fmt_float(self._offset_dict().get("x_inches"))

    @OffsetXText.setter
    def OffsetXText(self, value):
        self._offset_dict()["x_inches"] = _to_float(value, 0.0)

    @property
    def OffsetYText(self):
        return _fmt_float(self._offset_dict().get("y_inches"))

    @OffsetYText.setter
    def OffsetYText(self, value):
        self._offset_dict()["y_inches"] = _to_float(value, 0.0)

    @property
    def OffsetZText(self):
        return _fmt_float(self._offset_dict().get("z_inches"))

    @OffsetZText.setter
    def OffsetZText(self, value):
        self._offset_dict()["z_inches"] = _to_float(value, 0.0)

    @property
    def OffsetRotText(self):
        return _fmt_float(self._offset_dict().get("rotation_deg"))

    @OffsetRotText.setter
    def OffsetRotText(self, value):
        self._offset_dict()["rotation_deg"] = _to_float(value, 0.0)


# ---------------------------------------------------------------------
# Generic key/value sub-dialog
# ---------------------------------------------------------------------

class KeyValueDialog(object):
    """Modal: edit a flat string -> string dict.

    ``doc`` + ``label`` enable the "Auto-fill from family" button —
    the sub-dialog can resolve the given Family:Type label to a
    FamilySymbol and seed parameter rows from its current values
    (same logic the LED Parameters tab uses). When either is missing
    the auto-fill button is disabled.
    """

    def __init__(self, params_dict, header="Edit Parameters",
                 doc=None, label=None):
        self._params = params_dict if isinstance(params_dict, dict) else {}
        self._doc = doc
        self._autofill_label = label or ""
        self.window = _wpf.load_xaml(_KV_XAML)
        self._rows = ObservableCollection[_NetObject]()
        self._committed = False

        f = self.window.FindName
        self.header_label = f("HeaderLabel")
        self.grid = f("ParamGrid")
        self.autofill_btn = f("AutoFillButton")
        self.add_btn = f("AddRowButton")
        self.del_btn = f("DeleteRowButton")
        self.ok_btn = f("OkButton")
        self.cancel_btn = f("CancelButton")
        self.header_label.Text = header
        self.grid.ItemsSource = self._rows
        self._snapshot = copy.deepcopy(self._params)

        for k, v in self._params.items():
            self._rows.Add(_ParamRow(k, v))

        self._h_autofill = RoutedEventHandler(lambda s, e: self._on_autofill())
        self._h_add = RoutedEventHandler(lambda s, e: self._on_add())
        self._h_del = RoutedEventHandler(lambda s, e: self._on_delete())
        self._h_ok = RoutedEventHandler(lambda s, e: self._on_ok())
        self._h_cancel = RoutedEventHandler(lambda s, e: self._on_cancel())
        self.autofill_btn.Click += self._h_autofill
        self.add_btn.Click += self._h_add
        self.del_btn.Click += self._h_del
        self.ok_btn.Click += self._h_ok
        self.cancel_btn.Click += self._h_cancel

        if self._doc is None or not self._autofill_label:
            self.autofill_btn.IsEnabled = False

    def _on_autofill(self):
        if self._doc is None or not self._autofill_label:
            return
        symbol = _sym_index.find_symbol_by_label(
            self._doc, self._autofill_label,
        )
        if symbol is None:
            return
        existing = set()
        for row in self._rows:
            n = (getattr(row, "Name", "") or "").strip()
            if n:
                existing.add(n)
        defaults = _sym_index.symbol_parameter_defaults(self._doc, symbol)
        for name in sorted(defaults.keys(), key=lambda s: s.lower()):
            if name in existing:
                continue
            self._rows.Add(_ParamRow(name, defaults[name]))

    def _on_add(self):
        self._rows.Add(_ParamRow("", ""))
        self.grid.SelectedItem = self._rows[self._rows.Count - 1]

    def _on_delete(self):
        sel = self.grid.SelectedItem
        if isinstance(sel, _ParamRow):
            self._rows.Remove(sel)

    def _on_ok(self):
        try:
            self.grid.CommitEdit()
            self.grid.CommitEdit()
        except Exception:
            pass
        # Rebuild the source dict from the rows. Preserve order.
        new_data = {}
        for row in self._rows:
            name = (row.Name or "").strip()
            if not name:
                continue
            original = self._snapshot.get(name)
            new_data[name] = _coerce_from_text(name, row.Value, original)
        # Mutate the caller's dict in place (so references survive).
        self._params.clear()
        self._params.update(new_data)
        self._committed = True
        self.window.Close()

    def _on_cancel(self):
        # Restore the snapshot — caller's dict is mutated only on OK.
        self._params.clear()
        self._params.update(self._snapshot)
        self._committed = False
        self.window.Close()

    def show_modal(self, owner=None):
        if owner is not None:
            try:
                self.window.Owner = owner
            except Exception:
                pass
        self.window.ShowDialog()
        return self._committed


# ---------------------------------------------------------------------
# Main details dialog
# ---------------------------------------------------------------------

class SpaceLedDetailsController(object):
    """Edit one LED's parameters / offsets / annotations dicts.

    Mutates the passed-in LED dict in place. Parent caller (Manage
    Space Profiles) decides whether to persist to YAML.

    ``doc`` powers the parameter auto-fill button (resolves the LED's
    Family:Type to a FamilySymbol, lists its parameter names) and the
    annotation Label dropdown (lists every loaded annotation
    Family:Type). Both are no-ops when ``doc`` is None.
    """

    def __init__(self, led_dict, header="", doc=None, led_label=None):
        self._led = led_dict if isinstance(led_dict, dict) else {}
        # Take a deep snapshot so Cancel can fully restore.
        self._snapshot = copy.deepcopy(self._led)
        self._committed = False
        self._doc = doc
        self._led_label_for_autofill = led_label or self._led.get("label") or ""

        # Build per-kind Family:Type indexes. Each kind gets its own
        # CLR list + lookup, shared across all annotation rows. The
        # row-level resolver picks the right list based on the row's
        # current Kind, and re-resolves on Kind selection change.
        self._ann_options_by_kind = {
            "tag": _make_clr_string_list(()),
            "keynote": _make_clr_string_list(()),
            "text_note": _make_clr_string_list(()),
        }
        self._ann_lookups_by_kind = {
            "tag": {},
            "keynote": {},
            "text_note": {},
        }
        if doc is not None:
            for kind, builder in (
                ("tag", _sym_index.build_tag_symbol_index),
                ("keynote", _sym_index.build_keynote_symbol_index),
                ("text_note", _sym_index.build_text_note_type_index),
            ):
                try:
                    labels, lookup = builder(doc)
                except Exception:
                    labels, lookup = [], {}
                self._ann_options_by_kind[kind] = _make_clr_string_list(labels)
                self._ann_lookups_by_kind[kind] = lookup

        self.window = _wpf.load_xaml(_DETAILS_XAML)
        self._param_rows = ObservableCollection[_NetObject]()
        self._offset_rows = ObservableCollection[_NetObject]()
        self._ann_rows = ObservableCollection[_NetObject]()

        self._lookup_controls()
        self._wire_events()
        self.header_label.Text = header or "Edit LED: {} ({})".format(
            self._led.get("label") or "(no label)",
            self._led.get("id") or "?",
        )
        self._reload_all_tabs()
        self._set_status("Ready.")

    # ----- bootstrapping -------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.header_label = f("HeaderLabel")
        self.tabs = f("DetailTabs")
        self.status_label = f("StatusLabel")
        self.ok_btn = f("OkButton")
        self.cancel_btn = f("CancelButton")

        # Parameters tab
        self.param_grid = f("ParamGrid")
        self.param_add_btn = f("ParamAddButton")
        self.param_del_btn = f("ParamDeleteButton")
        self.param_autofill_btn = f("ParamAutoFillButton")
        self.param_grid.ItemsSource = self._param_rows

        # Offsets tab
        self.offset_grid = f("OffsetGrid")
        self.offset_add_btn = f("OffsetAddButton")
        self.offset_del_btn = f("OffsetDeleteButton")
        self.offset_grid.ItemsSource = self._offset_rows

        # Annotations tab
        self.ann_grid = f("AnnGrid")
        self.ann_add_btn = f("AnnAddButton")
        self.ann_del_btn = f("AnnDeleteButton")
        self.ann_grid.ItemsSource = self._ann_rows

    def _wire_events(self):
        self._h_param_add = RoutedEventHandler(lambda s, e: self._safe(self._on_param_add, "param-add"))
        self._h_param_del = RoutedEventHandler(lambda s, e: self._safe(self._on_param_delete, "param-del"))
        self._h_offset_add = RoutedEventHandler(lambda s, e: self._safe(self._on_offset_add, "offset-add"))
        self._h_offset_del = RoutedEventHandler(lambda s, e: self._safe(self._on_offset_delete, "offset-del"))
        self._h_ann_add = RoutedEventHandler(lambda s, e: self._safe(self._on_ann_add, "ann-add"))
        self._h_ann_del = RoutedEventHandler(lambda s, e: self._safe(self._on_ann_delete, "ann-del"))
        self._h_ok = RoutedEventHandler(lambda s, e: self._safe(self._on_ok, "ok"))
        self._h_cancel = RoutedEventHandler(lambda s, e: self._safe(self._on_cancel, "cancel"))

        self.param_add_btn.Click += self._h_param_add
        self.param_del_btn.Click += self._h_param_del
        self._h_param_autofill = RoutedEventHandler(
            lambda s, e: self._safe(self._on_param_autofill, "param-autofill")
        )
        self.param_autofill_btn.Click += self._h_param_autofill
        # Disable the auto-fill button when we have no doc to query.
        if self._doc is None or not self._led_label_for_autofill:
            self.param_autofill_btn.IsEnabled = False
        self.offset_add_btn.Click += self._h_offset_add
        self.offset_del_btn.Click += self._h_offset_del
        self.ann_add_btn.Click += self._h_ann_add
        self.ann_del_btn.Click += self._h_ann_del
        self.ok_btn.Click += self._h_ok
        self.cancel_btn.Click += self._h_cancel

        # Per-row wiring (combined): Click on Params... button +
        # ItemsSource on the per-row ComboBoxes, in a single Loaded
        # handler. Doing them separately interferes (the combo
        # ItemsSource pass invalidates the button Click attach).
        # The Kind combo always shows the same three kinds. The
        # Label combo varies per row based on Kind — resolved by
        # ``_resolve_ann_combo_items`` against the row's current
        # state. ``attach_per_row_handlers`` also subscribes to
        # SelectionChanged on every combo in each row so changing
        # Kind triggers a re-resolution of the Label combo's items.
        self._row_handles = _wpf.attach_per_row_handlers(
            self.ann_grid,
            on_button_click=lambda btn, e, item: self._safe(
                lambda: self._on_ann_params_clicked(item), "row-click",
            ),
            items_per_combo_name={
                "AnnKindCombo": _ANNOTATION_KIND_OPTIONS_NET,
            },
            combo_item_resolver=self._resolve_ann_combo_items,
        )

    def _resolve_ann_combo_items(self, combo_name, row_item):
        """Per-row ItemsSource resolver passed to ``attach_per_row_handlers``.

        ``AnnLabelCombo`` (the Family:Type cell) gets a kind-aware
        list — annotation symbols when the row's Kind is ``tag``,
        keynote symbols only when ``keynote``, TextNoteType names
        when ``text_note``. Other combos return ``None`` so the
        helper falls back to the static ``items_per_combo_name`` map.
        """
        if combo_name != "AnnLabelCombo":
            return None
        if not isinstance(row_item, _AnnotationRow):
            return None
        return row_item.options_for_current_kind()

    def _safe(self, fn, label):
        try:
            fn()
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _safe_with(self, sender, e, fn, label):
        try:
            fn(sender, e)
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _set_status(self, text):
        self.status_label.Text = text or ""

    # ----- reload --------------------------------------------------

    def _reload_all_tabs(self):
        self._reload_params()
        self._reload_offsets()
        self._reload_annotations()

    def _reload_params(self):
        self._param_rows.Clear()
        params = self._led.setdefault("parameters", {})
        if not isinstance(params, dict):
            params = {}
            self._led["parameters"] = params
        for k, v in params.items():
            self._param_rows.Add(_ParamRow(k, v))

    def _reload_offsets(self):
        self._offset_rows.Clear()
        offsets = self._led.setdefault("offsets", [])
        if not isinstance(offsets, list):
            offsets = []
            self._led["offsets"] = offsets
        for o in offsets:
            if isinstance(o, dict):
                self._offset_rows.Add(_OffsetRow(o))

    def _reload_annotations(self):
        self._ann_rows.Clear()
        anns = self._led.setdefault("annotations", [])
        if not isinstance(anns, list):
            anns = []
            self._led["annotations"] = anns
        for a in anns:
            if isinstance(a, dict):
                a.setdefault("kind", "tag")
                a.setdefault("id", _new_id("ANN"))
                self._ann_rows.Add(_AnnotationRow(
                    a,
                    options_by_kind=self._ann_options_by_kind,
                    lookups_by_kind=self._ann_lookups_by_kind,
                    doc=self._doc,
                ))

    # ----- parameter actions ---------------------------------------

    def _on_param_add(self):
        self._param_rows.Add(_ParamRow("", ""))
        self.param_grid.SelectedItem = self._param_rows[self._param_rows.Count - 1]

    def _on_param_delete(self):
        sel = self.param_grid.SelectedItem
        if isinstance(sel, _ParamRow):
            self._param_rows.Remove(sel)

    def _on_param_autofill(self):
        """Pull every parameter name visible on the LED's chosen
        Family:Type and add a blank row for each one not already
        present. Existing rows (with values the user typed) are
        preserved untouched.
        """
        label = self._led_label_for_autofill or ""
        if not label or self._doc is None:
            self._set_status("No Family:Type set on this LED — can't auto-fill.")
            return

        symbol = _sym_index.find_symbol_by_label(self._doc, label)
        if symbol is None:
            self._set_status(
                "Couldn't find loaded family for {!r}. Load it first.".format(label)
            )
            return

        existing = set()
        for row in self._param_rows:
            n = (getattr(row, "Name", "") or "").strip()
            if n:
                existing.add(n)

        defaults = _sym_index.symbol_parameter_defaults(self._doc, symbol)
        added = 0
        # Sort case-insensitively so the rows come in alphabetically,
        # matching the way the previous (names-only) auto-fill ordered
        # them.
        for name in sorted(defaults.keys(), key=lambda s: s.lower()):
            if name in existing:
                continue
            self._param_rows.Add(_ParamRow(name, defaults[name]))
            added += 1
        self._set_status(
            "Auto-fill from '{}': added {} parameter row(s) with current values.".format(
                label, added,
            )
        )

    # ----- offset actions ------------------------------------------

    def _on_offset_add(self):
        new = {
            "x_inches": 0.0, "y_inches": 0.0,
            "z_inches": 0.0, "rotation_deg": 0.0,
        }
        offsets = self._led.setdefault("offsets", [])
        offsets.append(new)
        self._offset_rows.Add(_OffsetRow(new))
        self.offset_grid.SelectedItem = self._offset_rows[self._offset_rows.Count - 1]

    def _on_offset_delete(self):
        sel = self.offset_grid.SelectedItem
        if not isinstance(sel, _OffsetRow):
            return
        try:
            self._led.get("offsets", []).remove(sel._data)
        except ValueError:
            pass
        self._offset_rows.Remove(sel)

    # ----- annotation actions --------------------------------------

    def _on_ann_add(self):
        new = {
            "id": _new_id("ANN"),
            "kind": "tag",
            "label": "",
            "family_name": "",
            "type_name": "",
            "parameters": {},
            "offsets": {"x_inches": 0.0, "y_inches": 0.0,
                        "z_inches": 0.0, "rotation_deg": 0.0},
        }
        anns = self._led.setdefault("annotations", [])
        anns.append(new)
        self._ann_rows.Add(_AnnotationRow(
            new,
            options_by_kind=self._ann_options_by_kind,
            lookups_by_kind=self._ann_lookups_by_kind,
            doc=self._doc,
        ))
        self.ann_grid.SelectedItem = self._ann_rows[self._ann_rows.Count - 1]

    def _on_ann_delete(self):
        sel = self.ann_grid.SelectedItem
        if not isinstance(sel, _AnnotationRow):
            return
        try:
            self._led.get("annotations", []).remove(sel._data)
        except ValueError:
            pass
        self._ann_rows.Remove(sel)

    def _on_ann_params_clicked(self, row):
        # Click handler attached directly to each per-row button by
        # ``wpf.attach_per_row_handlers``. ``row`` is the button's
        # DataContext (the _AnnotationRow).
        if not isinstance(row, _AnnotationRow):
            return
        params = row._data.setdefault("parameters", {})
        if not isinstance(params, dict):
            params = {}
            row._data["parameters"] = params
        # Forward the annotation's Family:Type label so the sub-dialog
        # can offer "Auto-fill from family" for tags / keynote symbols.
        # Keynotes (typically family "GA_Keynote Symbol_CED") resolve
        # here exactly like any other tag — they're just FamilySymbols.
        ann_label = row.Label or ""
        kind = row._data.get("kind") or "tag"

        # Text notes carry their content in a top-level ``text`` field
        # (not in ``parameters``). The placement engine reads it from
        # there when creating the TextNote. Splice a ``Text`` row into
        # the combined dict the dialog edits, then on commit pull it
        # back out and write to the annotation's top-level ``text``.
        is_text_note = (kind == "text_note")
        if is_text_note:
            # Existing data may have come from an older capture where
            # the content was stored in ``label`` (per
            # hosted_annotations.annotation_descriptor). Prefer
            # top-level ``text``; fall back to ``label`` for legacy
            # entries so users see their existing content prepopulated
            # rather than a blank cell.
            existing_text = (
                row._data.get("text")
                or row._data.get("label")
                or ""
            )
            combined = {"Text": existing_text}
            for k, v in params.items():
                if k == "Text":
                    # Avoid the collision — the params-dict version is
                    # ignored in favour of the top-level field.
                    continue
                combined[k] = v
            dialog_data = combined
        else:
            dialog_data = params

        dialog = KeyValueDialog(
            dialog_data,
            header="Annotation parameters: {} [{}]".format(
                ann_label or row._data.get("type_name") or "(no label)",
                kind,
            ),
            doc=self._doc,
            label=ann_label,
        )
        committed = dialog.show_modal(owner=self.window)

        if is_text_note and committed:
            # Split the combined dict back: ``Text`` -> top-level field,
            # everything else stays in ``parameters``.
            text_value = combined.pop("Text", "")
            if text_value is None:
                text_value = ""
            row._data["text"] = str(text_value)
            params.clear()
            params.update(combined)
        # Intentionally NOT calling self.ann_grid.Items.Refresh() —
        # the parameters dict isn't visible in any annotation grid
        # column, and forcing a refresh rebuilds row containers
        # without firing Loaded, which clobbers our programmatically
        # set ItemsSource on the Kind / Label combos and leaves them
        # blank.

    # ----- OK / Cancel ---------------------------------------------

    def _on_ok(self):
        # Commit any in-flight DataGrid edit so the last-typed cell makes it.
        for grid in (self.param_grid, self.offset_grid, self.ann_grid):
            try:
                grid.CommitEdit()
                grid.CommitEdit()
            except Exception:
                pass

        # Rebuild parameters from rows.
        params_out = {}
        original_params = self._snapshot.get("parameters") or {}
        for row in self._param_rows:
            name = (row.Name or "").strip()
            if not name:
                continue
            params_out[name] = _coerce_from_text(
                name, row.Value, original_params.get(name),
            )
        self._led["parameters"] = params_out

        # Offsets are already mutated in place via the row setters,
        # but rebuild the list from the current rows (handles deletes
        # and any reorder if we ever add it).
        self._led["offsets"] = [r._data for r in self._offset_rows]

        # Keep ``Elevation from Level`` (parameter) and ``offsets[0]
        # .z_inches`` (geometric offset) in lockstep. Both represent
        # the fixture's height above the space's level — having two
        # sources of truth for the same value lets them drift, with
        # the user editing one and being surprised the other "wins"
        # at placement / audit time. The parameter is the user-facing
        # control (it's how engineers think about elevation), so when
        # it changes, propagate the change into the offset's z_inches.
        # We also feed the parameter back from the offset when the
        # parameter is missing/blank — so authors who only set z_inches
        # still get a parameter value that reads correctly in the
        # Parameters tab on next open.
        self._reconcile_elevation_from_level(params_out)

        # Annotations same — mutated in place via setters, just rebuild
        # the list from current rows.
        self._led["annotations"] = [r._data for r in self._ann_rows]

        self._committed = True
        self.window.Close()

    def _reconcile_elevation_from_level(self, params_out):
        """Keep ``params['Elevation from Level']`` in lockstep with
        ``offsets[0].z_inches``, derived from the offset value.

        Equipment-side LEDs carry "Elevation from Level" inside their
        ``parameters`` dict (captured from the source instance), and
        placement writes it back via ``_apply_static_parameters`` ->
        ``SetValueString``. Spaces don't capture, so we have to
        synthesize the parameter value from the user-edited
        ``z_inches`` so the placement code path can be byte-for-byte
        identical to equipment.

        Authoritative direction: offset -> parameter. Whatever
        ``z_inches`` holds becomes the parameter's display string. If
        the user typed something different in the Parameters tab and
        ``z_inches`` is non-zero, ``z_inches`` wins (the Offsets tab
        is the source of truth, like equipment). If ``z_inches`` is
        zero AND the user typed a value into the Parameters tab,
        migrate that value into ``z_inches`` first so the dialog
        round-trips cleanly.
        """
        if not isinstance(params_out, dict):
            return
        offsets = self._led.setdefault("offsets", [])
        if not isinstance(offsets, list):
            offsets = []
            self._led["offsets"] = offsets

        # Read current values.
        z_inches_val = 0.0
        if offsets and isinstance(offsets[0], dict):
            try:
                z_inches_val = float(offsets[0].get("z_inches") or 0.0)
            except (TypeError, ValueError):
                z_inches_val = 0.0
        param_value = params_out.get("Elevation from Level")

        # If z_inches is zero but the parameter has a value, migrate
        # the parameter value into z_inches so the offset reflects it.
        if abs(z_inches_val) < 1e-6 and param_value not in (None, ""):
            try:
                import placement as _placement
                feet = _placement._parse_feet_inches(param_value)
            except Exception:
                feet = None
            if feet is not None:
                z_inches_val = float(feet) * 12.0
                if not offsets or not isinstance(offsets[0], dict):
                    offsets.insert(0, {
                        "x_inches": 0.0,
                        "y_inches": 0.0,
                        "z_inches": z_inches_val,
                        "rotation_deg": 0.0,
                    })
                else:
                    offsets[0]["z_inches"] = z_inches_val

        # Write the canonical "Elevation from Level" value into params
        # from z_inches, formatted as a Revit feet-inches string. This
        # is what placement (``_apply_static_parameters`` ->
        # ``SetValueString``) writes onto the placed instance —
        # mirroring exactly how equipment-side captured values land.
        feet = int(z_inches_val // 12) if z_inches_val >= 0 else -int(-z_inches_val // 12)
        inches = z_inches_val - feet * 12
        if abs(inches - round(inches)) < 1e-6:
            inches_str = '{}"'.format(int(round(inches)))
        else:
            inches_str = '{:g}"'.format(inches)
        params_out["Elevation from Level"] = "{}' - {}".format(
            feet, inches_str,
        )
        self._led["parameters"] = params_out

    def _on_cancel(self):
        # Restore the snapshot fully — every nested dict / list.
        self._led.clear()
        self._led.update(copy.deepcopy(self._snapshot))
        self._committed = False
        self.window.Close()

    # ----- entry point --------------------------------------------

    def show_modal(self, owner=None):
        if owner is not None:
            try:
                self.window.Owner = owner
            except Exception:
                pass
        self.window.ShowDialog()
        return self._committed


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _new_id(prefix):
    return "{}-{}".format(prefix, uuid.uuid4().hex[:8].upper())


def show_modal(led_dict, header="", owner=None, doc=None, led_label=None):
    """Open the Details dialog for ``led_dict``. Returns True on OK.

    ``doc`` enables the parameter auto-fill button and the annotation
    Family:Type dropdown. ``led_label`` is the LED's "Family : Type"
    label (used as the auto-fill source); falls back to
    ``led_dict['label']`` when omitted.
    """
    controller = SpaceLedDetailsController(
        led_dict=led_dict, header=header,
        doc=doc, led_label=led_label,
    )
    return controller.show_modal(owner=owner)
