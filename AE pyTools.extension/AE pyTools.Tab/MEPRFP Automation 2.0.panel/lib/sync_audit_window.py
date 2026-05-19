# -*- coding: utf-8 -*-
"""
Tree-view UI for the synced-relationship audit.

Hierarchy: Profile -> LED -> Conflict. This is a **flag-only** report —
each conflict row shows the directive source value, the child's actual
value, and the percentage difference between them. There is no
correction path (no per-row action combo, no Apply); the only verbs are
Refresh and Close.
"""

import os

import clr  # noqa: F401

from System.Windows.Controls import (  # noqa: E402
    Grid,
    ColumnDefinition,
    TextBlock,
    TreeViewItem,
)
from System.Windows import GridLength, GridUnitType, Thickness  # noqa: E402

import sync_audit
import wpf as _wpf


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_resources", "SyncAuditWindow.xaml"
)


class SyncAuditController(object):

    def __init__(self, doc, profile_data):
        self.doc = doc
        self.profile_data = profile_data
        # Cache of the last detected conflicts so the % filter / sort
        # controls can re-render WITHOUT re-running detect_conflicts
        # (which collects + reads the model — wasteful per keystroke,
        # and the window is modeless so we avoid touching the API
        # outside an explicit Refresh).
        self._all_conflicts = []
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._wire_controls()
        self.refresh()

    # ---- wiring ------------------------------------------------------

    def _wire_controls(self):
        f = self.window.FindName
        self.tree = f("ConflictTree")
        self.summary = f("SummaryLabel")
        self.detail_box = f("DetailBox")
        self.min_pct_box = f("MinPctBox")
        self.sort_combo = f("SortCombo")
        # Retain handler refs so pythonnet's GC can't drop them.
        self._h_refresh = lambda s, e: self._on_refresh(s, e)
        self._h_close = lambda s, e: self._on_close(s, e)
        self._h_tree_select = lambda s, e: self._on_tree_select(s, e)
        self._h_filter_sort = lambda s, e: self._on_filter_sort_changed(s, e)
        f("RefreshButton").Click += self._h_refresh
        f("CloseButton").Click += self._h_close
        self.tree.SelectedItemChanged += self._h_tree_select
        # Filter / sort re-render from the CACHED conflicts only —
        # never re-detects, so it's cheap per-keystroke and modeless-safe.
        self.min_pct_box.TextChanged += self._h_filter_sort
        self.sort_combo.SelectionChanged += self._h_filter_sort

    # ---- public ------------------------------------------------------

    def show_modeless(self):
        """Show the window non-modally so the user can keep working in
        Revit (pan/zoom, select elements) with the audit report open.

        Safe without an ExternalEvent gateway because this audit is
        read-only — ``detect_conflicts`` only collects + reads
        parameters, never opens a transaction (same precondition the
        circuit-audit workflow documents for its own modeless use).
        The first ``refresh()`` already ran in ``__init__`` while the
        pushbutton was still in the Revit API context, so the tree is
        populated before the window is ever shown. The defensive
        ``Loaded`` hook only re-collects if that first pass somehow
        produced nothing (mirrors ``circuit_window.show_modeless``).
        """
        self._h_loaded = lambda s, e: self._on_window_loaded(s, e)
        try:
            self.window.Loaded += self._h_loaded
        except Exception:
            pass
        self.window.Show()
        return self

    def _on_window_loaded(self, sender, e):
        try:
            if self.tree is not None and self.tree.Items.Count == 0:
                self.refresh()
        except Exception as exc:
            self._set_status_safe(
                "Loaded-event refresh failed: {}".format(exc)
            )

    def _set_status_safe(self, text):
        try:
            self.summary.Text = text or ""
        except Exception:
            pass

    def refresh(self):
        """Re-run detection (reads the model) and cache the result,
        then render through the active filter / sort."""
        self._all_conflicts = sync_audit.detect_conflicts(
            self.doc, self.profile_data
        )
        self._render()

    # ---- filter / sort ----------------------------------------------

    def _min_pct(self):
        """Parsed ``Min % diff`` box, or ``None`` when blank / invalid."""
        try:
            txt = (self.min_pct_box.Text or "").strip().rstrip("%").strip()
        except Exception:
            txt = ""
        if not txt:
            return None
        try:
            return float(txt)
        except ValueError:
            return None

    def _sort_mode(self):
        # 0 = Grouped (default), 1 = % high->low, 2 = % low->high
        try:
            return self.sort_combo.SelectedIndex
        except Exception:
            return 0

    def _filtered_sorted(self, conflicts):
        """Apply the % filter then the sort. A non-numeric conflict
        (``percent_difference is None``) has no measurable %, so it is
        dropped by any positive Min-% filter and always sorts AFTER
        numeric ones regardless of direction."""
        min_pct = self._min_pct()
        out = []
        for c in conflicts:
            pd = c.percent_difference
            if min_pct is not None and min_pct > 0:
                if pd is None or pd < min_pct:
                    continue
            out.append(c)
        mode = self._sort_mode()
        if mode in (1, 2):
            high_low = (mode == 1)

            def _key(c):
                pd = c.percent_difference
                numeric_last = 0 if pd is not None else 1
                v = pd if pd is not None else 0.0
                return (numeric_last, -v if high_low else v)

            out = sorted(out, key=_key)
        return out, min_pct, mode

    def _render(self):
        conflicts, min_pct, mode = self._filtered_sorted(self._all_conflicts)
        self._populate_tree(conflicts)
        total = len(self._all_conflicts)
        shown = len(conflicts)
        bits = []
        if shown == total:
            bits.append("{} flagged difference(s)".format(total))
        else:
            bits.append("{} of {} flagged difference(s)".format(shown, total))
        bits.append("across {} profile(s)".format(
            len({c.profile_id for c in conflicts})
        ))
        if min_pct is not None and min_pct > 0:
            bits.append("min {:g}%".format(min_pct))
        if mode == 1:
            bits.append("sorted % high→low")
        elif mode == 2:
            bits.append("sorted % low→high")
        self.summary.Text = (
            "  ".join(bits) + "  (report only — no changes are made)"
        )

    def _on_filter_sort_changed(self, sender, e):
        # Re-render from the cache only — never re-detects.
        try:
            self._render()
        except Exception as exc:
            self._set_status_safe("Filter/sort failed: {}".format(exc))

    # ---- tree --------------------------------------------------------

    def _populate_tree(self, conflicts):
        self.tree.Items.Clear()
        # Group: profile -> led -> [conflicts]
        by_profile = {}
        for c in conflicts:
            by_profile.setdefault(
                (c.profile_id, c.profile_name), {}
            ).setdefault((c.led_id, c.led_label), []).append(c)

        for (profile_id, profile_name), led_groups in by_profile.items():
            profile_node = TreeViewItem()
            profile_node.Header = "{}  ({})  -  {} conflict(s)".format(
                profile_name, profile_id,
                sum(len(v) for v in led_groups.values()),
            )
            profile_node.IsExpanded = True
            for (led_id, led_label), led_conflicts in led_groups.items():
                led_node = TreeViewItem()
                led_node.Header = "{}  ({})  -  {} conflict(s)".format(
                    led_label, led_id, len(led_conflicts)
                )
                led_node.IsExpanded = True
                for conflict in led_conflicts:
                    led_node.Items.Add(self._build_conflict_node(conflict))
                profile_node.Items.Add(led_node)
            self.tree.Items.Add(profile_node)

    def _build_conflict_node(self, conflict):
        node = TreeViewItem()
        node.IsExpanded = False
        node.Tag = ("conflict", conflict)
        grid = Grid()
        for w in (3, 2, 2, 2):
            col = ColumnDefinition()
            col.Width = GridLength(w, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        # Compose a label that already shows what the directive references,
        # so users see the source even before they click into the detail panel.
        if conflict.kind == "parent":
            ref_text = "(parent.{})".format(conflict.target_param_name or "?")
        elif conflict.kind == "sibling":
            ref_text = "(sibling.{})".format(conflict.target_param_name or "?")
        else:
            ref_text = ""
        param = TextBlock()
        param.Text = "{}  {}".format(conflict.parameter_name, ref_text)
        param.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(param, 0)
        grid.Children.Add(param)

        actual = TextBlock()
        actual.Text = "Actual: {}".format(_short(conflict.actual_value))
        actual.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(actual, 1)
        grid.Children.Add(actual)

        expected = TextBlock()
        expected.Text = "Expected: {}".format(_short(conflict.expected_value))
        expected.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(expected, 2)
        grid.Children.Add(expected)

        # Percentage difference replaces the old resolution combo —
        # this audit is flag-only.
        pct = TextBlock()
        pct.Text = "Diff: {}".format(conflict.percent_display)
        pct.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(pct, 3)
        grid.Children.Add(pct)

        node.Header = grid
        return node

    # ---- handlers ----------------------------------------------------

    def _on_refresh(self, sender, e):
        self.refresh()
        self._render_detail(None)

    def _on_close(self, sender, e):
        self.window.Close()

    def _on_tree_select(self, sender, e):
        item = self.tree.SelectedItem
        if item is None:
            self._render_detail(None)
            return
        tag = getattr(item, "Tag", None)
        if isinstance(tag, tuple) and len(tag) >= 2 and tag[0] == "conflict":
            self._render_detail(tag[1])
        else:
            self._render_detail(None)

    # ---- detail panel -----------------------------------------------

    def _render_detail(self, conflict):
        if not hasattr(self, "detail_box") or self.detail_box is None:
            return
        if conflict is None:
            self.detail_box.Text = "Select a conflict in the tree to see its comparison."
            return

        # Resolve the actual elements involved.
        from Autodesk.Revit.DB import ElementId
        try:
            child_id = ElementId(int(conflict.element_id)) if conflict.element_id else None
        except Exception:
            child_id = None
        try:
            target_id = ElementId(int(conflict.target_element_id)) if conflict.target_element_id else None
        except Exception:
            target_id = None

        child_elem = self.doc.GetElement(child_id) if child_id else None
        target_elem = self.doc.GetElement(target_id) if target_id else None

        kind = conflict.kind or "?"
        ref_label = "parent" if kind == "parent" else ("sibling" if kind == "sibling" else kind)

        lines = []
        lines.append("SELECTED CONFLICT")
        lines.append("=================")
        lines.append("Profile:        {}  ({})".format(
            conflict.profile_name or "?", conflict.profile_id or "?"))
        lines.append("LED:            {}  ({})".format(
            conflict.led_label or "?", conflict.led_id or "?"))
        lines.append("Parameter:      {}".format(conflict.parameter_name or "?"))
        lines.append("Directive kind: {}".format(kind))
        lines.append("Expected (src): {}".format(_short(conflict.expected_value)))
        lines.append("Actual (child): {}".format(_short(conflict.actual_value)))
        lines.append("Difference:     {}".format(conflict.percent_display))
        if conflict.target_param_name:
            lines.append("Source:         {}.{}{}".format(
                ref_label,
                conflict.target_param_name,
                "  (id {})".format(conflict.target_element_id)
                if conflict.target_element_id is not None else "",
            ))
        lines.append("")

        # ---- child parameters ---------------------------------------
        lines.append("CHILD ELEMENT  (id {})".format(
            conflict.element_id if conflict.element_id is not None else "?"))
        lines.append("=" * 60)
        if child_elem is None:
            lines.append("  (element not found in the active document)")
        else:
            child_params = _collect_params(child_elem)
            if not child_params:
                lines.append("  (no parameters)")
            else:
                lines.append(_format_params(
                    child_params,
                    highlight={conflict.parameter_name: " *** mismatch ***"},
                ))
        lines.append("")

        # ---- referenced (parent / sibling) parameters ----------------
        ref_label_upper = "PARENT ELEMENT" if kind == "parent" else \
                          "SIBLING ELEMENT" if kind == "sibling" else "TARGET ELEMENT"
        lines.append("{}  (id {})".format(
            ref_label_upper,
            conflict.target_element_id
            if conflict.target_element_id is not None else "?",
        ))
        lines.append("=" * 60)
        if target_elem is None:
            lines.append("  (element not found in the active document)")
        else:
            target_params = _collect_params(target_elem)
            if not target_params:
                lines.append("  (no parameters)")
            else:
                lines.append(_format_params(
                    target_params,
                    highlight={
                        conflict.target_param_name:
                            " *** referenced by directive ***"
                    } if conflict.target_param_name else {},
                ))

        self.detail_box.Text = "\n".join(lines)


def _short(value):
    if value is None:
        return "(empty)"
    s = str(value)
    if len(s) > 40:
        return s[:37] + "..."
    return s


def _collect_params(elem):
    """Return ``[(name, value_string), ...]`` for every parameter on an
    element. Sorted alphabetically by name. Includes empty-value
    parameters so the user sees the full picture."""
    out = []
    if elem is None:
        return out
    seen = set()
    try:
        params_iter = elem.Parameters
    except Exception:
        return out
    for p in params_iter:
        if p is None:
            continue
        try:
            name = p.Definition.Name
        except Exception:
            continue
        if not name or name in seen:
            continue
        seen.add(name)
        value = None
        try:
            value = p.AsValueString()
        except Exception:
            value = None
        if value is None:
            try:
                value = p.AsString()
            except Exception:
                value = None
        out.append((name, "" if value is None else str(value)))
    out.sort(key=lambda nv: nv[0].lower())
    return out


def _format_params(name_value_pairs, highlight=None):
    """Render ``[(name, value), ...]`` as aligned text. ``highlight`` is
    ``{name: marker_string}`` — names matching get an inline marker."""
    if not name_value_pairs:
        return "  (no parameters)"
    highlight = highlight or {}
    name_width = max(len(n) for n, _ in name_value_pairs)
    name_width = min(name_width, 40)  # don't run away on weird names
    lines = []
    for name, value in name_value_pairs:
        marker = highlight.get(name, "")
        lines.append("  {n:<{w}}  {v}{m}".format(
            n=name[:name_width],
            w=name_width,
            v=value if value else "(empty)",
            m=marker,
        ))
    return "\n".join(lines)


def show_modeless(doc, profile_data):
    """Open the synced-relationship audit as a non-modal window and
    return the controller. The window stays alive after the pushbutton
    script returns the same way the SuperCircuit modeless window does —
    Revit's window manager keeps the shown ``Window`` (and its
    handler-bound controller) alive."""
    return SyncAuditController(doc, profile_data).show_modeless()
