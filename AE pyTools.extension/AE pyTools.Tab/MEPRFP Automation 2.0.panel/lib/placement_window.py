# -*- coding: utf-8 -*-
"""
The Placement dialog controller.

The dialog is a single window with three sections:

    1. Source (radio): Host model / Linked Revit model / CSV
       + a source-specific picker (combo or file browse)

    2. Filters (two list boxes): category multi-select + profile-name
       multi-select. Both default to "select none" = include all.

    3. Match preview: per-row check + (target, profile) labels. Match
       button populates this; user can toggle individual rows; Place
       button commits.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import Thickness, VerticalAlignment, Visibility  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402
from System.Windows.Controls import (  # noqa: E402
    CheckBox,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    TextBlock,
)
from System.Windows import GridUnitType, GridLength  # noqa: E402

from Autodesk.Revit.DB import (  # noqa: E402
    FilteredElementCollector,
    Level,
    Phase,
)

import fuzzy_match
import merge_workflow
import placement
import placement_apply
import wpf as _wpf
import wpf_dialogs


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "_resources", "PlacementWindow.xaml"
)


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

class _SourceItem(object):
    """Wraps a source-side selection (a RevitLinkInstance, an
    ImportInstance, or a CSV path) for display in the source combo."""

    def __init__(self, label, kind, value):
        self.label = label
        self.kind = kind
        self.value = value

    def __str__(self):
        return self.label


class _PhaseItem(object):
    """Wraps one Revit Phase for display in the phase combo. ``value`` is
    the phase's ElementId integer — kept on the item so Match doesn't have
    to re-resolve it."""

    def __init__(self, label, phase_id_int, sequence_number):
        self.label = label
        self.value = phase_id_int
        self.sequence_number = sequence_number

    def __str__(self):
        return self.label


class _MatchRow(object):
    """Per-row UI state alongside a placement.Match."""

    def __init__(self, match, ui_grid, checkbox):
        self.match = match
        self.grid = ui_grid
        self.checkbox = checkbox

    @property
    def checked(self):
        return bool(self.checkbox.IsChecked)


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class PlacementController(object):

    def __init__(self, doc, profile_data, uidoc=None, output=None):
        self.doc = doc
        self.profile_data = profile_data
        self.profiles = list(profile_data.get("equipment_definitions") or [])
        self.matches = []
        self._match_rows = []
        self._csv_path = None
        self._all_profile_labels = []        # alphabetically sorted, full list
        self._selected_profile_labels = set()  # survives search-filtering
        self._suppress_profile_selection = False
        self.committed = False
        self._last_result = None
        self._output = output
        if uidoc is None:
            try:
                from pyrevit import revit
                uidoc = getattr(revit, "uidoc", None)
            except Exception:
                uidoc = None
        self.uidoc = uidoc
        # Resolve the per-session modeless ExternalEvent gateway now,
        # while we're in a valid Revit API context (pushbutton main()).
        # The placement itself runs on Revit's main thread via this
        # gateway so the active-view switch — which binds the Level on
        # workplane-based families — is legal. Doing it from this modal-
        # free, transaction-free context is the whole fix.
        try:
            self._gateway = placement_apply.get_or_create_gateway()
        except Exception as exc:
            self._gateway = None
            self._gateway_error = str(exc)
        # Separate save gateway so accepting a fuzzy-match alias can
        # persist the profile store from this modeless window without
        # colliding with a queued placement run.
        try:
            self._save_gateway = placement_apply.get_or_create_save_gateway()
        except Exception:
            self._save_gateway = None
        self.window = _wpf.load_xaml(_XAML_PATH)
        self._lookup_controls()
        self._populate_filters()
        self._wire_events()
        self._switch_source("host_model")
        self._set_status("Pick a source, optionally filter, then Match.")

    # ---- bootstrapping ---------------------------------------------

    def _lookup_controls(self):
        f = self.window.FindName
        self.src_host_radio = f("SrcHostRadio")
        self.src_linked_radio = f("SrcLinkedRevitRadio")
        self.src_csv_radio = f("SrcCsvRadio")
        self.src_label = f("SrcLabel")
        self.src_combo = f("SrcCombo")
        self.src_browse_btn = f("SrcBrowseButton")
        self.phase_label = f("PhaseLabel")
        self.phase_combo = f("PhaseCombo")
        self.host_height_row = f("HostHeightRow")
        self.host_height_check = f("HostHeightCheck")
        self.host_height_combo = f("HostHeightCombo")
        self.category_list = f("CategoryList")
        self.profile_list = f("ProfileList")
        self.profile_search_box = f("ProfileSearchBox")
        self.skip_placed_check = f("SkipPlacedCheck")
        self.one_profile_per_target_check = f("OneProfilePerTargetCheck")
        self.allow_type_sub_check = f("AllowTypeSubCheck")
        self.match_btn = f("MatchButton")
        self.check_all_btn = f("CheckAllButton")
        self.uncheck_all_btn = f("UncheckAllButton")
        self.place_btn = f("PlaceButton")
        self.close_btn = f("CloseButton")
        self.summary_label = f("SummaryLabel")
        self.status_label = f("StatusLabel")
        self.match_rows_panel = f("MatchRowsPanel")

    def _make_delegate(self, label, fn):
        """Wrap a Python callable into a retained ``RoutedEventHandler``.

        Each handler:
            * is wrapped in try/except so a Python error becomes a status
              message instead of being silently swallowed by pythonnet,
            * writes ``[label] ...`` to the status label up front so we
              can tell from the UI whether the handler ever fires,
            * is converted to a ``RoutedEventHandler`` delegate explicitly
              (pythonnet 3's implicit conversion has been unreliable
              specifically for RoutedEventHandler in some builds),
            * is stored as an attribute so neither pythonnet nor Python's
              GC drops the delegate / target.
        """
        def wrapped(sender, e):
            try:
                self._set_status("[{}] running...".format(label))
                fn(sender, e)
            except Exception as exc:
                self._set_status("[{}] error: {}".format(label, exc))
                raise

        delegate = RoutedEventHandler(wrapped)
        return delegate

    def _wire_events(self):
        # Build + retain delegates for every event we subscribe to.
        self._h_src_host = self._make_delegate(
            "src=host", lambda s, e: self._switch_source("host_model"))
        self._h_src_linked = self._make_delegate(
            "src=linked", lambda s, e: self._switch_source("linked_revit"))
        self._h_src_csv = self._make_delegate(
            "src=csv", lambda s, e: self._switch_source("csv"))
        self._h_browse = self._make_delegate(
            "browse", lambda s, e: self._on_browse_clicked(s, e))
        self._h_match = self._make_delegate(
            "match", lambda s, e: self._on_match_clicked(s, e))
        self._h_check_all = self._make_delegate(
            "check-all", lambda s, e: self._on_check_all(s, e))
        self._h_uncheck_all = self._make_delegate(
            "uncheck-all", lambda s, e: self._on_uncheck_all(s, e))
        self._h_place = self._make_delegate(
            "place", lambda s, e: self._on_place_clicked(s, e))
        self._h_close = self._make_delegate(
            "close", lambda s, e: self.window.Close())

        self.src_host_radio.Checked += self._h_src_host
        self.src_linked_radio.Checked += self._h_src_linked
        self.src_csv_radio.Checked += self._h_src_csv
        self.src_browse_btn.Click += self._h_browse
        self.match_btn.Click += self._h_match
        self.check_all_btn.Click += self._h_check_all
        self.uncheck_all_btn.Click += self._h_uncheck_all
        self.place_btn.Click += self._h_place
        self.close_btn.Click += self._h_close

        # Profile search + selection-survival wiring. pythonnet wraps
        # the bound methods automatically; the bound-method instances are
        # kept alive by the class itself.
        self.profile_search_box.TextChanged += self._on_profile_search
        self.profile_list.SelectionChanged += self._on_profile_selection
        # Re-populate the phase combo whenever the linked-model picker
        # changes. No-op for host_model (host doc was already populated
        # in _switch_source) and CSV (no phase concept).
        self.src_combo.SelectionChanged += self._on_source_combo_changed

    # ---- source handling -------------------------------------------

    def _switch_source(self, kind):
        """``kind`` in ('host_model', 'linked_revit', 'csv')."""
        self._source_kind = kind
        self.src_combo.Items.Clear()
        if kind == "host_model":
            self.src_label.Text = "Active document"
            self.src_browse_btn.Visibility = Visibility.Collapsed
            self.src_combo.IsEnabled = False
            self.src_combo.Items.Add(
                _SourceItem("(this document)", "host_model", self.doc)
            )
            self.src_combo.SelectedIndex = 0
            self._show_phase_combo(True)
            self._populate_phase_combo(self.doc)
        elif kind == "linked_revit":
            self.src_label.Text = "Linked model:"
            self.src_browse_btn.Visibility = Visibility.Collapsed
            self.src_combo.IsEnabled = True
            self._show_phase_combo(True)
            self._populate_phase_combo(None)
            for inst in placement.collect_linked_revit_link_instances(self.doc):
                link_doc = inst.GetLinkDocument()
                title = getattr(link_doc, "Title", "") or "(unnamed)"
                self.src_combo.Items.Add(_SourceItem(title, "linked_revit", inst))
            if self.src_combo.Items.Count > 0:
                # SelectedIndex assignment fires SelectionChanged, which
                # populates the phase combo from the picked link's doc.
                self.src_combo.SelectedIndex = 0
        elif kind == "csv":
            self.src_label.Text = "CSV file:"
            self.src_browse_btn.Visibility = Visibility.Visible
            self.src_combo.IsEnabled = True
            self._show_phase_combo(False)
            self._populate_phase_combo(None)
            if self._csv_path:
                self.src_combo.Items.Add(
                    _SourceItem(self._csv_path, "csv", self._csv_path)
                )
                self.src_combo.SelectedIndex = 0
        # Host-height override is meaningful only for the host-model
        # source (fixtures placed relative to a live host parent).
        self.host_height_row.Visibility = (
            Visibility.Visible if kind == "host_model" else Visibility.Collapsed
        )
        # Clear preview when source changes.
        self._clear_match_rows()
        self._set_status("Source switched. Match again.")

    def _show_phase_combo(self, visible):
        vis = Visibility.Visible if visible else Visibility.Collapsed
        self.phase_label.Visibility = vis
        self.phase_combo.Visibility = vis

    def _populate_phase_combo(self, phase_source_doc):
        """Fill the phase combo from ``phase_source_doc``'s Phase list.

        Sort by ``SequenceNumber`` so the order matches the project's
        phase timeline. Default selection: an item named "New
        Construction" (case-insensitive); otherwise the first item.
        Pass ``None`` to clear the combo (e.g. between switching to
        linked_revit and the user picking a specific link).
        """
        self.phase_combo.Items.Clear()
        if phase_source_doc is None:
            return
        try:
            phases = list(
                FilteredElementCollector(phase_source_doc).OfClass(Phase)
            )
        except Exception:
            phases = []
        items = []
        for phase in phases:
            try:
                name = phase.Name
                eid = phase.Id
                eid_val = (
                    getattr(eid, "Value", None)
                    or getattr(eid, "IntegerValue", None)
                )
                seq = int(getattr(phase, "SequenceNumber", 0) or 0)
            except Exception:
                continue
            if eid_val is None:
                continue
            items.append(_PhaseItem(name, eid_val, seq))
        items.sort(key=lambda it: (it.sequence_number, it.label.lower()))
        for it in items:
            self.phase_combo.Items.Add(it)
        if self.phase_combo.Items.Count == 0:
            return
        default_idx = 0
        for idx in range(self.phase_combo.Items.Count):
            item = self.phase_combo.Items[idx]
            if (item.label or "").strip().lower() == "new construction":
                default_idx = idx
                break
        self.phase_combo.SelectedIndex = default_idx

    def _on_source_combo_changed(self, sender, e):
        """When the user picks a different link, re-populate the phase
        combo from that link's doc. Host_model and CSV are no-ops
        (host doc was already populated in _switch_source; CSV doesn't
        use phases)."""
        if self._source_kind != "linked_revit":
            return
        try:
            item = self.src_combo.SelectedItem
            if item is None or item.value is None:
                self._populate_phase_combo(None)
                return
            link_doc = item.value.GetLinkDocument()
            self._populate_phase_combo(link_doc)
        except Exception as exc:
            self._set_status("[phase-populate] error: {}".format(exc))

    def _on_browse_clicked(self, sender, e):
        if self._source_kind != "csv":
            return
        # Reuse forms_compat.pick_file via the wpf path.
        import forms_compat as forms
        path = forms.pick_file(file_ext="csv", title="Pick rebased-coords CSV")
        if not path:
            return
        self._csv_path = path
        self.src_combo.Items.Clear()
        self.src_combo.Items.Add(_SourceItem(path, "csv", path))
        self.src_combo.SelectedIndex = 0

    # ---- filter population -----------------------------------------

    def _populate_filters(self):
        # Categories come from each LED's ``category`` field across every
        # profile — i.e. the categories of the *fixture children*, not
        # the profile's parent. Selecting one or more categories:
        #   1. Keeps every profile that contains AT LEAST ONE LED in a
        #      selected category (profile-level filter in
        #      ``_filtered_profiles``).
        #   2. Restricts placement to LEDs whose category is in the
        #      selected set (LED-level filter applied by
        #      ``execute_placement`` via
        #      ``PlacementOptions.category_filter``).
        # So a profile mixing Data Devices + Mechanical Equipment LEDs,
        # with "Data Devices" selected, places only its Data Devices
        # LEDs — the Mechanical Equipment LEDs are skipped.
        cats = set()
        for p in self.profiles:
            if not isinstance(p, dict):
                continue
            for s in p.get("linked_sets") or []:
                if not isinstance(s, dict):
                    continue
                for led in s.get("linked_element_definitions") or []:
                    if not isinstance(led, dict):
                        continue
                    c = (led.get("category") or "").strip()
                    if c:
                        cats.add(c)
        self.category_list.Items.Clear()
        for c in sorted(cats):
            self.category_list.Items.Add(c)

        labels = []
        for p in self.profiles:
            if not isinstance(p, dict):
                continue
            labels.append("{}  ({})".format(
                p.get("name") or "(unnamed)", p.get("id") or "?"
            ))
        labels.sort(key=lambda s: s.lower())
        self._all_profile_labels = labels
        self._render_profile_list("")

    def _render_profile_list(self, search_text):
        needle = (search_text or "").strip().lower()
        self._suppress_profile_selection = True
        try:
            self.profile_list.Items.Clear()
            visible = []
            for label in self._all_profile_labels:
                if needle and needle not in label.lower():
                    continue
                self.profile_list.Items.Add(label)
                visible.append(label)
            for label in visible:
                if label in self._selected_profile_labels:
                    self.profile_list.SelectedItems.Add(label)
        finally:
            self._suppress_profile_selection = False

    def _on_profile_search(self, sender, e):
        try:
            self._render_profile_list(self.profile_search_box.Text or "")
        except Exception as exc:
            self._set_status("[search] error: {}".format(exc))

    def _on_profile_selection(self, sender, e):
        if self._suppress_profile_selection:
            return
        try:
            for item in e.AddedItems:
                self._selected_profile_labels.add(str(item))
            for item in e.RemovedItems:
                self._selected_profile_labels.discard(str(item))
        except Exception as exc:
            self._set_status("[profile-select] error: {}".format(exc))

    def _filtered_profiles(self):
        selected_cats = {str(item) for item in self.category_list.SelectedItems}
        selected_names = {
            label.split("  (", 1)[0]
            for label in self._selected_profile_labels
        }
        out = []
        for p in self.profiles:
            if not isinstance(p, dict):
                continue
            if selected_cats:
                # Keep the profile if ANY of its LEDs is in a selected
                # category. The LED-level filter (applied by
                # ``execute_placement`` via ``PlacementOptions.category_filter``)
                # then drops the LEDs of unselected categories — so a
                # profile that mixes Electrical Fixtures + Mechanical
                # Equipment only places the LEDs in the categories the
                # user actually selected.
                led_cats = set()
                for s in p.get("linked_sets") or []:
                    if not isinstance(s, dict):
                        continue
                    for led in s.get("linked_element_definitions") or []:
                        if not isinstance(led, dict):
                            continue
                        c = (led.get("category") or "").strip()
                        if c:
                            led_cats.add(c)
                if not (led_cats & selected_cats):
                    continue
            if selected_names:
                if (p.get("name") or "") not in selected_names:
                    continue
            out.append(p)
        return out

    # ---- match button ----------------------------------------------

    def _selected_source_value(self):
        item = self.src_combo.SelectedItem
        return item.value if item is not None else None

    def _on_match_clicked(self, sender, e):
        self._run_match(offer_fuzzy=True)

    def _run_match(self, offer_fuzzy=True):
        source_value = self._selected_source_value()
        if source_value is None and self._source_kind != "csv":
            self._set_status("Pick a source first")
            return

        targets = []
        mode = placement.MATCH_FAMILY_NAME_STRIP_SUFFIX
        if self._source_kind in ("host_model", "linked_revit"):
            phase_item = self.phase_combo.SelectedItem
            if phase_item is None:
                self._set_status("Pick a phase first")
                return
            phase_id = phase_item.value
            if self._source_kind == "host_model":
                targets = placement.find_targets_in_host_model(
                    self.doc, phase_id=phase_id,
                )
            else:
                targets = placement.find_targets_in_linked_revit(
                    source_value, phase_id=phase_id,
                )
            mode = placement.MATCH_FAMILY_NAME_STRIP_SUFFIX
        elif self._source_kind == "csv":
            if not self._csv_path:
                self._set_status("Browse to a CSV first")
                return
            try:
                targets = placement.find_targets_in_csv(self._csv_path)
            except placement.CsvParseError as exc:
                self._set_status(str(exc))
                return
            mode = placement.MATCH_CAD_ALIASES

        profiles = self._filtered_profiles()
        if not profiles:
            self._set_status("No profiles match the current filters")
            self.matches = []
            self._render_matches([])
            return
        if not targets:
            self._set_status("No targets found in the selected source")
            self.matches = []
            self._render_matches([])
            return

        raw_matches = placement.match_targets(targets, profiles, mode)

        # Matching is strictly exact, so near-miss family names (renames,
        # _2 suffixes, typos) produce zero matches silently. Offer to
        # record ≥80%-similar unmatched names as aliases on their
        # closest profile, then re-match once so they land immediately.
        if offer_fuzzy and self._offer_fuzzy_aliases(
                targets, raw_matches, profiles, mode):
            return self._run_match(offer_fuzzy=False)

        deduped_count = 0
        if self.one_profile_per_target_check.IsChecked:
            self.matches = placement.dedupe_matches_per_target(raw_matches)
            deduped_count = len(raw_matches) - len(self.matches)
        else:
            self.matches = raw_matches
        self.matches.sort(key=lambda m: (
            (m.profile.get("name") or "").lower(),
            (m.target.name or "").lower(),
        ))
        self._render_matches(self.matches)
        summary = "{} target(s) -> {} match(es) across {} profile(s)".format(
            len(targets), len(self.matches), len(profiles),
        )
        if deduped_count:
            summary += "  ({} duplicate match(es) suppressed)".format(deduped_count)
        self.summary_label.Text = summary
        self._set_status(
            "Review the list, uncheck rows to skip, then Place." if self.matches
            else "No matches. Try different filters or a different source."
        )

    # ---- fuzzy alias proposals -------------------------------------

    def _offer_fuzzy_aliases(self, targets, raw_matches, profiles, mode):
        """Prompt to alias unmatched target names onto their closest
        profile (>= 80% similar). Returns True when at least one alias
        was added (caller re-runs the match once)."""
        matched_names = set()
        for m in raw_matches:
            name = (getattr(m.target, "name", "") or "").strip().lower()
            if name:
                matched_names.add(name)
        unmatched = []
        seen = set()
        # Every type seen per unmatched family name — shown in the
        # proposal label so the reviewer can judge whether the linked
        # element's TYPE (not just its family) fits the profile.
        types_by_name = {}
        for t in targets:
            name = (getattr(t, "name", "") or "").strip()
            key = name.lower()
            if not name or key in matched_names:
                continue
            type_name = (getattr(t, "type_name", "") or "").strip()
            if type_name:
                types_by_name.setdefault(key, set()).add(type_name)
            if key in seen:
                continue
            seen.add(key)
            unmatched.append(name)
        if not unmatched:
            return False

        profile_keys = []
        for idx, p in enumerate(profiles):
            keys = set(placement.profile_family_names_raw(p))
            if mode == placement.MATCH_CAD_ALIASES:
                keys |= placement.collect_profile_aliases_raw(p)
            if keys:
                profile_keys.append((idx, sorted(keys)))
        if not profile_keys:
            return False

        proposals = fuzzy_match.propose_aliases(unmatched, profile_keys)
        # A name some profile already answers to as an alias must not be
        # re-proposed onto a different profile.
        proposals = [
            prop for prop in proposals
            if merge_workflow.find_alias_owner(self.profile_data, prop[0]) is None
        ]
        if not proposals:
            return False

        def _label(prop):
            name, idx, _key, score = prop
            profile = profiles[idx]
            profile_name = profile.get("name") or "(unnamed)"
            types = sorted(types_by_name.get(name.lower()) or ())
            left = "{} : {}".format(name, " | ".join(types)) if types else name
            pf = profile.get("parent_filter") or {}
            prof_fam = (pf.get("family_name_pattern") or "").strip()
            prof_type = (pf.get("type_name_pattern") or "").strip()
            right = profile_name
            if prof_fam or prof_type:
                right += "  [{} : {}]".format(prof_fam or "*", prof_type or "*")
            return "{}  ->  {}  ({:.0f}%)".format(left, right, score)

        chosen = wpf_dialogs.multi_select_from_list(
            proposals,
            title="Close matches found",
            prompt=(
                "These source families didn't match any profile but are "
                "close to one. Check the ones to add as aliases:\n"
                "Format:  Family : Type  ->  Profile  "
                "[family pattern : type pattern]  (similarity)"
            ),
            display_func=_label,
        )
        if not chosen:
            return False

        added = 0
        for name, idx, _key, _score in chosen:
            if merge_workflow.add_alias(profiles[idx], name):
                added += 1
        if not added:
            return False

        # ``profiles`` holds live references into ``profile_data``, so
        # the re-match sees the aliases immediately; persistence goes
        # through the dedicated ExternalEvent (modeless window — no
        # direct transaction allowed here).
        if self._save_gateway is not None:
            self._save_gateway.request_save(
                self.doc, self.profile_data,
                action="MEPRFP 2.0: add matching aliases",
                on_complete=self._on_alias_save_done,
            )
        else:
            self._set_status(
                "Alias(es) applied for this session only — save gateway "
                "unavailable; use Manage Profiles to persist them."
            )
        return True

    def _on_alias_save_done(self, error):
        if error:
            self._set_status(error)
        else:
            self._set_status("Alias(es) saved to the profile store.")

    # ---- preview rendering -----------------------------------------

    def _clear_match_rows(self):
        self.match_rows_panel.Children.Clear()
        self._match_rows = []
        self.summary_label.Text = ""
        self.place_btn.IsEnabled = False

    def _render_matches(self, matches):
        self._clear_match_rows()
        for match in matches:
            grid, checkbox = self._build_match_row(match)
            self.match_rows_panel.Children.Add(grid)
            self._match_rows.append(_MatchRow(match, grid, checkbox))
        self.place_btn.IsEnabled = bool(matches)

    def _build_match_row(self, match):
        grid = Grid()
        for star in (0.0, 4.0, 4.0, 3.0):
            col = ColumnDefinition()
            if star == 0.0:
                col.Width = GridLength(28)
            else:
                col.Width = GridLength(star, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        checkbox = CheckBox()
        checkbox.IsChecked = True
        checkbox.Margin = Thickness(4, 2, 0, 2)
        checkbox.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(checkbox, 0)
        grid.Children.Add(checkbox)

        target_label = TextBlock()
        target_label.Text = "{}  @ ({:.2f}, {:.2f}, {:.2f})  rot {:.1f}°".format(
            match.target.name,
            match.target.world_pt[0], match.target.world_pt[1], match.target.world_pt[2],
            match.target.rotation_deg,
        )
        target_label.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(target_label, 1)
        grid.Children.Add(target_label)

        arrow = TextBlock()
        arrow.Text = "  ->  "
        arrow.Margin = Thickness(0, 4, 0, 4)
        Grid.SetColumn(arrow, 2)
        grid.Children.Add(arrow)

        profile_label = TextBlock()
        profile_label.Text = "{}  ({})".format(
            match.profile.get("name") or "(unnamed)",
            match.profile.get("id") or "?",
        )
        profile_label.Margin = Thickness(0, 4, 8, 4)
        Grid.SetColumn(profile_label, 3)
        grid.Children.Add(profile_label)

        return grid, checkbox

    def _on_check_all(self, sender, e):
        for row in self._match_rows:
            row.checkbox.IsChecked = True

    def _on_uncheck_all(self, sender, e):
        for row in self._match_rows:
            row.checkbox.IsChecked = False

    # ---- place ------------------------------------------------------

    def _on_place_clicked(self, sender, e):
        chosen = []
        for row in self._match_rows:
            if not row.checked:
                continue
            chosen.append(row.match)
        if not chosen:
            self._set_status("Nothing checked to place")
            return

        if self._gateway is None:
            self._set_status(
                "Placement gateway unavailable: {}".format(
                    getattr(self, "_gateway_error", "unknown error")
                )
            )
            return

        selected_cats = {str(item) for item in self.category_list.SelectedItems}
        default_level_id = self._resolve_default_level_id()

        # Host-only mounting-height override. Only honoured for the
        # host-model source; parse the write-in value and abort the run
        # (rather than silently ignoring it) if the user checked the box
        # but typed something we can't read.
        host_height_inches = None
        if (self._source_kind == "host_model"
                and bool(self.host_height_check.IsChecked)):
            raw = (self.host_height_combo.Text or "").strip()
            host_height_inches = placement.parse_relative_height_inches(raw)
            if host_height_inches is None:
                self._set_status(
                    "Can't read height '{}'. Use e.g. 66, 66\", 5' 6\", "
                    "or 5' - 6\".".format(raw)
                )
                return

        options = placement.PlacementOptions(
            skip_already_placed=bool(self.skip_placed_check.IsChecked),
            allow_type_substitution=bool(self.allow_type_sub_check.IsChecked),
            default_level_id=default_level_id,
            category_filter=selected_cats or None,
            uidoc=self.uidoc,
            host_relative_height_inches=host_height_inches,
        )

        # Hand the run to Revit's main thread via the modeless
        # ExternalEvent gateway. That context (no modal dialog, no open
        # transaction) is the only place ``uidoc.ActiveView = view`` is
        # legal — and switching the active view to each target level's
        # plan before ``NewFamilyInstance`` is what binds the Level on
        # workplane-based families. The dialog stays open; results arrive
        # via ``_on_place_complete``.
        self.place_btn.IsEnabled = False
        self._set_status("Placing on Revit's main thread...")
        self._gateway.request_placement(
            self.doc, self.uidoc, chosen, options,
            on_complete=self._on_place_complete,
        )

    def _on_place_complete(self, result):
        """Runs on Revit's main thread after every level's transaction
        commits. Re-enables Place, updates status, and prints the run
        report to the pyRevit output window."""
        self.committed = True
        self._last_result = result
        try:
            self.place_btn.IsEnabled = True
        except Exception:
            pass
        self._set_status(
            "Placed {} fixture(s); skipped {} already-placed; "
            "{} warning(s).".format(
                result.placed_fixture_count,
                result.skipped_already_placed,
                len(result.warnings),
            )
        )
        self._print_report(result)

    def _print_report(self, result):
        output = self._output
        if output is None:
            return
        try:
            output.print_md(
                "**Placement run complete**\n\n"
                "- Fixtures placed: {}\n"
                "- Element_Linker writes: {}\n"
                "- Static parameter writes: {}\n"
                "- Already-placed (skipped): {}\n"
                "- Normalized-name matches: {}\n"
                "- Type substitutions: {}\n"
                "- Warnings: {}\n".format(
                    result.placed_fixture_count,
                    result.element_linker_writes,
                    getattr(result, "static_param_writes", 0),
                    result.skipped_already_placed,
                    getattr(result, "normalized_match_count", 0),
                    getattr(result, "substituted_type_count", 0),
                    len(result.warnings),
                )
            )
            if result.errors:
                output.print_md(
                    "\n**Errors:**\n"
                    + "\n".join("- {}".format(x) for x in result.errors[:50])
                )
            if result.warnings:
                output.print_md(
                    "\n**Warnings:**\n"
                    + "\n".join("- {}".format(w) for w in result.warnings[:50])
                )
        except Exception:
            pass

    # ---- misc -------------------------------------------------------

    def _resolve_default_level_id(self):
        """Pick a Level ElementId to feed Revit's level-bearing
        ``NewFamilyInstance`` overload.

        Why this exists: the no-level overload
        ``NewFamilyInstance(point, symbol, NonStructural)`` produces
        instances whose ``Level`` parameter is missing entirely from the
        Properties palette for level-based families. Supplying a level
        at creation time gives the user the editable ``Level`` dropdown
        they expect. Face/workplane-hosted families ignore the level
        arg — Revit raises, and ``_place_fixture``'s existing try/except
        falls back to the no-level overload for them, so passing a level
        is safe for every family kind.

        Resolution order:
            1. Active view's ``GenLevel`` (set on plan views).
            2. Lowest-elevation real Level in the project — excluding
               legend / placeholder levels (e.g. ``XX - Legend Level``
               at -100 ft). Without this filter, non-plan active views
               would default every placement to the legend level.
            3. ``None`` if the project has no usable levels at all.
        """
        try:
            active = self.doc.ActiveView
            if active is not None:
                gen = getattr(active, "GenLevel", None)
                if gen is not None:
                    eid = gen.Id
                    val = getattr(eid, "Value", None)
                    if val is None:
                        val = getattr(eid, "IntegerValue", None)
                    if val is not None:
                        return val
        except Exception:
            pass
        try:
            levels = [
                lv for lv in FilteredElementCollector(self.doc).OfClass(Level)
                if placement._is_user_level(lv)
            ]
            levels.sort(key=lambda lv: getattr(lv, "Elevation", 0.0))
            if levels:
                eid = levels[0].Id
                val = getattr(eid, "Value", None)
                if val is None:
                    val = getattr(eid, "IntegerValue", None)
                return val
        except Exception:
            pass
        return None

    def _set_status(self, text):
        self.status_label.Text = text or ""

    def show_modeless(self):
        # Modeless is REQUIRED, not cosmetic: the placement runs through
        # an ExternalEvent, which Revit only services when its main
        # thread is idle. A modal ShowDialog blocks that thread, so the
        # event would never fire (and the active-view switch that binds
        # the Level could never run). Show() returns immediately; the
        # window stays alive because Revit roots modeless windows and the
        # retained event delegates keep this controller referenced.
        self.window.Show()
        return self


def show_modeless(doc, profile_data, uidoc=None, output=None):
    return PlacementController(
        doc, profile_data, uidoc=uidoc, output=output
    ).show_modeless()
