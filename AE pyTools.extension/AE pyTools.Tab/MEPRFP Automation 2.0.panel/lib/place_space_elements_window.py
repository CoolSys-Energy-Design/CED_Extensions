# -*- coding: utf-8 -*-
"""
Modal preview UI for the Place Space Elements pushbutton.

Renders every placement plan from the workflow as a flat row, then
hands the lot to ``space_apply.apply_plans`` when the user clicks
*Place all*. Status (placed / failed / skipped) lands back into each
row so the user can see in-line which Family/Type names didn't
resolve.

Modal — NOT modeless. Spaces placement runs inside a single
transaction kicked off from a button click while the dialog still
holds the API context, so the ExternalEvent gateway used by
SuperCircuit isn't necessary here.
"""

import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Object as _NetObject  # noqa: E402
from System.Collections.ObjectModel import ObservableCollection  # noqa: E402
from System.Windows import RoutedEventHandler  # noqa: E402

import wpf as _wpf  # noqa: E402
import space_placement_workflow as _spw  # noqa: E402


_XAML_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_resources", "PlaceSpaceElementsWindow.xaml",
)


# ---------------------------------------------------------------------
# Row binding object
# ---------------------------------------------------------------------

class _PreviewRow(object):
    """One DataGrid row.

    Selected is the per-row opt-in for the place run; ``Selected =
    True`` rows get applied, the rest are skipped. ``@property`` +
    ``@setter`` are required so WPF's TwoWay checkbox binding can
    write back through pythonnet.
    """

    def __init__(self, plan, bucket_index=None):
        self.plan = plan
        self.Status = "Pending"
        # Informational plans (no world_pt) start un-checked because
        # they can't be placed anyway — the user can still see them
        # in the preview with a comment explaining why.
        self._selected = bool(getattr(plan, "is_placeable", True))
        # bucket_id -> bucket_name lookup, shared across rows.
        self._bucket_index = bucket_index or {}

    @property
    def Selected(self):
        return self._selected

    @Selected.setter
    def Selected(self, value):
        self._selected = bool(value)

    @property
    def SpaceLabel(self):
        s = self.plan.space
        if s is None:
            return ""
        bits = []
        if s.number:
            bits.append(s.number)
        if s.name:
            bits.append(s.name)
        return " - ".join(bits) or "(unnamed)"

    @property
    def BucketLabel(self):
        """Resolve the profile's bucket_id to a human-readable name.

        The profile is matched to this Space because its bucket_id is
        in the Space's assigned-bucket list — but a profile only ever
        belongs to one bucket, so we display that single bucket here.
        """
        p = self.plan.profile
        if p is None:
            return ""
        bid = (p.bucket_id or "").strip()
        if not bid:
            return "(no bucket)"
        name = self._bucket_index.get(bid)
        if name:
            return "{}  ({})".format(name, bid)
        return bid

    @property
    def ProfileLabel(self):
        p = self.plan.profile
        if p is None:
            return ""
        return "{}  ({})".format(p.name or "(unnamed)", p.id or "??")

    @property
    def Label(self):
        return self.plan.label or ""

    @property
    def KindLabel(self):
        if self.plan.led is None:
            return ""
        return self.plan.led.placement_rule.kind

    @property
    def XText(self):
        return _fmt_float(self.plan.world_pt[0]) if self.plan.world_pt else ""

    @property
    def YText(self):
        return _fmt_float(self.plan.world_pt[1]) if self.plan.world_pt else ""

    @property
    def ZText(self):
        return _fmt_float(self.plan.world_pt[2]) if self.plan.world_pt else ""

    @property
    def RotText(self):
        return _fmt_float(self.plan.rotation_deg)

    @property
    def Comment(self):
        return getattr(self.plan, "comment", "") or ""


def _fmt_float(v):
    try:
        f = float(v)
    except (ValueError, TypeError):
        return ""
    return "{:.3f}".format(f)


# ---------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------

class PlaceSpaceElementsController(object):

    def __init__(self, doc, profile_data=None, door_choices=None):
        self.doc = doc
        self.profile_data = profile_data or {}

        self.window = _wpf.load_xaml(_XAML_PATH)
        self._rows = ObservableCollection[_NetObject]()
        # ``door_choices`` is pre-populated by the pushbutton script
        # via ``space_door_picker.pre_pick_doors`` BEFORE this modal
        # opens — Selection.PickObject can't run while a modal is up,
        # so the pick happens up front and the choices ride into the
        # workflow here. Spaces without an entry fall back to the
        # first-door default inside the workflow.
        self.run = _spw.SpacePlacementRun(
            doc=doc,
            profile_data=self.profile_data,
            door_choices=door_choices or {},
        )

        # bucket_id -> bucket_name lookup so each preview row can show
        # the bucket label alongside the profile.
        self._bucket_index = {}
        for b in (self.profile_data.get("space_buckets") or ()):
            if isinstance(b, dict) and b.get("id"):
                self._bucket_index[b["id"]] = b.get("name") or ""

        self._lookup_controls()
        self._wire_events()
        self._refresh()

    def _lookup_controls(self):
        f = self.window.FindName
        self.summary_label = f("SummaryLabel")
        self.preview_grid = f("PreviewGrid")
        self.refresh_btn = f("RefreshButton")
        self.place_btn = f("PlaceButton")
        self.close_btn = f("CloseButton")
        self.select_all_btn = f("SelectAllButton")
        self.select_none_btn = f("SelectNoneButton")
        self.status_label = f("StatusLabel")
        self.preview_grid.ItemsSource = self._rows

    def _wire_events(self):
        self._h_refresh = RoutedEventHandler(
            lambda s, e: self._safe(self._refresh, "refresh")
        )
        self._h_place = RoutedEventHandler(
            lambda s, e: self._safe(self._on_place, "place")
        )
        self._h_close = RoutedEventHandler(
            lambda s, e: self.window.Close()
        )
        self._h_select_all = RoutedEventHandler(
            lambda s, e: self._safe(lambda: self._set_all_selected(True), "select-all")
        )
        self._h_select_none = RoutedEventHandler(
            lambda s, e: self._safe(lambda: self._set_all_selected(False), "select-none")
        )
        self.refresh_btn.Click += self._h_refresh
        self.place_btn.Click += self._h_place
        self.close_btn.Click += self._h_close
        self.select_all_btn.Click += self._h_select_all
        self.select_none_btn.Click += self._h_select_none

        # Per-row checkbox wiring. The DataGridCheckBoxColumn's
        # TwoWay binding doesn't display the bound value reliably
        # under pythonnet 3, so the per-row helper sets IsChecked
        # programmatically at row.Loaded and routes Checked/Unchecked
        # write-backs to ``row.Selected`` (the @property+@setter pair
        # on _PreviewRow).
        self._row_handles = _wpf.attach_per_row_handlers(
            self.preview_grid,
            checkboxes_per_name={"SelectCheckBox": "Selected"},
        )

    def _set_all_selected(self, value):
        for row in self._rows:
            row.Selected = bool(value)
        self.preview_grid.Items.Refresh()
        n = sum(1 for r in self._rows if r.Selected)
        self._set_status(
            "{} of {} row(s) selected.".format(n, self._rows.Count)
        )

    def _safe(self, fn, label):
        try:
            fn()
        except Exception as exc:
            self._set_status("[{}] error: {}".format(label, exc))
            raise

    def _set_status(self, text):
        self.status_label.Text = text or ""

    # ----- pipeline ------------------------------------------------

    def _refresh(self):
        self._set_status("Collecting placement plans...")
        plans = self.run.collect()
        self._rows.Clear()
        for plan in plans:
            self._rows.Add(_PreviewRow(plan, bucket_index=self._bucket_index))
        n_plans = len(plans)
        n_warns = len(self.run.warnings)
        self._refresh_summary()
        if n_plans == 0:
            extra = "  (See output panel for details.)" if n_warns else ""
            self._set_status("Nothing to place." + extra)
        else:
            self._set_status(
                "Ready. Tick rows to opt in/out, then 'Place selected'. "
                "{} warning(s).".format(n_warns)
            )

    def _refresh_summary(self):
        n_total = self._rows.Count
        n_sel = sum(1 for r in self._rows if r.Selected)
        n_warns = len(self.run.warnings)
        self.summary_label.Text = (
            "{} planned placement(s); {} selected; {} warning(s)".format(
                n_total, n_sel, n_warns,
            )
        )

    def _selected_plans(self):
        # Commit any in-flight cell edit (the Selected checkbox value
        # might still be pending if the user clicked Place while a
        # checkbox was mid-toggle).
        try:
            self.preview_grid.CommitEdit()
            self.preview_grid.CommitEdit()
        except Exception:
            pass
        # Filter out informational rows — they have no world point and
        # ``apply_plans`` would mark them ``no_anchor`` anyway, but
        # excluding them up front gives a cleaner status message.
        return [
            r.plan for r in self._rows
            if r.Selected and getattr(r.plan, "is_placeable", True)
        ]

    def _on_place(self):
        if not self._rows.Count:
            self._set_status("Nothing to place.")
            return
        selected_plans = self._selected_plans()
        if not selected_plans:
            self._set_status("No rows selected — nothing to place.")
            return
        self._set_status("Placing {} of {} row(s)... (one transaction)".format(
            len(selected_plans), self._rows.Count,
        ))
        # Disable the buttons during the run to avoid a double-click.
        self.place_btn.IsEnabled = False
        self.refresh_btn.IsEnabled = False
        try:
            result = self.run.apply(plans=selected_plans)
        finally:
            self.place_btn.IsEnabled = True
            self.refresh_btn.IsEnabled = True
        # Map plan -> row for status writeback.
        plan_to_row = {id(r.plan): r for r in self._rows}
        for plan, _elem in result.placed:
            row = plan_to_row.get(id(plan))
            if row is not None:
                row.Status = "Placed"
        for plan, status, info in result.failed:
            row = plan_to_row.get(id(plan))
            if row is None:
                continue
            if status == "family_missing":
                row.Status = "Family missing: {}".format(info.get("requested_family"))
            elif status == "type_missing":
                row.Status = "Type missing under {}".format(info.get("requested_family"))
            elif status == "no_label":
                row.Status = "No label"
            elif status == "no_anchor":
                row.Status = "Skipped (no anchor)"
            elif status == "create_failed":
                row.Status = "Create failed"
            elif status == "exception":
                row.Status = "Exception: {}".format(info.get("message", ""))
            else:
                row.Status = status

        self.preview_grid.Items.Refresh()
        self._set_status(
            "Done. Placed {} / Failed {}.".format(result.n_placed, result.n_failed)
        )


# ---------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------

def show_modal(doc, profile_data=None, door_choices=None):
    """Open the placement modal.

    ``door_choices`` is a ``{space_element_id: (origin_xy, inward_xy)}``
    map — typically built by ``space_door_picker.pre_pick_doors`` in
    the calling script before this modal opens. Spaces without an
    entry default to the first door in the workflow.
    """
    controller = PlaceSpaceElementsController(
        doc=doc, profile_data=profile_data, door_choices=door_choices,
    )
    controller.window.ShowDialog()
    return controller
