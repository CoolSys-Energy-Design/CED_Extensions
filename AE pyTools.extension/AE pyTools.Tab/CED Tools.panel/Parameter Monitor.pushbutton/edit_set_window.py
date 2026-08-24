# -*- coding: utf-8 -*-
"""Themed, exact-identity parameter picker for editing a Tracking Set."""

from __future__ import print_function

import os

from System import Object
from System.Collections.Generic import List
from System.Windows import Visibility

import parameter_service
import text_service
from UIClasses import pathing
from UIClasses.ui_bases import CEDWindowBase

THIS_DIR = os.path.abspath(os.path.dirname(__file__))


def _resources_root():
    lib_root = pathing.resolve_lib_root(THIS_DIR)
    resources_root = pathing.resolve_ui_resources_root(lib_root)
    if not resources_root:
        raise RuntimeError("Could not resolve CEDLib UI resources for Edit Tracking Set.")
    return resources_root


class ParameterChoiceRow(object):
    def __init__(self, descriptor, checked=False):
        self.descriptor = descriptor
        self.checked = bool(checked)
        self.visibility = Visibility.Visible
        self.name = text_service.to_text(
            descriptor.get("name") or "Unnamed Parameter"
        )
        scope = "Type" if descriptor.get("scope") == "type" else "Instance"
        kind = text_service.to_text(
            descriptor.get("identity_kind") or "name"
        ).replace("_", " ").title()
        self.identity_text = "{} | {} Parameter".format(scope, kind)
        self.availability_text = "{} / {} elements".format(
            int(descriptor.get("available_count", 0) or 0),
            int(descriptor.get("element_count", 0) or 0),
        )
        self.value_shape_text = text_service.to_text(
            parameter_service.spec_type_label(descriptor.get("spec_type"))
            or "Other"
        )
        self.search_text = " ".join([
            self.name,
            self.identity_text,
            self.availability_text,
            self.value_shape_text,
        ]).lower()


class EditSetWindow(CEDWindowBase):
    theme_aware = True
    use_config_theme = True

    def __init__(
        self,
        descriptors,
        selected_keys,
        set_name,
        track_new_elements,
    ):
        self._rows = []
        self._result = None
        selected_keys = set([
            text_service.to_text(item or "") for item in list(selected_keys or [])
        ])
        for descriptor in list(descriptors or []):
            self._rows.append(ParameterChoiceRow(
                descriptor,
                checked=text_service.to_text(descriptor.get("key") or "") in selected_keys,
            ))
        xaml_path = os.path.join(THIS_DIR, "EditSetWindow.xaml")
        CEDWindowBase.__init__(
            self,
            xaml_source=xaml_path,
            resources_root=_resources_root(),
            theme_aware=True,
            use_config_theme=True,
            handle_esc=True,
            set_owner=True,
        )
        source = List[Object]()
        for row in self._rows:
            source.Add(row)
        self._source = source
        self.ParameterList.ItemsSource = source
        self.SetNameTextBox.Text = text_service.to_text(set_name or "Tracking Set")
        self.TrackNewElementsCheckBox.IsChecked = bool(track_new_elements)
        self._update_count()
        try:
            self.SetNameTextBox.Focus()
            self.SetNameTextBox.SelectAll()
        except Exception:
            pass

    def _visible_rows(self):
        return [
            row for row in self._rows if row.visibility == Visibility.Visible
        ]

    def _refresh_filter(self):
        query = text_service.to_text(
            getattr(self.ParameterSearchTextBox, "Text", "") or ""
        ).strip().lower()
        checked_only = bool(getattr(self.ShowCheckedOnlyCheckBox, "IsChecked", False))
        for row in self._rows:
            visible = (not query or query in row.search_text) and (
                not checked_only or row.checked
            )
            row.visibility = Visibility.Visible if visible else Visibility.Collapsed
        try:
            self.ParameterList.Items.Refresh()
        except Exception:
            pass

    def _update_count(self):
        count = len([row for row in self._rows if row.checked])
        self.SelectedCountText.Text = "{} parameter{} selected".format(
            count, "" if count == 1 else "s"
        )

    def parameter_search_changed(self, sender, args):
        if hasattr(self, "ParameterList"):
            self._refresh_filter()

    def parameter_filter_changed(self, sender, args):
        if hasattr(self, "ParameterList"):
            self._refresh_filter()

    def parameter_checked_changed(self, sender, args):
        row = getattr(sender, "Tag", None)
        if row is None:
            return
        row.checked = bool(sender.IsChecked)
        self._update_count()
        if bool(getattr(self.ShowCheckedOnlyCheckBox, "IsChecked", False)):
            self._refresh_filter()

    def _set_visible(self, checked):
        for row in self._visible_rows():
            row.checked = bool(checked)
        self.ParameterList.Items.Refresh()
        self._update_count()
        self._refresh_filter()

    def select_visible_clicked(self, sender, args):
        self._set_visible(True)

    def clear_visible_clicked(self, sender, args):
        self._set_visible(False)

    def accept_clicked(self, sender, args):
        name = text_service.to_text(self.SetNameTextBox.Text or "").strip()
        if not name:
            self.ValidationText.Text = "Enter a Tracking Set name."
            self.ValidationText.Visibility = Visibility.Visible
            self.SetNameTextBox.Focus()
            return
        self._result = {
            "name": name,
            "descriptors": [
                row.descriptor for row in self._rows if row.checked
            ],
            "track_new_elements": bool(self.TrackNewElementsCheckBox.IsChecked),
        }
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()

    def get_result(self):
        return self._result


def show_edit_set_dialog(
    descriptors,
    selected_keys,
    set_name,
    track_new_elements,
):
    window = EditSetWindow(
        descriptors,
        selected_keys,
        set_name,
        track_new_elements,
    )
    accepted = window.ShowDialog()
    return window.get_result() if bool(accepted) else None
