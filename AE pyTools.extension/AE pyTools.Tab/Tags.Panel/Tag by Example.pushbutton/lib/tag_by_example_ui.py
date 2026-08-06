# -*- coding: utf-8 -*-
"""Modeless WPF UI for Tag by Example."""

from System import Action
from System.Windows import Visibility
from System.Windows.Input import Key
from pyrevit import DB, UI, forms, script


class TagByExampleWindow(forms.WPFWindow):
    def __init__(self, xaml_path, gateway, config):
        self.gateway = gateway
        self.config = config
        self.ready = False
        self.is_closed = False
        self.view_supported = True
        self.example_available = False
        self.selection_active = False
        self.target_count = 0
        forms.WPFWindow.__init__(self, xaml_path)

        self.example_type_text = self.FindName("ExampleTypeText")
        self.host_category_text = self.FindName("HostCategoryText")
        self.host_family_text = self.FindName("HostFamilyText")
        self.host_type_text = self.FindName("HostTypeText")
        self.owner_view_text = self.FindName("OwnerViewText")
        self.reference_count_text = self.FindName("ReferenceCountText")
        self.leader_text = self.FindName("LeaderText")
        self.selection_instruction_text = self.FindName("SelectionInstructionsText")
        self.target_mode_combo = self.FindName("TargetModeCombo")
        self.new_targets_button = self.FindName("NewTargetsButton")
        self.edit_targets_button = self.FindName("EditTargetsButton")
        self.preview_selection_button = self.FindName("PreviewSelectionButton")
        self.target_count_text = self.FindName("TargetCountText")
        self.status_text = self.FindName("StatusText")
        self.create_button = self.FindName("CreateButton")
        self.pick_example_button = self.FindName("PickExampleButton")
        self.include_nested_toggle = self.FindName("IncludeNestedToggle")
        self.preserve_rotation_toggle = self.FindName("PreserveRotationToggle")
        self.use_model_orientation_toggle = self.FindName("UseModelOrientationToggle")
        self.copy_leader_toggle = self.FindName("CopyLeaderToggle")
        self.replace_all_radio = self.FindName("ReplaceAllRadio")
        self.replace_matching_radio = self.FindName("ReplaceMatchingRadio")
        self.skip_matching_radio = self.FindName("SkipMatchingRadio")

        self._load_config()
        self._set_enabled_for_example(False)
        self.ready = True

    def _load_config(self):
        mode = str(getattr(self.config, "target_mode", "type"))
        mode_indexes = {"type": 0, "family": 1, "category": 2, "manual": 3}
        self.target_mode_combo.SelectedIndex = mode_indexes.get(mode, 0)
        toggle_values = {
            self.include_nested_toggle: "include_nested",
            self.preserve_rotation_toggle: "preserve_rotation",
            self.use_model_orientation_toggle: "use_model_orientation",
            self.copy_leader_toggle: "copy_leader",
        }
        for control, name in toggle_values.items():
            default_value = bool(control.IsChecked)
            control.IsChecked = bool(getattr(self.config, name, default_value))

        existing_behavior = str(
            getattr(self.config, "existing_behavior", "skip_matching")
        )
        legacy_values = {
            "keep": "skip_matching",
            "any": "replace_all",
            "same_type": "replace_matching",
        }
        existing_behavior = legacy_values.get(existing_behavior, existing_behavior)
        self.replace_all_radio.IsChecked = existing_behavior == "replace_all"
        self.replace_matching_radio.IsChecked = existing_behavior == "replace_matching"
        self.skip_matching_radio.IsChecked = existing_behavior != "replace_all" and existing_behavior != "replace_matching"

    def _save_config(self):
        if not self.ready:
            return
        self.config.target_mode = self._target_mode()
        values = self._options()
        for name, value in values.items():
            setattr(self.config, name, value)
        script.save_config()

    def _target_mode(self):
        item = self.target_mode_combo.SelectedItem
        if item is None:
            return "type"
        try:
            return str(item.Tag)
        except Exception:
            return "type"

    def _options(self):
        def checked(control):
            return bool(control.IsChecked)

        existing_behavior = "skip_matching"
        if bool(self.replace_all_radio.IsChecked):
            existing_behavior = "replace_all"
        elif bool(self.replace_matching_radio.IsChecked):
            existing_behavior = "replace_matching"
        return {
            "include_nested": checked(self.include_nested_toggle),
            "preserve_rotation": checked(self.preserve_rotation_toggle),
            "use_model_orientation": checked(self.use_model_orientation_toggle),
            "copy_leader": checked(self.copy_leader_toggle),
            "existing_behavior": existing_behavior,
        }

    def _payload(self):
        return {
            "mode": self._target_mode(),
            "options": self._options(),
            "target_ids": list(self.gateway.manual_target_ids),
        }

    def _queue_target_refresh(self):
        def refresh_targets():
            if self.is_closed:
                return
            if self.example_available and self.view_supported and not self.selection_active:
                self.gateway.raise_action("refresh_targets", self._payload())
        try:
            self.Dispatcher.BeginInvoke(Action(refresh_targets))
        except Exception:
            refresh_targets()

    def _set_status(self, text):
        self.status_text.Text = str(text or "")

    def _set_enabled_for_view(self, enabled):
        self.view_supported = bool(enabled)
        self._refresh_control_state()

    def _set_enabled_for_example(self, enabled):
        self.example_available = bool(enabled)
        self._refresh_control_state()

    def _refresh_control_state(self):
        active = self.view_supported and self.example_available and not self.selection_active
        manual_mode = self._target_mode() == "manual"
        self.pick_example_button.IsEnabled = self.view_supported and not self.selection_active
        self.target_mode_combo.IsEnabled = active
        self.include_nested_toggle.IsEnabled = active
        self.preserve_rotation_toggle.IsEnabled = active
        self.use_model_orientation_toggle.IsEnabled = active
        self.copy_leader_toggle.IsEnabled = active
        self.replace_all_radio.IsEnabled = active
        self.replace_matching_radio.IsEnabled = active
        self.skip_matching_radio.IsEnabled = active
        self.new_targets_button.IsEnabled = active and manual_mode
        self.edit_targets_button.IsEnabled = active and manual_mode and self.target_count > 0
        self.preview_selection_button.IsEnabled = active and manual_mode and self.target_count > 0
        self.create_button.IsEnabled = active and self.target_count > 0

    def _set_selection_state(self, active, message):
        self.selection_active = bool(active)
        self.selection_instruction_text.Text = str(message or "")
        self.selection_instruction_text.Visibility = (
            Visibility.Visible if active else Visibility.Collapsed
        )
        self._refresh_control_state()

    def _set_target_count(self, count):
        self.target_count = int(count or 0)
        self.target_count_text.Text = "Targets: {}".format(self.target_count)
        self._refresh_control_state()

    def receive_result(self, status, action_name, result, error):
        if self.is_closed:
            return

        def apply_result():
            if self.is_closed:
                return
            self._apply_result(status, action_name, result, error)

        try:
            if self.Dispatcher.CheckAccess():
                apply_result()
            else:
                self.Dispatcher.BeginInvoke(Action(apply_result))
        except Exception:
            apply_result()

    def _apply_result(self, status, action_name, result, error):
        result = result or {}
        if status == "unavailable":
            self._set_selection_state(False, "")
            self._set_enabled_for_view(False)
            message = str(error) if error is not None else "Revit is not available."
            self._set_status(message)
            return
        if status == "error":
            self._set_selection_state(False, "")
            message = str(error) if error is not None else "Unknown Revit operation error."
            self._set_status(message)
            if ("document" in message.lower()
                    or "reference" in message.lower()
                    or "owner view" in message.lower()
                    or "example" in message.lower()):
                self._set_enabled_for_example(False)
            return

        if status == "lifecycle":
            if action_name == "view_activated":
                self._set_selection_state(False, "")
                view_supported = bool(result.get("view_supported", False))
                self._set_enabled_for_view(view_supported)
                self._set_target_count(0)
                if not view_supported:
                    self._set_status(
                        "Tag by Example is disabled in {}. Switch to a floor plan, "
                        "reflected ceiling plan, or drafting view.".format(
                            result.get("view_name", "this view")
                        )
                    )
                elif self.example_available:
                    self._set_status(
                        "Active view changed. References retained; targets cleared."
                    )
                    self._queue_target_refresh()
                return
            if action_name == "reference_deleted":
                self._clear_snapshot()
                self._set_enabled_for_example(False)
                self._set_target_count(0)
                self._set_status(result.get(
                    "message", "A reference tag was deleted. Pick a new reference."
                ))
                return
            if action_name == "document_changed":
                self._clear_snapshot()
                self._set_enabled_for_view(False)
                self._set_enabled_for_example(False)
                self._set_target_count(0)
                self._set_status(
                    "The active document changed. Close this window and start Tag by Example again."
                )
                return
            if action_name == "document_closed":
                self._clear_snapshot()
                self._set_enabled_for_view(False)
                self._set_enabled_for_example(False)
                self._set_target_count(0)
                self._set_status(result.get(
                    "message", "The reference document was closed."
                ))
                return
        if status == "cancelled":
            self._set_selection_state(False, "")
            self._set_target_count(result.get("count", self.target_count))
            self._set_status("Selection cancelled. {} valid targets retained.".format(
                self.target_count
            ))
            return

        if action_name == "sync":
            if not result.get("view_supported", True):
                self._set_enabled_for_view(False)
                self._set_target_count(0)
                self._set_status(
                    "Tag by Example is disabled in {}. Switch to a floor plan, "
                    "reflected ceiling plan, or drafting view.".format(
                        result.get("view_name", "this view")
                    )
                )
                return
            self._set_enabled_for_view(True)
            if result.get("snapshot"):
                self._apply_snapshot(result.get("snapshot"))
                self._set_enabled_for_example(True)
                if result.get("view_changed"):
                    self._set_target_count(0)
                    self._set_status(
                        "Active view changed. References retained; targets cleared."
                    )
                else:
                    self._set_status("Reference tags are ready. Choose targets.")
                self._queue_target_refresh()
            elif result.get("example_invalid"):
                self._clear_snapshot()
                self._set_enabled_for_example(False)
                self._set_target_count(0)
                self._set_status(result.get("message", "The reference tags are no longer valid."))
            else:
                self._set_enabled_for_example(False)
                self._set_target_count(0)
                self._set_status("Pick one or more reference tags to begin.")
            return

        if action_name == "pick_example":
            self._set_selection_state(False, "")
            if result.get("snapshot"):
                self._apply_snapshot(result.get("snapshot"))
                self._set_enabled_for_view(True)
                self._set_enabled_for_example(True)
                self._set_target_count(0)
                self._set_status("Reference tags are ready. Choose targets.")
                self._queue_target_refresh()
            return

        if action_name == "pick_targets":
            self._set_selection_state(False, "")
            self._set_target_count(result.get("count", 0))
            self._set_status("{} valid manual targets selected.".format(
                self.target_count
            ))
            return

        if action_name == "preview_selection":
            self._set_status("Revit selection updated to the saved target set.")
            return

        if action_name == "refresh_targets":
            self._set_target_count(result.get("count", 0))
            if result.get("skipped_count", 0):
                self._set_status("{} targets available; {} skipped.".format(
                    result.get("count", 0), result.get("skipped_count", 0)
                ))
            else:
                self._set_status("{} targets available.".format(
                    result.get("count", 0)
                ))
            return

        if action_name == "create":
            self._set_status(
                "Created {} tags; deleted {}, matching skipped {}, failures {}."
                .format(
                    result.get("created", 0),
                    result.get("deleted", 0),
                    result.get("existing_skips", 0),
                    len(result.get("failures", [])),
                )
            )
            self._set_enabled_for_example(True)

    def _apply_snapshot(self, snapshot):
        self.example_type_text.Text = snapshot.get("tag_type", "-")
        self.host_category_text.Text = snapshot.get("host_category", "-")
        self.host_family_text.Text = snapshot.get("host_family", "-")
        self.host_type_text.Text = snapshot.get("host_type", "-")
        self.owner_view_text.Text = snapshot.get("owner_view", "-")
        self.reference_count_text.Text = str(snapshot.get("reference_count", 0))
        self.leader_text.Text = "{} / {}".format(
            snapshot.get("leader_summary", "-"),
            snapshot.get("orientation", "-"),
        )

    def _clear_snapshot(self):
        self.example_type_text.Text = "No reference selected"
        self.host_category_text.Text = "-"
        self.host_family_text.Text = "-"
        self.host_type_text.Text = "-"
        self.owner_view_text.Text = "-"
        self.reference_count_text.Text = "0"
        self.leader_text.Text = "-"

    def report_gateway_error(self, message):
        self.receive_result("error", "gateway", None, Exception(message))

    def option_changed(self, sender, args):
        if not self.ready:
            return
        self._save_config()
        if self.gateway.example_tag_ids and self.view_supported and not self.selection_active:
            self.gateway.raise_action("refresh_targets", self._payload())

    def target_mode_changed(self, sender, args):
        if not self.ready:
            return
        self._save_config()
        self._refresh_control_state()
        if self.gateway.example_tag_ids and self.view_supported and not self.selection_active:
            self.gateway.raise_action("refresh_targets", self._payload())

    def pick_example_clicked(self, sender, args):
        self._set_selection_state(
            True,
            "Select one or more reference tags on the same host in Revit. "
            "Finish or cancel the selection to return here.",
        )
        if not self.gateway.raise_action("pick_example"):
            self._set_selection_state(False, "")

    def new_targets_clicked(self, sender, args):
        self._set_selection_state(
            True,
            "Select a new set of compatible target hosts in Revit. "
            "Finish or cancel the selection to return here.",
        )
        payload = self._payload()
        payload["selection_mode"] = "new"
        if not self.gateway.raise_action("pick_targets", payload):
            self._set_selection_state(False, "")

    def edit_targets_clicked(self, sender, args):
        self._set_selection_state(
            True,
            "Edit the saved target set in Revit. Add or remove compatible hosts, "
            "then finish or cancel to return here.",
        )
        payload = self._payload()
        payload["selection_mode"] = "edit"
        if not self.gateway.raise_action("pick_targets", payload):
            self._set_selection_state(False, "")

    def preview_selection_clicked(self, sender, args):
        self._set_status("Updating the Revit selection...")
        if not self.gateway.raise_action("preview_selection", self._payload()):
            self._set_status("Could not update the Revit selection.")

    def create_clicked(self, sender, args):
        self._save_config()
        self.create_button.IsEnabled = False
        self._set_status("Creating tags in Revit...")
        if not self.gateway.raise_action("create", self._payload()):
            self._refresh_control_state()

    def close_clicked(self, sender, args):
        self.Close()

    def window_preview_key_down(self, sender, args):
        if args.Key == Key.Escape:
            args.Handled = True
            self._set_status("Escape is disabled for this tool. Use Close.")

    def window_closing(self, sender, args):
        self.is_closed = True
        self.gateway.detach_lifecycle()
