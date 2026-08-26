# -*- coding: utf-8 -*-
"""Modeless WPF UI for the multi-scheme Wire Tools command."""

import re

from System import Action
from System.Windows import DataFormats, DataObject, Visibility
from System.Windows.Input import Key
from pyrevit import DB, UI, forms, script

from UIClasses import resource_loader
from wire_tools_logic import (
    SCHEME_INDIVIDUAL_HOMERUN,
    SCHEME_INTERCONNECT,
    SCHEME_WIRE_BY_CIRCUIT,
    SCHEME_WIRE_TO_NODE,
)


class _Choice(object):
    def __init__(self, value, label):
        self.value = value
        self.label = str(label or "<Unnamed>")

    def __str__(self):
        return self.label


class WireToolsWindow(forms.WPFWindow):
    def __init__(self, xaml_path, gateway, config, resources_root=None,
                 theme_mode="light", accent_mode="blue"):
        self.gateway = gateway
        self.config = config
        self.resources_root = resources_root
        self.theme_mode = resource_loader.normalize_theme_mode(
            theme_mode,
            "light",
        )
        self.accent_mode = resource_loader.normalize_accent_mode(
            accent_mode,
            "blue",
        )
        self.ready = False
        self.suspend_option_events = False
        self.view_supported = False
        self.selection_active = False
        self.invalid_context = False
        self.invalid_context_message = ""
        self.scheme = SCHEME_WIRE_BY_CIRCUIT
        self.device_count = 0
        self.invalid_device_count = 0
        self.selected_device_count = 0
        self.circuit_count = 0
        self.system_type_choices = []
        self.system_type_key = None
        self.system_type_status = ""
        self.node_id = None
        self.node_connector_count = 0
        self.homerun_count = 0
        self.wire_choices = []
        self.tag_choices = []
        self.option_choices_initialized = False
        forms.WPFWindow.__init__(self, xaml_path)
        resource_loader.apply_theme(
            self,
            resources_root=self.resources_root,
            theme_mode=self.theme_mode,
            accent_mode=self.accent_mode,
        )

        self.scheme_combo = self.FindName("SchemeCombo")
        self.scheme_instructions_text = self.FindName("SchemeInstructionsText")
        self.target_view_text = self.FindName("TargetViewText")
        self.select_devices_button = self.FindName("SelectDevicesButton")
        self.use_current_button = self.FindName("UseCurrentButton")
        self.clear_devices_button = self.FindName("ClearDevicesButton")
        self.device_count_text = self.FindName("DeviceCountText")
        self.device_status_text = self.FindName("DeviceStatusText")
        self.system_type_panel = self.FindName("SystemTypePanel")
        self.system_type_combo = self.FindName("SystemTypeCombo")
        self.node_group = self.FindName("NodeGroup")
        self.node_selection_panel = self.FindName("NodeSelectionPanel")
        self.interconnect_scope_panel = self.FindName("InterconnectScopePanel")
        self.select_node_button = self.FindName("SelectNodeButton")
        self.clear_node_button = self.FindName("ClearNodeButton")
        self.node_info_text = self.FindName("NodeInfoText")
        self.wire_type_combo = self.FindName("WireTypeCombo")
        self.branch_chamfer_radio = self.FindName("BranchChamferRadio")
        self.branch_arc_radio = self.FindName("BranchArcRadio")
        self.homerun_arc_radio = self.FindName("HomerunArcRadio")
        self.homerun_chamfer_radio = self.FindName("HomerunChamferRadio")
        self.homerun_length_text = self.FindName("HomerunLengthText")
        self.redraw_existing_toggle = self.FindName("RedrawExistingToggle")
        self.homerun_settings_group = self.FindName("HomerunSettingsGroup")
        self.homerun_panel_direction_radio = self.FindName("HomerunPanelDirectionRadio")
        self.homerun_device_facing_radio = self.FindName("HomerunDeviceFacingRadio")
        self.bend_offset_text = self.FindName("BendOffsetText")
        self.skip_single_device_toggle = self.FindName("SkipSingleDeviceToggle")
        self.interconnect_scope_combo = self.FindName("InterconnectScopeCombo")
        self.select_homeruns_button = self.FindName("SelectHomerunsButton")
        self.clear_homeruns_button = self.FindName("ClearHomerunsButton")
        self.homerun_count_text = self.FindName("HomerunCountText")
        self.tag_type_combo = self.FindName("TagTypeCombo")
        self.add_leaders_toggle = self.FindName("AddLeadersToggle")
        self.existing_tag_behavior_combo = self.FindName("ExistingTagBehaviorCombo")
        self.tag_homeruns_button = self.FindName("TagHomerunsButton")
        self.homerun_tagging_toggle = self.FindName("HomerunTaggingToggle")
        self.homerun_tagging_panel = self.FindName("HomerunTaggingPanel")
        self.bend_info_button = self.FindName("BendInfoButton")
        self.bend_info_popup = self.FindName("BendInfoPopup")
        self.direction_info_button = self.FindName("DirectionInfoButton")
        self.direction_info_popup = self.FindName("DirectionInfoPopup")
        self.selection_instruction_panel = self.FindName("SelectionInstructionsPanel")
        self.selection_instruction_text = self.FindName("SelectionInstructionsText")
        self.run_button = self.FindName("RunButton")
        self.status_text = self.FindName("StatusText")
        self.invalid_context_overlay = self.FindName("InvalidContextOverlay")
        self.invalid_context_message_text = self.FindName("InvalidContextMessageText")
        self._attach_numeric_paste_handlers()

        self._load_config()
        self._set_selection_state(False, "")
        self.ready = True
        self._refresh_control_state()

    def _load_config(self):
        scheme_indexes = {
            SCHEME_WIRE_BY_CIRCUIT: 0,
            SCHEME_INTERCONNECT: 1,
            SCHEME_INDIVIDUAL_HOMERUN: 2,
            SCHEME_WIRE_TO_NODE: 3,
        }
        configured_scheme = str(
            getattr(self.config, "scheme", SCHEME_WIRE_BY_CIRCUIT)
        )
        self.scheme_combo.SelectedIndex = scheme_indexes.get(
            configured_scheme,
            0,
        )
        self.scheme = configured_scheme if configured_scheme in scheme_indexes else SCHEME_WIRE_BY_CIRCUIT
        self.initial_wire_type_name = str(
            getattr(self.config, "wire_type_name", "") or ""
        )
        self.initial_tag_type_name = str(
            getattr(self.config, "tag_type_name", "") or ""
        )
        branch_wiring_type = str(
            getattr(self.config, "branch_wiring_type", "Chamfer")
        )
        homerun_wiring_type = str(
            getattr(self.config, "homerun_wiring_type", "Arc")
        )
        self.branch_arc_radio.IsChecked = branch_wiring_type.lower() == "arc"
        self.branch_chamfer_radio.IsChecked = not bool(self.branch_arc_radio.IsChecked)
        self.homerun_arc_radio.IsChecked = homerun_wiring_type.lower() == "arc"
        self.homerun_chamfer_radio.IsChecked = not bool(self.homerun_arc_radio.IsChecked)
        self.homerun_length_text.Text = str(
            getattr(self.config, "homerun_length", 4.0)
        )
        self.redraw_existing_toggle.IsChecked = bool(
            getattr(self.config, "redraw_existing_wires", True)
        )
        configured_direction = getattr(self.config, "homerun_direction", None)
        if configured_direction not in ("panel", "device"):
            legacy_shape = str(
                getattr(self.config, "homerun_shape", None)
                or "revit_default"
            )
            configured_direction = (
                "device"
                if legacy_shape == "straight_from_last_device"
                else "panel"
            )
        self.homerun_device_facing_radio.IsChecked = configured_direction == "device"
        self.homerun_panel_direction_radio.IsChecked = not bool(
            self.homerun_device_facing_radio.IsChecked
        )
        configured_shape = str(
            getattr(self.config, "homerun_shape", "straight") or "straight"
        )
        configured_offset = getattr(self.config, "bend_offset", None)
        if configured_offset is None:
            configured_offset = 1.0 if configured_shape == "bend" else 0.0
        self.bend_offset_text.Text = str(
            configured_offset
        )
        configured_scope = str(
            getattr(self.config, "interconnect_scope", "selected_circuits")
        )
        self.interconnect_scope_combo.SelectedIndex = (
            1 if configured_scope == "selected_only" else 0
        )
        self.skip_single_device_toggle.IsChecked = bool(
            getattr(self.config, "skip_single_device", False)
        )
        self.add_leaders_toggle.IsChecked = bool(
            getattr(self.config, "add_leaders", True)
        )
        existing_behavior = str(
            getattr(self.config, "existing_tag_behavior", "skip_existing")
        )
        for index in range(self.existing_tag_behavior_combo.Items.Count):
            item = self.existing_tag_behavior_combo.Items[index]
            if str(getattr(item, "Tag", "")) == existing_behavior:
                self.existing_tag_behavior_combo.SelectedIndex = index
                break
        if self.existing_tag_behavior_combo.SelectedIndex < 0:
            self.existing_tag_behavior_combo.SelectedIndex = 0
        self._update_scheme_text()

    def _save_config(self):
        if not self.ready:
            return
        try:
            self.config.scheme = self._scheme()
            self.config.branch_wiring_type = self._branch_wiring_type()
            self.config.homerun_wiring_type = self._homerun_wiring_type()
            self.config.homerun_length = self._homerun_length()
            self.config.redraw_existing_wires = bool(
                self.redraw_existing_toggle.IsChecked
            )
            self.config.homerun_direction = self._homerun_direction()
            self.config.homerun_shape = self._homerun_shape()
            self.config.bend_offset = self._bend_offset()
            self.config.interconnect_scope = self._interconnect_scope()
            self.config.skip_single_device = bool(
                self.skip_single_device_toggle.IsChecked
            )
            self.config.add_leaders = bool(self.add_leaders_toggle.IsChecked)
            self.config.existing_tag_behavior = self._existing_tag_behavior()
            selected_wire = self.wire_type_combo.SelectedItem
            selected_tag = self.tag_type_combo.SelectedItem
            self.config.wire_type_name = (
                selected_wire.label if selected_wire is not None else ""
            )
            self.config.tag_type_name = (
                selected_tag.label if selected_tag is not None else ""
            )
            script.save_config()
        except Exception as error:
            script.get_logger().warning(
                "Could not save Wire Tools settings: {}".format(error)
            )

    def _scheme(self):
        selected_item = self.scheme_combo.SelectedItem
        if selected_item is None:
            return self.scheme
        try:
            return str(selected_item.Tag)
        except Exception:
            return self.scheme

    def _homerun_length(self):
        try:
            value = float(str(self.homerun_length_text.Text))
            if value > 0:
                return value
        except Exception:
            pass
        return 4.0

    def _branch_wiring_type(self):
        return "Arc" if bool(self.branch_arc_radio.IsChecked) else "Chamfer"

    def _homerun_wiring_type(self):
        return "Arc" if bool(self.homerun_arc_radio.IsChecked) else "Chamfer"

    def _selected_tag(self, combo_box, fallback):
        selected_item = combo_box.SelectedItem
        if selected_item is None:
            return fallback
        return str(getattr(selected_item, "Tag", fallback))

    def _homerun_direction(self):
        return "device" if bool(self.homerun_device_facing_radio.IsChecked) else "panel"

    def _homerun_shape(self):
        return "bend" if abs(self._bend_offset()) > 1e-9 else "straight"

    def _bend_offset(self):
        try:
            return float(str(self.bend_offset_text.Text))
        except Exception:
            pass
        return 0.0

    def _interconnect_scope(self):
        return self._selected_tag(
            self.interconnect_scope_combo,
            "selected_circuits",
        )

    def _existing_tag_behavior(self):
        selected_item = self.existing_tag_behavior_combo.SelectedItem
        if selected_item is None:
            return "skip_existing"
        return str(getattr(selected_item, "Tag", "skip_existing"))

    def _selected_choice(self, combo_box):
        selected_item = combo_box.SelectedItem
        if selected_item is None:
            return None
        return selected_item

    def _system_type_key(self):
        selected_item = self.system_type_combo.SelectedItem
        if selected_item is not None:
            return selected_item.value
        return self.system_type_key

    def _settings(self):
        wire_choice = self._selected_choice(self.wire_type_combo)
        tag_choice = self._selected_choice(self.tag_type_combo)
        return {
            "wire_type_id": wire_choice.value if wire_choice is not None else None,
            "branch_wiring_type": self._branch_wiring_type(),
            "homerun_wiring_type": self._homerun_wiring_type(),
            "homerun_length": self._homerun_length(),
            "redraw_existing_wires": bool(
                self.redraw_existing_toggle.IsChecked
            ),
            "homerun_direction": self._homerun_direction(),
            "homerun_shape": self._homerun_shape(),
            "bend_offset": self._bend_offset(),
            "interconnect_scope": self._interconnect_scope(),
            "system_type_key": self._system_type_key(),
            "skip_single_device": bool(self.skip_single_device_toggle.IsChecked),
            "tag_type_id": tag_choice.value if tag_choice is not None else None,
            "add_leaders": bool(self.add_leaders_toggle.IsChecked),
            "existing_tag_behavior": self._existing_tag_behavior(),
        }

    def _payload(self):
        return {
            "scheme": self._scheme(),
            "settings": self._settings(),
        }

    def _set_status(self, message):
        self.status_text.Text = str(message or "")

    def tagging_expander_changed(self, sender, args):
        if self.homerun_tagging_panel is None:
            return
        self.homerun_tagging_panel.Visibility = (
            Visibility.Visible
            if bool(sender.IsChecked)
            else Visibility.Collapsed
        )

    def bend_info_toggled(self, sender, args):
        if self.bend_info_popup is not None:
            self.bend_info_popup.IsOpen = bool(sender.IsChecked)

    def bend_info_popup_closed(self, sender, args):
        if self.bend_info_button is not None and self.bend_info_button.IsChecked:
            self.bend_info_button.IsChecked = False

    def direction_info_toggled(self, sender, args):
        if self.direction_info_popup is not None:
            self.direction_info_popup.IsOpen = bool(sender.IsChecked)

    def direction_info_popup_closed(self, sender, args):
        if self.direction_info_button is not None and self.direction_info_button.IsChecked:
            self.direction_info_button.IsChecked = False

    def _attach_numeric_paste_handlers(self):
        for text_box in (self.homerun_length_text, self.bend_offset_text):
            try:
                text_box.AddHandler(
                    DataObject.PastingEvent,
                    self.numeric_input_pasting,
                )
            except Exception:
                # The normal PreviewTextInput filter remains active if an
                # older WPF host does not expose the attached paste event.
                pass

    def _numeric_candidate(self, sender, inserted_text):
        current_text = str(getattr(sender, "Text", "") or "")
        try:
            start = int(getattr(sender, "SelectionStart", 0) or 0)
        except Exception:
            start = 0
        try:
            selected_length = int(getattr(sender, "SelectionLength", 0) or 0)
        except Exception:
            selected_length = 0
        end = start + selected_length
        return current_text[:start] + str(inserted_text or "") + current_text[end:]

    def _is_numeric_text(self, text):
        # Permit temporary editing states such as '-', '.', and '-.' while
        # still rejecting misplaced signs, duplicate decimal points, spaces,
        # and all other non-numeric characters.
        return re.match(r"^-?(?:\d+(?:\.\d*)?|\.\d*)?$", str(text or "")) is not None

    def numeric_input_got_focus(self, sender, args):
        try:
            sender.SelectAll()
        except Exception:
            pass

    def numeric_input_mouse_down(self, sender, args):
        try:
            sender.Focus()
            sender.SelectAll()
            args.Handled = True
        except Exception:
            pass

    def numeric_preview_text_input(self, sender, args):
        candidate = self._numeric_candidate(sender, getattr(args, "Text", ""))
        if not self._is_numeric_text(candidate):
            args.Handled = True

    def numeric_input_lost_focus(self, sender, args):
        text = str(getattr(sender, "Text", "") or "").strip()
        invalid_partial = text in ("", "-", ".", "-.")
        try:
            numeric_value = float(text)
        except Exception:
            numeric_value = None
        if sender is self.homerun_length_text:
            if invalid_partial or numeric_value is None or numeric_value <= 0:
                sender.Text = "4.0"
        elif invalid_partial or numeric_value is None:
            sender.Text = "1.0"
        self.option_changed(sender, args)

    def numeric_input_pasting(self, sender, args):
        try:
            data_object = args.DataObject
            if data_object.GetDataPresent(DataFormats.UnicodeText):
                pasted_text = data_object.GetData(DataFormats.UnicodeText)
            elif data_object.GetDataPresent(DataFormats.Text):
                pasted_text = data_object.GetData(DataFormats.Text)
            else:
                args.CancelCommand()
                return
            candidate = self._numeric_candidate(sender, pasted_text)
            if not self._is_numeric_text(candidate):
                args.CancelCommand()
        except Exception:
            args.CancelCommand()

    def _update_scheme_text(self):
        instructions = {
            SCHEME_WIRE_BY_CIRCUIT: (
                "Connect selected circuits or devices and all other devices sharing their underlying circuits."
            ),
            SCHEME_INTERCONNECT: (
                "Connect selected devices in one continuous wiring sequence, even across multiple circuits."
            ),
            SCHEME_INDIVIDUAL_HOMERUN: (
                "Create one homerun from each selected device using its primary connector and the selected settings."
            ),
            SCHEME_WIRE_TO_NODE: (
                "Connect each selected device directly to one compatible electrical node."
            ),
        }
        self.scheme_instructions_text.Text = instructions.get(
            self._scheme(),
            "Choose a wiring scheme.",
        )

    def _set_selection_state(self, active, message):
        self.selection_active = bool(active)
        self.selection_instruction_text.Text = str(message or "")
        self.selection_instruction_panel.Visibility = (
            Visibility.Visible if active else Visibility.Collapsed
        )
        self.selection_instruction_text.Visibility = (
            Visibility.Visible if active else Visibility.Collapsed
        )
        self._refresh_control_state()

    def set_invalid_context(self, message):
        if self.invalid_context:
            return
        self.invalid_context = True
        self.invalid_context_message = str(message or "Active document changed.")
        self.selection_active = False
        self.selection_instruction_panel.Visibility = Visibility.Collapsed
        self.selection_instruction_text.Visibility = Visibility.Collapsed
        self.invalid_context_message_text.Text = self.invalid_context_message
        self.invalid_context_overlay.Visibility = Visibility.Visible
        self._set_status(self.invalid_context_message)
        self._refresh_control_state()

    def _enforce_scheme_defaults(self, scheme):
        if scheme != SCHEME_INDIVIDUAL_HOMERUN:
            return
        if not bool(self.skip_single_device_toggle.IsChecked):
            return
        self.suspend_option_events = True
        try:
            self.skip_single_device_toggle.IsChecked = False
        finally:
            self.suspend_option_events = False

    def _refresh_control_state(self):
        active = (
            self.view_supported
            and not self.selection_active
            and not self.invalid_context
        )
        scheme = self._scheme()
        self._enforce_scheme_defaults(scheme)
        node_scheme = scheme == SCHEME_WIRE_TO_NODE
        interconnect_scheme = scheme == SCHEME_INTERCONNECT
        node_options_visible = node_scheme or interconnect_scheme
        homerun_scheme = scheme in (
            SCHEME_WIRE_BY_CIRCUIT,
            SCHEME_INDIVIDUAL_HOMERUN,
            SCHEME_INTERCONNECT,
            SCHEME_WIRE_TO_NODE,
        )
        self.scheme_combo.IsEnabled = active
        self.select_devices_button.IsEnabled = active
        self.use_current_button.IsEnabled = active
        self.clear_devices_button.IsEnabled = active
        self.system_type_panel.Visibility = (
            Visibility.Visible
            if len(self.system_type_choices) > 1
            else Visibility.Collapsed
        )
        self.system_type_panel.IsEnabled = active and bool(self.system_type_choices)
        self.system_type_combo.IsEnabled = active and bool(self.system_type_choices)
        self.node_group.Visibility = (
            Visibility.Visible
            if node_options_visible
            else Visibility.Collapsed
        )
        self.node_group.IsEnabled = active and node_options_visible
        self.node_selection_panel.Visibility = (
            Visibility.Visible if node_scheme else Visibility.Collapsed
        )
        self.interconnect_scope_panel.Visibility = (
            Visibility.Visible if interconnect_scheme else Visibility.Collapsed
        )
        self.select_node_button.IsEnabled = active and node_scheme
        self.clear_node_button.IsEnabled = active and node_scheme
        self.interconnect_scope_combo.IsEnabled = active and interconnect_scheme
        self.wire_type_combo.IsEnabled = active
        self.branch_chamfer_radio.IsEnabled = active
        self.branch_arc_radio.IsEnabled = active
        self.homerun_arc_radio.IsEnabled = active
        self.homerun_chamfer_radio.IsEnabled = active
        self.homerun_settings_group.IsEnabled = active and homerun_scheme
        self.homerun_length_text.IsEnabled = active and homerun_scheme
        self.redraw_existing_toggle.IsEnabled = active
        self.homerun_panel_direction_radio.IsEnabled = active and homerun_scheme
        self.homerun_device_facing_radio.IsEnabled = active and homerun_scheme
        self.bend_offset_text.IsEnabled = active and homerun_scheme
        self.skip_single_device_toggle.IsEnabled = (
            active and scheme == SCHEME_WIRE_BY_CIRCUIT
        )
        self.select_homeruns_button.IsEnabled = active
        self.clear_homeruns_button.IsEnabled = active
        self.tag_type_combo.IsEnabled = active
        self.add_leaders_toggle.IsEnabled = active
        self.existing_tag_behavior_combo.IsEnabled = active
        self.homerun_tagging_toggle.IsEnabled = active
        self.bend_info_button.IsEnabled = active and homerun_scheme
        self.direction_info_button.IsEnabled = active and homerun_scheme
        self.run_button.IsEnabled = (
            active
            and self.device_count > 0
            and self._selected_choice(self.wire_type_combo) is not None
            and (not node_scheme or self.node_id is not None)
        )
        self.tag_homeruns_button.IsEnabled = (
            active
            and self.homerun_count > 0
            and self._selected_choice(self.tag_type_combo) is not None
        )

    def _set_device_counts(
            self,
            valid_count,
            invalid_count,
            circuit_count=0,
            selected_count=None,
            system_type_status=""):
        self.device_count = int(valid_count or 0)
        self.invalid_device_count = int(invalid_count or 0)
        if selected_count is None:
            selected_count = self.device_count + self.invalid_device_count
        self.selected_device_count = int(selected_count or 0)
        self.circuit_count = int(circuit_count or 0)
        self.system_type_status = str(system_type_status or "")
        self.device_count_text.Text = (
            "Selected: {} | Valid: {} | Invalid: {}".format(
                self.selected_device_count,
                self.device_count,
                self.invalid_device_count,
            )
        )
        if self.system_type_status:
            self.device_status_text.Text = self.system_type_status
        elif self.device_count:
            if self._scheme() == SCHEME_WIRE_BY_CIRCUIT:
                self.device_status_text.Text = (
                    "{} valid device(s); {} eligible circuit(s) ready."
                    .format(self.device_count, self.circuit_count)
                )
            else:
                self.device_status_text.Text = "{} valid device(s) ready.".format(
                    self.device_count
                )
        else:
            self.device_status_text.Text = "No valid devices selected."
        self._refresh_control_state()

    def _set_homerun_count(self, count):
        self.homerun_count = int(count or 0)
        self.homerun_count_text.Text = "Selected homeruns: {}".format(
            self.homerun_count
        )
        self._refresh_control_state()

    def _populate_choices(self, combo_box, choices, configured_name):
        combo_box.Items.Clear()
        selected_index = -1
        for index, choice_data in enumerate(list(choices or [])):
            choice = _Choice(choice_data.get("id"), choice_data.get("name"))
            combo_box.Items.Add(choice)
            if str(choice.label) == str(configured_name):
                selected_index = index
        if combo_box.Items.Count:
            combo_box.SelectedIndex = selected_index if selected_index >= 0 else 0

    def _populate_system_types(self, choices, selected_key):
        self.system_type_choices = list(choices or [])
        self.system_type_combo.Items.Clear()
        selected_index = -1
        for index, choice_data in enumerate(self.system_type_choices):
            choice = _Choice(choice_data.get("id"), choice_data.get("name"))
            self.system_type_combo.Items.Add(choice)
            if choice.value == selected_key:
                selected_index = index
        if self.system_type_combo.Items.Count:
            self.system_type_combo.SelectedIndex = (
                selected_index if selected_index >= 0 else 0
            )
            self.system_type_key = self.system_type_combo.SelectedItem.value
        else:
            self.system_type_key = None

    def _apply_sync(self, result):
        self.view_supported = bool(result.get("view_supported", False))
        self.target_view_text.Text = "Target View: {}".format(
            result.get("view_name", "(none)")
        )
        self.suspend_option_events = True
        try:
            if not self.option_choices_initialized:
                self._populate_choices(
                    self.wire_type_combo,
                    result.get("wire_types", []),
                    self.initial_wire_type_name,
                )
                self._populate_choices(
                    self.tag_type_combo,
                    result.get("tag_types", []),
                    self.initial_tag_type_name,
                )
                self.option_choices_initialized = True
            self._populate_system_types(
                result.get("system_type_choices", []),
                result.get("system_type_key"),
            )
        finally:
            self.suspend_option_events = False
        self._set_device_counts(
            result.get("device_count", 0),
            result.get("invalid_device_count", 0),
            result.get("circuit_count", 0),
            result.get("selected_device_count"),
            result.get("system_type_status", ""),
        )
        self.node_id = result.get("node_id")
        self.node_connector_count = int(result.get("node_connector_count", 0) or 0)
        if self.node_id:
            self.node_info_text.Text = "{} / {} (ID {}) | {} usable connector(s)".format(
                result.get("node_name", "<Unnamed>"),
                result.get("node_type", "<Unnamed type>"),
                self.node_id,
                self.node_connector_count,
            )
        else:
            self.node_info_text.Text = "No node selected."
        self._set_homerun_count(result.get("homerun_count", 0))
        self._update_scheme_text()
        self._refresh_control_state()

    def receive_result(self, status, action_name, result, error):
        def apply_result():
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
        if status == "error":
            self._set_selection_state(False, "")
            message = str(error) if error is not None else "Unknown Revit operation error."
            lower_message = message.lower()
            if ("floor plan" in lower_message
                    or "reflected ceiling plan" in lower_message):
                self.view_supported = False
                self._refresh_control_state()
            self._set_status(message)
            return
        if status == "invalid_context":
            self.set_invalid_context(
                result.get("message", "Active document changed. Close and reopen Wire Tools.")
            )
            return
        if status == "lifecycle" and action_name == "active_view_changed":
            self.view_supported = bool(result.get("view_supported", False))
            view_name = result.get("view_name", "(none)")
            self.target_view_text.Text = "Target View: {}".format(view_name)
            self._set_selection_state(False, "")
            self.suspend_option_events = True
            try:
                self._populate_system_types([], None)
            finally:
                self.suspend_option_events = False
            self._set_device_counts(0, 0, selected_count=0)
            self.node_id = None
            self.node_connector_count = 0
            self.node_info_text.Text = "No node selected."
            self._set_homerun_count(0)
            self._refresh_control_state()
            if self.view_supported:
                self._set_status(
                    "Target view changed to {}; saved selections were cleared."
                    .format(view_name)
                )
            else:
                self._set_status(
                    "Wire Tools is disabled in {}. Switch to a floor plan or "
                    "reflected ceiling plan.".format(view_name)
                )
            return
        if status == "unavailable":
            self.view_supported = False
            self._set_selection_state(False, "")
            self._refresh_control_state()
            self._set_status(str(error or "Revit is not available."))
            return
        if status == "lifecycle" and action_name == "document_closed":
            self.view_supported = False
            self._set_selection_state(False, "")
            self._refresh_control_state()
            self._set_status(result.get("message", "The project document was closed."))
            return
        if status == "cancelled":
            self._set_selection_state(False, "")
            self._set_status("Selection cancelled. Saved selections were retained.")
            return
        if action_name == "sync":
            self._apply_sync(result)
            if self.view_supported:
                self._set_status(
                    "Ready in {}. Select devices or use the current selection.".format(
                        result.get("view_name", "the active view")
                    )
                )
            else:
                self._set_status(
                    "Wire Tools is disabled in {}. Switch to a floor plan or reflected "
                    "ceiling plan.".format(result.get("view_name", "this view"))
                )
            return
        if action_name in ("select_devices", "use_current_selection"):
            self._set_selection_state(False, "")
            self._set_device_counts(
                result.get("device_count", 0),
                result.get("invalid_device_count", 0),
                result.get("circuit_count", 0),
                result.get("selected_device_count"),
                result.get("system_type_status", ""),
            )
            self.suspend_option_events = True
            try:
                self._populate_system_types(
                    result.get("system_type_choices", []),
                    result.get("system_type_key"),
                )
            finally:
                self.suspend_option_events = False
            self._refresh_control_state()
            if result.get("node_excluded"):
                self._set_status("Device selection updated; the picked node was excluded.")
            else:
                self._set_status("Device selection updated.")
            return
        if action_name == "clear_devices":
            self._populate_system_types([], None)
            self._set_device_counts(0, 0, selected_count=0)
            self._set_status("Device selection cleared.")
            return
        if action_name == "select_node":
            self._set_selection_state(False, "")
            self.node_id = result.get("node_id")
            self._set_status("Node selection updated.")
            self.gateway.raise_action("sync", self._payload())
            return
        if action_name == "clear_node":
            self.node_id = None
            self.node_info_text.Text = "No node selected."
            self._refresh_control_state()
            self._set_status("Node selection cleared.")
            return
        if action_name == "select_homeruns":
            self._set_selection_state(False, "")
            self._set_homerun_count(result.get("homerun_count", 0))
            self._set_status("Homerun selection updated.")
            return
        if action_name == "clear_homeruns":
            self._set_homerun_count(0)
            self._set_status("Homerun selection cleared.")
            return
        if action_name == "run_scheme":
            self._set_status(
                "Created {} wire(s), deleted {}; {} homerun(s); {} failure(s).".format(
                    result.get("created", 0),
                    result.get("deleted", 0),
                    len(result.get("homeruns", [])),
                    len(result.get("failures", [])),
                )
            )
            self._refresh_control_state()
            self.gateway.raise_action("sync", self._payload())
            return
        if action_name == "tag_homeruns":
            self._set_status(
                "Created {} tag(s); deleted {}; skipped {}; {} failure(s).".format(
                    result.get("created", 0),
                    result.get("deleted", 0),
                    len(result.get("skipped", [])),
                    len(result.get("failures", [])),
                )
            )

    def report_gateway_error(self, message):
        self.receive_result("error", "gateway", None, Exception(message))

    def scheme_changed(self, sender, args):
        if not self.ready or self.invalid_context:
            return
        self.scheme = self._scheme()
        self._enforce_scheme_defaults(self.scheme)
        self._update_scheme_text()
        self._refresh_control_state()
        if self.gateway is not None and self.view_supported and not self.selection_active:
            self.gateway.raise_action("sync", self._payload())

    def option_changed(self, sender, args):
        if not self.ready or self.suspend_option_events or self.invalid_context:
            return
        self._update_scheme_text()
        self._refresh_control_state()

    def system_type_changed(self, sender, args):
        del sender
        del args
        if not self.ready or self.suspend_option_events or self.invalid_context:
            return
        selected_item = self.system_type_combo.SelectedItem
        self.system_type_key = (
            selected_item.value if selected_item is not None else None
        )
        self._refresh_control_state()
        if self.gateway is not None and self.view_supported and not self.selection_active:
            self.gateway.raise_action("sync", self._payload())

    def select_devices_clicked(self, sender, args):
        if self.invalid_context:
            return
        self._set_selection_state(
            True,
            "Select electrical devices. Finish or cancel to return to Wire Tools.",
        )
        if not self.gateway.raise_action("select_devices", self._payload()):
            self._set_selection_state(False, "")

    def use_current_clicked(self, sender, args):
        if self.invalid_context:
            return
        if not self.gateway.raise_action("use_current_selection", self._payload()):
            self._set_status("Could not use the current Revit selection.")

    def clear_devices_clicked(self, sender, args):
        if self.invalid_context:
            return
        self.gateway.raise_action("clear_devices")

    def select_node_clicked(self, sender, args):
        if self.invalid_context:
            return
        self._set_selection_state(
            True,
            "Select one electrical node. Finish or cancel to return to Wire Tools.",
        )
        if not self.gateway.raise_action("select_node", self._payload()):
            self._set_selection_state(False, "")

    def clear_node_clicked(self, sender, args):
        if self.invalid_context:
            return
        self.gateway.raise_action("clear_node")

    def select_homeruns_clicked(self, sender, args):
        if self.invalid_context:
            return
        self._set_status("Finding all open-ended homerun wires in the active view...")
        if not self.gateway.raise_action("select_homeruns", self._payload()):
            self._set_status("Could not find homerun wires in the active view.")

    def clear_homeruns_clicked(self, sender, args):
        if self.invalid_context:
            return
        self.gateway.raise_action("clear_homeruns")

    def run_clicked(self, sender, args):
        if self.invalid_context:
            return
        request = self._payload()
        self.run_button.IsEnabled = False
        self._set_status("Creating wires in Revit...")
        if not self.gateway.raise_action("run_scheme", request):
            self._refresh_control_state()

    def tag_homeruns_clicked(self, sender, args):
        if self.invalid_context:
            return
        request = self._payload()
        self.tag_homeruns_button.IsEnabled = False
        self._set_status("Tagging homeruns in Revit...")
        if not self.gateway.raise_action("tag_homeruns", request):
            self._refresh_control_state()

    def close_clicked(self, sender, args):
        self.Close()

    def window_preview_key_down(self, sender, args):
        if args.Key == Key.Escape:
            args.Handled = True
            self._set_status("Escape is disabled for this tool. Use Close.")

    def window_closing(self, sender, args):
        if self.gateway is not None:
            self.gateway.detach_lifecycle()
        self._save_config()
