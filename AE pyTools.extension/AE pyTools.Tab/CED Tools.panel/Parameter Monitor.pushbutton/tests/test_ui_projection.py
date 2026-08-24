# -*- coding: utf-8 -*-
from __future__ import print_function

import io
import os
import sys
import types
import unittest

import models

tracking_service_stub = types.ModuleType("tracking_service")
tracking_service_stub.set_summary = lambda tracking_set: {}
previous_tracking_service = sys.modules.get("tracking_service")
sys.modules["tracking_service"] = tracking_service_stub
import viewmodel
if previous_tracking_service is None:
    del sys.modules["tracking_service"]
else:
    sys.modules["tracking_service"] = previous_tracking_service


def _record(name, accepted, current, changed=False):
    key = "name:sample:instance"
    return {
        "metadata": {
            "friendly_name": name,
            "family_type": "Family : Type",
            "element_id": name[-1],
            "level": "Level 1",
        },
        "state": models.ELEMENT_TRACKED,
        "accepted_properties": {
            key: {"state": models.VALUE_VALID, "display": accepted},
        },
        "current_properties": {
            key: {"state": models.VALUE_VALID, "display": current},
        },
        "changed_property_keys": [key] if changed else [],
        "change_count": 1 if changed else 0,
        "missing_count": 0,
        "track_location": False,
    }


class UiProjectionTests(unittest.TestCase):
    def test_modeless_window_pins_unicode_service_for_callback_lifetime(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        with io.open(os.path.join(bundle, "script.py"), "r", encoding="utf-8") as stream:
            source = stream.read()
        class_source = source.split("class ParameterMonitorWindow", 1)[1].split(
            "\ndef main():", 1
        )[0]

        self.assertIn("self._text_service = text_service", class_source)
        global_uses = [
            line for line in class_source.splitlines()
            if "text_service." in line and "self._text_service." not in line
        ]
        self.assertEqual([], global_uses)

    def test_checkbox_column_is_fixed_centered_and_frozen(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        with io.open(path, "r", encoding="utf-8") as stream:
            xaml = stream.read()

        self.assertIn('FrozenColumnCount="1"', xaml)
        self.assertIn('Width="38" MinWidth="38" MaxWidth="38"', xaml)
        self.assertIn('CanUserResize="False" CanUserReorder="False"', xaml)
        self.assertIn('HeaderStyle="{StaticResource PM.CheckboxColumnHeader}"', xaml)
        self.assertIn('CellStyle="{StaticResource PM.CheckboxCell}"', xaml)

    def test_type_column_uses_type_name_not_friendly_element_label(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        with io.open(path, "r", encoding="utf-8") as stream:
            xaml = stream.read()

        self.assertIn(
            'Header="Type" Binding="{Binding type}"',
            xaml,
        )
        self.assertNotIn(
            'Header="Type" Binding="{Binding element}"',
            xaml,
        )

    def test_grid_has_reset_all_filters_and_active_column_style(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        xaml_path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        script_path = os.path.join(bundle, "script.py")
        with io.open(xaml_path, "r", encoding="utf-8") as stream:
            xaml = stream.read()
        with io.open(script_path, "r", encoding="utf-8") as stream:
            source = stream.read()

        self.assertIn('x:Name="ResetFiltersButton"', xaml)
        self.assertIn('Click="reset_filters_clicked"', xaml)
        self.assertIn('<Trigger Property="Tag" Value="Active">', xaml)
        self.assertIn('def reset_filters_clicked(', source)

    def test_grid_refresh_preserves_wpf_sort_descriptions(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        with io.open(os.path.join(bundle, "script.py"), "r", encoding="utf-8") as stream:
            source = stream.read()

        self.assertIn("from System.ComponentModel import SortDescription", source)
        self.assertIn("for item in list(control.Items.SortDescriptions)", source)
        self.assertIn("preserve_sort=True", source)
        self.assertIn("column.SortDirection = direction", source)

    def test_location_is_one_semantic_column_and_console_starts_collapsed(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        with io.open(path, "r", encoding="utf-8") as stream:
            xaml = stream.read()

        self.assertIn(
            'Header="Location" Binding="{Binding location_text}"', xaml
        )
        self.assertNotIn('Header="Location Track"', xaml)
        self.assertNotIn('Header="Location Δ"', xaml)
        self.assertIn('x:Key="PM.LocationCell"', xaml)
        self.assertIn('Binding="{Binding location_state}" Value="Untracked"', xaml)
        self.assertIn('<Setter Property="FontStyle" Value="Italic" />', xaml)
        self.assertIn('x:Name="ConsoleExpander" IsExpanded="False"', xaml)

    def test_location_projection_uses_requested_three_states(self):
        untracked = viewmodel.ElementRow("one", _record("Element 1", "A", "A"))
        self.assertEqual(("Untracked", "Untracked"), (
            untracked.location_text, untracked.location_state
        ))

        unchanged_record = _record("Element 2", "A", "A")
        unchanged_record["track_location"] = True
        unchanged = viewmodel.ElementRow("two", unchanged_record)
        self.assertEqual(("Unchanged", "Unchanged"), (
            unchanged.location_text, unchanged.location_state
        ))

        changed_record = _record("Element 3", "A", "A")
        changed_record["track_location"] = True
        changed_record["changed_property_keys"] = [models.LOCATION_PROPERTY_KEY]
        changed_record["change_count"] = 1
        changed = viewmodel.ElementRow("three", changed_record)
        self.assertEqual(("Changed", "Changed"), (
            changed.location_text, changed.location_state
        ))

    def test_inspector_has_family_type_context_and_status_badge_controls(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        xaml_path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        script_path = os.path.join(bundle, "script.py")
        with io.open(xaml_path, "r", encoding="utf-8") as stream:
            xaml = stream.read()
        with io.open(script_path, "r", encoding="utf-8") as stream:
            source = stream.read()

        self.assertIn('x:Name="ElementTypeText"', xaml)
        self.assertIn('x:Name="ElementStatusBadge"', xaml)
        self.assertIn('x:Name="ElementStatusText"', xaml)
        self.assertIn("self.ElementTitleText.Text = element_row.family", source)
        self.assertIn("self.ElementTypeText.Text = element_row.type", source)
        self.assertIn('self.ElementStatusBadge.Tag = element_row.status', source)

    def test_edit_set_uses_custom_checked_only_picker(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        external_path = os.path.join(bundle, "external_events.py")
        picker_path = os.path.join(bundle, "EditSetWindow.xaml")
        with io.open(external_path, "r", encoding="utf-8") as stream:
            external_source = stream.read()
        with io.open(picker_path, "r", encoding="utf-8") as stream:
            picker_xaml = stream.read()

        edit_source = external_source.split("def _edit_set_interactive", 1)[1].split(
            "\ndef _export_definitions", 1
        )[0]
        self.assertIn("edit_set_window.show_edit_set_dialog", edit_source)
        self.assertNotIn("forms.SelectFromList.show", edit_source)
        self.assertIn('x:Name="SetNameTextBox"', picker_xaml)
        self.assertIn('x:Name="ShowCheckedOnlyCheckBox"', picker_xaml)

    def test_monitor_local_theme_polish_preserves_visuals_while_busy(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        xaml_path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        picker_path = os.path.join(bundle, "EditSetWindow.xaml")
        script_path = os.path.join(bundle, "script.py")
        with io.open(xaml_path, "r", encoding="utf-8") as stream:
            xaml = stream.read()
        with io.open(picker_path, "r", encoding="utf-8") as stream:
            picker_xaml = stream.read()
        with io.open(script_path, "r", encoding="utf-8") as stream:
            source = stream.read()

        self.assertIn('UseLayoutRounding="True"', xaml)
        self.assertIn('RowHeight="28"', xaml)
        self.assertIn('<Setter Property="MinHeight" Value="28" />', xaml)
        self.assertIn('<Setter Property="BorderThickness" Value="0" />', xaml)
        self.assertIn('Value="{DynamicResource CED.Brush.AccentBlue}"', xaml)
        self.assertIn('Style="{StaticResource PM.PrimaryButton}" Content="Add Device"', xaml)
        self.assertIn('self.MainContent.IsEnabled = True', source)
        self.assertIn('self.MainContent.IsHitTestVisible = not bool(busy)', source)
        self.assertNotIn('self.MainContent.IsEnabled = not bool(busy)', source)
        self.assertIn(
            'Background="{DynamicResource CED.Brush.ListItemBackground}"',
            picker_xaml,
        )
        self.assertIn('Style="{DynamicResource CED.Button.Apply}"', picker_xaml)

    def test_grid_uses_global_checkbox_style_and_id_has_no_filter_button(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        xaml_path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        script_path = os.path.join(bundle, "script.py")
        with io.open(xaml_path, "r", encoding="utf-8") as stream:
            xaml = stream.read()
        with io.open(script_path, "r", encoding="utf-8") as stream:
            source = stream.read()

        self.assertIn('Style="{DynamicResource CED.Input.CheckBox}"', xaml)
        self.assertNotIn("StaticResource CED.Input.CheckBox", xaml)
        self.assertIn('select_all.Style = self.FindResource("CED.Input.CheckBox")', source)
        self.assertIn('Header="ID" Binding="{Binding element_id}"', xaml)
        self.assertIn('if field == "element_id":', source)
        self.assertIn('column.Header = label', source)

    def test_linked_child_elements_panel_consolidates_relationship_controls(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        with io.open(path, "r", encoding="utf-8") as stream:
            xaml = stream.read()

        self.assertIn('Text="LINKED CHILD ELEMENTS"', xaml)
        self.assertNotIn('Text="HOST DEVICE / CIRCUIT"', xaml)
        self.assertEqual(1, xaml.count('x:Name="SyncElementLinkerButton"'))
        self.assertIn(
            'x:Name="UnlinkDeviceButton" Style="{StaticResource PM.DangerButton}"',
            xaml,
        )
        self.assertIn('x:Name="RelationshipText"', xaml)

    def test_status_actions_are_fixed_to_main_grid_right_edge(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(bundle, "ParameterMonitorWindow.xaml")
        with io.open(path, "r", encoding="utf-8") as stream:
            xaml = stream.read()

        middle = xaml.split('x:Name="MiddleColumnGrid"', 1)[1].split(
            'x:Name="RightPanelBorder"', 1
        )[0]
        self.assertIn('x:Name="StatusBarText"', middle)
        self.assertIn('Grid.Column="0"', middle)
        self.assertIn('x:Name="StatusActionsPanel" Grid.Column="1"', middle)
        self.assertIn('TextWrapping="Wrap"', middle)

    def test_untracked_projection_keeps_last_known_identity_metadata(self):
        record = _record("Element 1", "A", "A")
        record["metadata"].update({
            "element_id": "101",
            "family_name": "Family A",
            "type_name": "Type B",
            "level": "Level 2",
        })
        tracking_set = {
            "elements": {"host:one": record},
            "untracked_ids": ["host:one"],
        }

        self.assertEqual([], viewmodel.element_rows(tracking_set))
        rows = viewmodel.element_rows(
            tracking_set, filter_key=viewmodel.FILTER_UNTRACKED
        )

        self.assertEqual(1, len(rows))
        self.assertEqual("Untracked", rows[0].status)
        self.assertEqual("101", rows[0].element_id)
        self.assertEqual("Family A", rows[0].family)
        self.assertEqual("Type B", rows[0].type)
        self.assertEqual("Level 2", rows[0].level)

    def test_last_check_uses_system_local_date_time_display(self):
        raw = "2026-08-09T12:00:00Z"
        rendered = viewmodel.system_datetime_text(raw)

        self.assertNotEqual(raw, rendered)
        self.assertNotIn("T", rendered)

    def test_edit_picker_hides_raw_parameter_identity_and_storage_details(self):
        bundle = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(bundle, "edit_set_window.py")
        with io.open(path, "r", encoding="utf-8") as stream:
            source = stream.read()
        row_source = source.split("class ParameterChoiceRow", 1)[1].split(
            "\n\nclass EditSetWindow", 1
        )[0]

        self.assertIn('parameter_service.spec_type_label(', row_source)
        self.assertNotIn('descriptor.get("identity_value")', row_source)
        self.assertNotIn('descriptor.get("parameter_id")', row_source)
        self.assertNotIn('descriptor.get("storage_type")', row_source)

    def test_property_grid_values_follow_selected_element_record(self):
        key = "name:sample:instance"
        tracking_set = {
            "tracked_properties": [
                {"key": key, "name": "Sample", "scope": "instance"},
            ],
        }
        first = viewmodel.ElementRow("first", _record("Element 1", "A1", "C1"))
        second = viewmodel.ElementRow(
            "second",
            _record("Element 2", "A2", "C2", changed=True),
        )

        first_rows = viewmodel.property_rows(tracking_set, first)
        second_rows = viewmodel.property_rows(tracking_set, second)

        first_property = [row for row in first_rows if row.key == key][0]
        second_property = [row for row in second_rows if row.key == key][0]
        self.assertEqual((first_property.accepted, first_property.current), ("A1", "C1"))
        self.assertEqual((second_property.accepted, second_property.current), ("A2", "C2"))
        self.assertFalse(first_property.changed)
        self.assertTrue(second_property.changed)

    def test_missing_parameter_uses_its_cell_indicator_not_a_row_status(self):
        record = _record("Element 1", "Accepted", "Accepted")
        record["missing_count"] = 2

        row = viewmodel.ElementRow("first", record)

        self.assertEqual("Unchanged", row.status)
        self.assertEqual("2", row.missing_text)
        self.assertEqual("Missing", row.missing_state)

    def test_linker_children_hidden_from_grid_and_listed_under_parent(self):
        location = {
            "state": models.VALUE_VALID,
            "x": 0.0, "y": 0.0, "z": 0.0, "rotation": 0.0,
        }
        parent = _record("Parent 1", "A", "A")
        parent["persistent_id"] = "host:p1"
        parent["track_location"] = True
        parent["accepted_location"] = dict(location)
        parent["current_location"] = dict(location, x=5.0)
        child_profile = _record("Child 1", "A", "A")
        child_profile["persistent_id"] = "host:c1"
        child_profile["parent_persistent_id"] = "host:p1"
        child_profile["linker_meta"] = {"role": "child", "led_id": "LED-1"}
        child_profile["metadata"]["family_type"] = "Recep Family : Quad"
        child_manual = _record("Child 2", "A", "A")
        child_manual["persistent_id"] = "host:c2"
        child_manual["parent_persistent_id"] = "host:p1"
        child_removed = _record("Child 3", "A", "A")
        child_removed["persistent_id"] = "host:c3"
        child_removed["parent_persistent_id"] = "host:p1"
        child_removed["state"] = models.ELEMENT_REMOVED
        child_linked = _record("Child 4", "A", "A")
        child_linked["persistent_id"] = "link:L1:c4"
        child_linked["parent_persistent_id"] = "host:p1"
        child_linked["linker_meta"] = {"role": "child"}
        tracking_set = {
            "elements": {
                "host:p1": parent,
                "host:c1": child_profile,
                "host:c2": child_manual,
                "host:c3": child_removed,
                "link:L1:c4": child_linked,
            },
            "location_defaults": {
                "translation_tolerance": 0.001,
                "angular_tolerance": 0.0017453292519943296,
            },
        }

        grid_rows = viewmodel.element_rows(tracking_set)
        self.assertEqual([row.persistent_id for row in grid_rows], ["host:p1"])

        info = viewmodel.linked_children_info(tracking_set, parent)
        self.assertEqual(info["count"], 4)
        by_id = dict([(row.persistent_id, row) for row in info["children"]])
        self.assertEqual(by_id["host:c1"].origin, "Profile")
        self.assertEqual(by_id["host:c1"].family, "Recep Family")
        self.assertEqual(by_id["host:c1"].type, "Quad")
        self.assertEqual(by_id["host:c2"].origin, "Manual")
        self.assertTrue(info["parent_moved"])
        # Removed children and linked-model children are not movable.
        self.assertEqual(
            sorted(info["movable_child_ids"]), ["host:c1", "host:c2"]
        )

    def test_main_grid_summary_excludes_inspector_only_children(self):
        parent = _record("Parent 1", "A", "B", changed=True)
        child = _record("Child 1", "A", "B", changed=True)
        child["parent_persistent_id"] = "host:p1"
        tracking_set = {
            "elements": {"host:p1": parent, "host:c1": child},
        }

        summary = viewmodel.main_grid_summary(tracking_set)

        self.assertEqual(1, summary["changed"])
        self.assertEqual(0, summary["unchanged"])

    def test_linked_children_info_in_sync_parent(self):
        parent = _record("Parent 1", "A", "A")
        parent["persistent_id"] = "host:p1"
        child = _record("Child 1", "A", "A")
        child["persistent_id"] = "host:c1"
        child["parent_persistent_id"] = "host:p1"
        tracking_set = {"elements": {"host:p1": parent, "host:c1": child}}
        info = viewmodel.linked_children_info(tracking_set, parent)
        self.assertEqual(info["count"], 1)
        self.assertFalse(info["parent_moved"])
        self.assertEqual(info["movable_child_ids"], [])
        empty = viewmodel.linked_children_info(tracking_set, child)
        self.assertEqual(empty["count"], 0)

    def test_element_rows_keep_their_own_record_objects(self):
        tracking_set = {
            "elements": {
                "first": _record("Element 1", "A1", "C1"),
                "second": _record("Element 2", "A2", "C2", changed=True),
            },
        }

        rows = viewmodel.element_rows(tracking_set)
        by_id = dict((row.persistent_id, row) for row in rows)

        self.assertIsNot(by_id["first"].record, by_id["second"].record)
        self.assertEqual(by_id["first"].record["current_properties"]["name:sample:instance"]["display"], "C1")
        self.assertEqual(by_id["second"].record["current_properties"]["name:sample:instance"]["display"], "C2")

    def test_projects_two_thousand_element_rows(self):
        tracking_set = {"elements": {}}
        for index in range(2000):
            persistent_id = "host:uid-{}".format(index)
            tracking_set["elements"][persistent_id] = _record(
                "Element {}".format(index),
                "A{}".format(index),
                "C{}".format(index),
                changed=(index % 10 == 0),
            )

        rows = viewmodel.element_rows(tracking_set)

        self.assertEqual(2000, len(rows))
        self.assertEqual(200, len([row for row in rows if row.status == "Changed"]))


if __name__ == "__main__":
    unittest.main()
