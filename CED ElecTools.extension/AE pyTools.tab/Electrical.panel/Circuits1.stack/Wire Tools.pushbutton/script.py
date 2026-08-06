# -*- coding: utf-8 -*-
"""Launch the modeless multi-scheme Wire Tools window."""

import os
import sys

from pyrevit import revit, DB, UI, forms, script

from UIClasses import load_theme_state_from_config
from UIClasses import pathing as ui_pathing


def _load_command_modules(command_directory):
    library_directory = os.path.join(command_directory, "lib")
    if library_directory not in sys.path:
        sys.path.append(library_directory)
    from wire_tools_events import WireToolsExternalEventGateway
    from wire_tools_ui import WireToolsWindow
    return WireToolsExternalEventGateway, WireToolsWindow


def _get_config():
    config = script.get_config("wire_tools_config")
    defaults = {
        "scheme": "wire_by_circuit",
        "wire_type_name": "",
        "branch_wiring_type": "Chamfer",
        "homerun_wiring_type": "Arc",
        "homerun_length": 4.0,
        "redraw_existing_wires": True,
        "homerun_direction": "panel",
        "homerun_shape": "straight",
        "bend_offset": 0.0,
        "interconnect_scope": "selected_circuits",
        "skip_single_device": False,
        "add_leaders": True,
        "existing_tag_behavior": "skip_existing",
        "tag_type_name": "",
    }
    for name, value in defaults.items():
        if not hasattr(config, name):
            setattr(config, name, value)
    return config


def main():
    ui_document = revit.uidoc
    document = revit.doc
    if ui_document is None or document is None:
        forms.alert(
            "Wire Tools requires an active project document.",
            title="Wire Tools",
        )
        return

    command_directory = os.path.abspath(os.path.dirname(__file__))
    library_root = ui_pathing.ensure_lib_root_on_syspath(command_directory)
    resources_root = ui_pathing.resolve_ui_resources_root(library_root)
    theme_mode, accent_mode = load_theme_state_from_config(
        default_theme="light",
        default_accent="blue",
    )
    gateway_class, window_class = _load_command_modules(command_directory)
    config = _get_config()
    try:
        ui_application = __revit__
    except NameError:
        ui_application = None

    window = window_class(
        os.path.join(command_directory, "WireTools.xaml"),
        None,
        config,
        resources_root,
        theme_mode,
        accent_mode,
    )
    gateway = gateway_class(window, document, ui_application)
    window.gateway = gateway
    window.Show()
    gateway.attach_lifecycle()
    if gateway.check_context():
        gateway.raise_action("sync")


if __name__ == "__main__":
    main()
