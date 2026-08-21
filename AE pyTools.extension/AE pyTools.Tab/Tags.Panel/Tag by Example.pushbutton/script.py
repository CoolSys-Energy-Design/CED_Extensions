# -*- coding: utf-8 -*-

import os
import sys

from pyrevit import revit, DB, UI, forms, script

from Snippets import revit_helpers
from UIClasses import load_theme_state_from_config
from UIClasses import pathing as ui_pathing


# Set to True when the detailed pyRevit output report is needed for debugging.
SHOW_OUTPUT_REPORT = False


def _load_command_modules(command_directory):
    library_directory = os.path.join(command_directory, "lib")
    if library_directory not in sys.path:
        sys.path.append(library_directory)
    from tag_by_example_events import TagByExampleExternalEventGateway
    from tag_by_example_ui import TagByExampleWindow
    return TagByExampleExternalEventGateway, TagByExampleWindow


def _initial_tag_ids(document, uidoc):
    tag_ids = []
    selected_ids = list(uidoc.Selection.GetElementIds())
    for selected_id in selected_ids:
        selected_element = document.GetElement(selected_id)
        if isinstance(selected_element, DB.IndependentTag):
            tag_ids.append(revit_helpers.get_elementid_value(selected_element.Id))
    return tag_ids


def _get_config():
    config = script.get_config("tag_by_example_config")
    defaults = {
        "target_mode": "type",
        "include_nested": False,
        "preserve_rotation": True,
        "use_model_orientation": True,
        "copy_leader": True,
        "existing_behavior": "skip_matching",
    }
    for name, value in defaults.items():
        if not hasattr(config, name):
            setattr(config, name, value)
    script.save_config()
    return config


def main():
    uidoc = revit.uidoc
    document = revit.doc
    if uidoc is None or document is None:
        forms.alert("Tag by Example requires an active project document.", title="Tag by Example")
        return
    active_view = uidoc.ActiveView
    if active_view is None:
        forms.alert("Tag by Example requires an active view.", title="Tag by Example")
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
    initial_tag_ids = _initial_tag_ids(document, uidoc)
    try:
        ui_application = __revit__
    except NameError:
        ui_application = None
    window = window_class(
        os.path.join(command_directory, "TagByExample.xaml"),
        None,
        config,
        resources_root,
        theme_mode,
        accent_mode,
    )
    gateway = gateway_class(
        window,
        document,
        active_view,
        initial_tag_ids,
        ui_application,
        SHOW_OUTPUT_REPORT,
    )
    window.gateway = gateway
    window.Show()
    gateway.attach_lifecycle()
    gateway.raise_action("sync")


if __name__ == "__main__":
    main()
