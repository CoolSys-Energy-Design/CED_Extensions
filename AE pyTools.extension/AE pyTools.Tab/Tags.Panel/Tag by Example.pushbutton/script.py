# -*- coding: utf-8 -*-

import os
import sys

from pyrevit import revit, DB, UI, forms, script

from Snippets import revit_helpers


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
    )
    gateway = gateway_class(
        window,
        document,
        active_view,
        initial_tag_ids,
        ui_application,
    )
    window.gateway = gateway
    window.Show()
    gateway.attach_lifecycle()
    gateway.raise_action("sync")


if __name__ == "__main__":
    main()
