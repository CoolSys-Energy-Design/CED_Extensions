# -*- coding: utf-8 -*-
"""Shared "Show Me Around" tours for AE pyTools.

Adding a tour to a tool takes two files in the tool's bundle folder.

1. ``tutorial_content.py`` - the narration, and nothing else:

       TOUR_TITLE = "Wire Tools - Show Me Around"
       TOUR_KEY   = "wire_tools"          # unique; one open window per key

       # Omit HOST (or set it to None) for a narration-only tour.
       HOST = {
           "kind": "window",              # or "dockable_panel"
           "tag": "_my_window_marker",    # window: match on Tag ...
           "title": "Wire Tools",         #         ... or on Title
           "host_element": "RootGrid",    # optional Panel to draw into
       }
       # HOST = {
       #     "kind": "dockable_panel",
       #     "panel_module": "ced_electools_circuit_manager_panel",
       #     "panel_id": "36c3fd8d-98c4-4cf4-92a4-4ac7f3f8c4f2",
       #     "host_element": "DockFrameHost",
       # }

       TARGET_LABELS = {"ApplyButton": "the Apply button"}

       STEPS = [
           {
               "chapter": "Getting oriented",
               "title": "What this does",
               "target": None,            # or an element Name to spotlight
               "body": "One or two sentences.",
               "bullets": ["detail", "detail"],
           },
       ]

2. ``config.py`` - the Shift+Click entry point, two lines:

       from UIClasses.tutorial import show_tour
       show_tour(__file__)

A tool with a window of its own should also carry a Help button, which is the
discoverable route - Shift+Click stays as the shortcut. Add a button named
``TutorialHelpButton`` to the tool's XAML and one line to its script:

       from UIClasses.tutorial import wire_help
       wire_help(window, __file__)

See ``help_button.py`` for the markup to paste.

A step whose ``target`` cannot be resolved falls back to narration-only, so a
tour still reads correctly with the tool closed. Content files are re-read when
they change on disk, so rewording a step needs no pyRevit reload.
"""

import os

from UIClasses.tutorial import engine
from UIClasses.tutorial.help_button import wire_help
from UIClasses.tutorial.engine import (
    DockablePanelHost,
    NoHost,
    SpotlightOverlay,
    TourWindow,
    WindowHost,
    build_host,
    load_module,
    run_tour,
)

__all__ = [
    "DockablePanelHost",
    "NoHost",
    "SpotlightOverlay",
    "TourWindow",
    "WindowHost",
    "build_host",
    "load_module",
    "run_tour",
    "show_tour",
    "show_tour_for",
    "wire_help",
]

DEFAULT_CONTENT_NAME = "tutorial_content.py"


def _module_name_for(path):
    """Stable, unique sys.modules key for a content file."""
    normalized = os.path.normcase(os.path.abspath(path)).replace("\\", "/")
    safe = "".join(char if char.isalnum() else "_" for char in normalized)
    return "ced_tour_content_" + safe


def show_tour_for(content_path, modal=False):
    """Open the tour defined by an explicit tutorial_content.py path."""
    path = os.path.abspath(content_path)
    if not os.path.exists(path):
        from pyrevit import forms

        forms.alert(
            "No tour content found at:\n{0}".format(path),
            title="Show Me Around",
        )
        return None
    content = load_module(_module_name_for(path), path)
    return run_tour(content, modal=modal)


def show_tour(caller_file, content_name=DEFAULT_CONTENT_NAME, modal=False):
    """Open the tour whose content sits next to ``caller_file``.

    Call from a tool's config.py as ``show_tour(__file__)``.
    """
    directory = os.path.dirname(os.path.abspath(caller_file))
    return show_tour_for(os.path.join(directory, content_name), modal=modal)
