# -*- coding: utf-8 -*-
"""Shift+Click on the Circuit Manager button opens the guided tour.

This is the fallback path for the ribbon right-click menu: if the AdWindows
hook in ribbon_context_menu.py ever stops attaching on a future Revit release,
the tour is still one Shift+Click away.
"""

import os
import sys

from pyrevit import forms, script

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
TUTORIAL_MODULE_NAME = "ced_electools_circuit_manager_tutorial"
TITLE = "Circuit Manager Tutorial"


def _tutorial_module():
    module = sys.modules.get(TUTORIAL_MODULE_NAME)
    if module is not None:
        return module
    import imp

    return imp.load_source(
        TUTORIAL_MODULE_NAME,
        os.path.join(THIS_DIR, "tutorial_guide.py"),
    )


try:
    _tutorial_module().show_tutorial()
except Exception as tutorial_exc:
    script.get_logger().warning("Circuit Manager tutorial failed to open: %s", tutorial_exc)
    forms.alert(
        "The Circuit Manager tutorial could not be opened.\n\n{}".format(tutorial_exc),
        title=TITLE,
    )
