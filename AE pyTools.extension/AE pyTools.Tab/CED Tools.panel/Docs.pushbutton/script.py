# -*- coding: utf-8 -*-
"""Open the modeless offline CED documentation browser."""

from __future__ import print_function

import clr

for _assembly in ("WindowsBase", "PresentationCore", "PresentationFramework", "System.Data"):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass

from System.Windows import Application, WindowState
TITLE = "CED Documentation"
WINDOW_MARKER = "_ced_documentation_viewer_modeless_v1"


def _find_existing_window():
    try:
        application = Application.Current
        if application is None:
            return None
        for window in application.Windows:
            try:
                if str(getattr(window, "Tag", "") or "") == WINDOW_MARKER:
                    return window
            except Exception:
                continue
    except Exception:
        pass
    return None


def _activate(window):
    try:
        if window.WindowState == WindowState.Minimized:
            window.WindowState = WindowState.Normal
        window.Show()
        window.Activate()
        window.Topmost = True
        window.Topmost = False
        return True
    except Exception:
        return False


def main():
    existing = _find_existing_window()
    if existing is not None and _activate(existing):
        return
    try:
        from pyrevit import script
        from Documentation.viewer import DocumentationViewerWindow

        logger = script.get_logger()
        window = DocumentationViewerWindow(
            documentation_root=None,
            uiapp=globals().get("__revit__"),
            logger=logger,
        )
        window.Show()
    except Exception as error:
        from pyrevit import forms, script

        script.get_logger().error("Documentation viewer failed to open: %s", error)
        forms.alert(
            "The documentation viewer could not open.\n\n{}".format(error),
            title=TITLE,
            warn_icon=True,
        )


if __name__ == "__main__":
    main()
