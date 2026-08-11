# -*- coding: utf-8 -*-
"""Standalone Show One Line diagram.

Temporarily split out of Circuit Manager for testing. Reuses the pane
module's window, model builder, and gateways through a small host adapter,
so there is a single source of truth and reconnecting to the pane later is
just deleting this button.
"""

import imp
import os
import sys

from pyrevit import forms, revit, script

logger = script.get_logger()

PANEL_MODULE_NAME = "ced_electools_circuit_manager_panel"
TITLE = "Show One Line"
HOST_ATTR = "_standalone_one_line_host"
WINDOW_ATTR = "_standalone_one_line_window"


def _load_panel_module():
    """The Circuit Manager module (already loaded at startup, or from source)."""
    module = sys.modules.get(PANEL_MODULE_NAME)
    if module is not None:
        return module
    path = os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            "..",
            "Circuit Manager.pushbutton",
            "CircuitBrowserPanel.py",
        )
    )
    return imp.load_source(PANEL_MODULE_NAME, path)


def _theme_settings(module):
    theme = getattr(module, "CURRENT_THEME_MODE", "light") or "light"
    accent = getattr(module, "CURRENT_ACCENT_MODE", "blue") or "blue"
    try:
        from pyrevit.userconfig import user_config
        section = user_config.get_section(module.THEME_CONFIG_SECTION)
        theme = (section.get_option(module.THEME_CONFIG_THEME_KEY, theme) or theme).lower()
        accent = (section.get_option(module.THEME_CONFIG_ACCENT_KEY, accent) or accent).lower()
    except Exception:
        pass
    return theme, accent


class _IdleGateway(object):
    """Stand-in for pane gateways the standalone window only polls."""

    def is_busy(self):
        return False


class _OneLineHost(object):
    """Minimal Circuit Manager pane stand-in for OneLineDiagramWindow.

    IMPORTANT: every method resolves Revit/pyRevit access through
    self._module (the startup-loaded Circuit Manager module). This script's
    own globals are cleaned up after the button command ends, so referencing
    them from later WPF callbacks would NameError silently - that presented
    as 'not a drafting view' even inside one.
    """

    def __init__(self, module):
        self._module = module
        self._logger = module.script.get_logger()
        self._move_gateway = module.MoveCircuitsExternalEventGateway(
            logger=self._logger,
            alert_parameter_name=module.ALERT_DATA_PARAM,
        )
        self._build_branch_gateway = module.BuildBranchExternalEventGateway(logger=self._logger)
        self._operation_gateway = _IdleGateway()

    def _uidoc(self):
        module = self._module
        try:
            uidoc = module.HOST_APP.uidoc
            if uidoc is not None:
                return uidoc
        except Exception:
            pass
        try:
            return module.revit.uidoc
        except Exception:
            return None

    def _get_active_doc(self):
        module = self._module
        try:
            uidoc = self._uidoc()
            if uidoc is not None:
                return uidoc.Document
        except Exception:
            pass
        try:
            return module.revit.doc
        except Exception:
            return None

    @property
    def _active_view_is_drafting(self):
        module = self._module
        view = None
        try:
            uidoc = self._uidoc()
            view = getattr(uidoc, "ActiveView", None) if uidoc is not None else None
        except Exception:
            view = None
        if view is None:
            try:
                doc = self._get_active_doc()
                view = doc.ActiveView if doc is not None else None
            except Exception:
                view = None
        try:
            return view is not None and view.ViewType == module.DB.ViewType.DraftingView
        except Exception:
            return False

    def _set_revit_selection(self, elements):
        module = self._module
        try:
            ids = module.List[module.DB.ElementId]()
            for element in list(elements or []):
                if element is not None:
                    ids.Add(element.Id)
            uidoc = self._uidoc()
            if uidoc is not None:
                uidoc.Selection.SetElementIds(ids)
        except Exception as ex:
            try:
                self._logger.debug("Standalone one line selection failed: %s", ex)
            except Exception:
                pass

    def _safe_load_items(self, *args, **kwargs):
        # No pane circuit list to refresh in standalone mode.
        pass


def main():
    try:
        module = _load_panel_module()
    except Exception as ex:
        logger.exception("Failed to load Circuit Manager module: %s", ex)
        forms.alert(
            "Could not load the Circuit Manager module that hosts the "
            "one line code:\n\n{}".format(ex),
            title=TITLE,
        )
        return

    # Reuse an already-open window instead of stacking duplicates.
    window = getattr(module, WINDOW_ATTR, None)
    if window is not None:
        try:
            if window.IsLoaded:
                window.Activate()
                return
        except Exception:
            pass
        setattr(module, WINDOW_ATTR, None)

    doc = revit.doc
    if doc is None:
        forms.alert("Open a model document first.", title=TITLE)
        return

    host = getattr(module, HOST_ATTR, None)
    if host is None:
        host = _OneLineHost(module)
        setattr(module, HOST_ATTR, host)

    try:
        model = module._build_one_line_model(doc)
    except Exception as ex:
        logger.exception("One line model build failed: %s", ex)
        forms.alert("Failed to build the one line diagram:\n\n{}".format(ex), title=TITLE)
        return

    theme, accent = _theme_settings(module)
    window = module.OneLineDiagramWindow(host, model, theme_mode=theme, accent_mode=accent)
    setattr(module, WINDOW_ATTR, window)
    window.Show()


if __name__ == "__main__":
    main()
