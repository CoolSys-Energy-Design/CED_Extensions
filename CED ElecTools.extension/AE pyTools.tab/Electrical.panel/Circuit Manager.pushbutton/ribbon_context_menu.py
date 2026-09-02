# -*- coding: utf-8 -*-
"""Right-click "Tutorial" menu on the Circuit Manager ribbon button.

Revit owns the ribbon, and neither Revit nor pyRevit exposes a way to add items
to a ribbon button's context menu. What we can do is attach handlers to the
AdWindows RibbonControl - which is an ordinary WPF Control - suppress Revit's
own quick-access context menu when the right-click lands on our button, and put
up our own ContextMenu instead.

The button is identified at click time by walking up the visual tree to the
first element whose DataContext is an Autodesk.Windows.RibbonItem, so nothing
has to be resolved or cached while the ribbon is still being built.

Install from the extension startup script:

    import ribbon_context_menu
    ribbon_context_menu.install(button_dir)
"""

import os
import sys

import clr

for _asm in ("AdWindows", "PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Windows import ComponentManager, RibbonItem
from System import Uri
from System.Windows import ContextMenuEventHandler, FrameworkElement, LogicalTreeHelper, ResourceDictionary
from System.Windows.Controls import ContextMenu, Image, MenuItem
from System.Windows.Controls.Primitives import PlacementMode
from System.Windows.Input import MouseButtonEventHandler
from System.Windows.Media import VisualTreeHelper

from pyrevit import HOST_APP, script

TUTORIAL_MODULE_NAME = "ced_electools_circuit_manager_tutorial"
HOOK_ENVVAR = "CED_CIRCUIT_MANAGER_TUTORIAL_RIBBON_HOOK"
MENU_HEADER = "Tutorial"
BUTTON_MATCH = "circuit manager"
MAX_IDLING_RETRIES = 40

_LOGGER = script.get_logger()


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _collapse(value):
    return " ".join(str(value or "").split()).strip().lower()


def _parent_of(node):
    if node is None:
        return None
    try:
        parent = VisualTreeHelper.GetParent(node)
        if parent is not None:
            return parent
    except Exception:
        pass
    try:
        return LogicalTreeHelper.GetParent(node)
    except Exception:
        return None


def _ribbon_item_for(source):
    """First RibbonItem found walking up from an event's OriginalSource."""
    node = source
    hops = 0
    while node is not None and hops < 40:
        hops += 1
        try:
            context = getattr(node, "DataContext", None)
        except Exception:
            context = None
        if context is not None and isinstance(context, RibbonItem):
            return context
        node = _parent_of(node)
    return None


def _is_circuit_manager(item):
    if item is None:
        return False
    identifier = _collapse(getattr(item, "Id", ""))
    if BUTTON_MATCH in identifier:
        return True
    for attribute in ("AutomationName", "Text"):
        if _collapse(getattr(item, attribute, "")) == BUTTON_MATCH:
            return True
    return False


def _load_tutorial_module(button_dir):
    module = sys.modules.get(TUTORIAL_MODULE_NAME)
    if module is not None:
        return module
    import imp

    path = os.path.join(button_dir, "tutorial_guide.py")
    return imp.load_source(TUTORIAL_MODULE_NAME, path)


# ---------------------------------------------------------------------------
# external event plumbing
# ---------------------------------------------------------------------------

class _TutorialLaunchHandler(IExternalEventHandler):
    """Runs show_tutorial() back inside a valid Revit API context."""

    def __init__(self, button_dir):
        self._button_dir = button_dir

    def Execute(self, uiapp):
        try:
            module = _load_tutorial_module(self._button_dir)
            module.show_tutorial()
        except Exception as exc:
            _LOGGER.warning("Circuit Manager tutorial failed to open: %s", exc)

    def GetName(self):
        return "CED Circuit Manager Tutorial Launcher"


# ---------------------------------------------------------------------------
# the hook
# ---------------------------------------------------------------------------

class CircuitManagerRibbonHook(object):
    def __init__(self, button_dir):
        self.button_dir = os.path.abspath(button_dir)
        self._ribbon = None
        self._external_event = None
        self._handler = None
        self._glyph = None
        self._idling_handler = None
        self._idling_attempts = 0
        self._pending_show = False

        self._down_handler = MouseButtonEventHandler(self._preview_right_down)
        self._up_handler = MouseButtonEventHandler(self._preview_right_up)
        self._context_handler = ContextMenuEventHandler(self._context_menu_opening)

    # -- install / uninstall ----------------------------------------------

    def install(self):
        self._create_external_event()
        if self._attach():
            return True
        self._start_idling_retry()
        return False

    def uninstall(self):
        self._stop_idling_retry()
        ribbon = self._ribbon
        if ribbon is None:
            return
        try:
            ribbon.PreviewMouseRightButtonDown -= self._down_handler
        except Exception:
            pass
        try:
            ribbon.PreviewMouseRightButtonUp -= self._up_handler
        except Exception:
            pass
        try:
            ribbon.RemoveHandler(FrameworkElement.ContextMenuOpeningEvent, self._context_handler)
        except Exception:
            pass
        self._ribbon = None

    def _create_external_event(self):
        if self._external_event is not None:
            return
        try:
            self._handler = _TutorialLaunchHandler(self.button_dir)
            self._external_event = ExternalEvent.Create(self._handler)
        except Exception as exc:
            # Not fatal - the menu click falls back to a direct call.
            _LOGGER.debug("Tutorial external event unavailable: %s", exc)
            self._external_event = None

    def _attach(self):
        try:
            ribbon = ComponentManager.Ribbon
        except Exception:
            ribbon = None
        if ribbon is None:
            return False
        try:
            ribbon.PreviewMouseRightButtonDown += self._down_handler
            ribbon.PreviewMouseRightButtonUp += self._up_handler
            ribbon.AddHandler(
                FrameworkElement.ContextMenuOpeningEvent,
                self._context_handler,
                True,
            )
        except Exception as exc:
            _LOGGER.warning("Circuit Manager right-click hook could not attach: %s", exc)
            return False
        self._ribbon = ribbon
        _LOGGER.debug("Circuit Manager right-click hook attached.")
        return True

    # -- deferred attach ---------------------------------------------------

    def _start_idling_retry(self):
        if self._idling_handler is not None:
            return
        uiapp = getattr(HOST_APP, "uiapp", None)
        if uiapp is None:
            return
        self._idling_handler = self._on_idling
        try:
            uiapp.Idling += self._idling_handler
        except Exception:
            self._idling_handler = None

    def _stop_idling_retry(self):
        if self._idling_handler is None:
            return
        uiapp = getattr(HOST_APP, "uiapp", None)
        if uiapp is not None:
            try:
                uiapp.Idling -= self._idling_handler
            except Exception:
                pass
        self._idling_handler = None

    def _on_idling(self, sender, args):
        self._idling_attempts += 1
        if self._attach() or self._idling_attempts >= MAX_IDLING_RETRIES:
            self._stop_idling_retry()

    # -- ribbon events -----------------------------------------------------

    def _hit_is_ours(self, args):
        try:
            source = args.OriginalSource
        except Exception:
            return False
        return _is_circuit_manager(_ribbon_item_for(source))

    def _preview_right_down(self, sender, args):
        if not self._hit_is_ours(args):
            self._pending_show = False
            return
        self._pending_show = True
        try:
            args.Handled = True
        except Exception:
            pass

    def _preview_right_up(self, sender, args):
        if not self._pending_show:
            return
        self._pending_show = False
        if not self._hit_is_ours(args):
            return
        try:
            args.Handled = True
        except Exception:
            pass
        self._show_menu()

    def _context_menu_opening(self, sender, args):
        # Suppress Revit's quick-access-toolbar menu on our button only.
        if self._hit_is_ours(args):
            try:
                args.Handled = True
            except Exception:
                pass

    # -- our menu ----------------------------------------------------------

    def _glyph_image(self):
        if self._glyph is None:
            path = os.path.join(self.button_dir, "ColonelMascot.xaml")
            if os.path.exists(path):
                try:
                    dictionary = ResourceDictionary()
                    dictionary.Source = Uri(path)
                    self._glyph = dictionary["CED.Colonel.Glyph"]
                except Exception as exc:
                    _LOGGER.debug("Colonel glyph unavailable: %s", exc)
                    self._glyph = False
            else:
                self._glyph = False
        if not self._glyph:
            return None
        image = Image()
        image.Source = self._glyph
        image.Width = 16
        image.Height = 16
        return image

    def _show_menu(self):
        try:
            menu = ContextMenu()
            item = MenuItem()
            item.Header = MENU_HEADER
            item.ToolTip = "Open the guided walkthrough of Circuit Manager."
            icon = self._glyph_image()
            if icon is not None:
                item.Icon = icon
            item.Click += self._tutorial_clicked
            menu.Items.Add(item)
            menu.Placement = PlacementMode.MousePoint
            menu.PlacementTarget = self._ribbon
            menu.IsOpen = True
        except Exception as exc:
            _LOGGER.warning("Circuit Manager context menu failed: %s", exc)

    def _tutorial_clicked(self, sender, args):
        if self._external_event is not None:
            try:
                self._external_event.Raise()
                return
            except Exception as exc:
                _LOGGER.debug("Tutorial external event raise failed: %s", exc)
        try:
            _load_tutorial_module(self.button_dir).show_tutorial()
        except Exception as exc:
            _LOGGER.warning("Circuit Manager tutorial failed to open: %s", exc)


def install(button_dir):
    """Attach the hook, replacing any hook left behind by an earlier reload."""
    try:
        previous = script.get_envvar(HOOK_ENVVAR)
    except Exception:
        previous = None
    if previous is not None:
        try:
            previous.uninstall()
        except Exception:
            pass

    hook = CircuitManagerRibbonHook(button_dir)
    hook.install()
    try:
        script.set_envvar(HOOK_ENVVAR, hook)
    except Exception:
        pass
    return hook
