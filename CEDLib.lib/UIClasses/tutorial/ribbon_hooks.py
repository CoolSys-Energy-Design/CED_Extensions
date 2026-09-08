# -*- coding: utf-8 -*-
"""Ribbon gestures that launch a "Show Me Around" tour.

Two gestures that pyRevit does not provide on its own:

* ``right_click`` - a "Tutorial" item on a ribbon button's context menu. Still
  supported, but nothing registers it now that every tool with a window carries
  its own Help button. Revit owns the ribbon and exposes no API for this, so we
  attach handlers to the AdWindows RibbonControl (an ordinary WPF Control),
  suppress Revit's own quick-access menu on a registered item, and show ours.

* ``shift_click`` - for PULLDOWNS only. A tool with a window of its own reaches
  its tour from the Help button inside that window, so no gesture is registered
  for it. A pulldown is a container (``PullDownButtonGroup``) with nowhere to
  put a button, and pyRevit only runs config.py for pushbuttons - so without
  this its overview would be unreachable.

Registered items are matched at click time by walking up the visual tree to the
first ``Autodesk.Windows.RibbonItem``, so nothing needs resolving while the
ribbon is still being built.

Extensions register from their startup script:

    from UIClasses.tutorial import ribbon_hooks
    ribbon_hooks.register([
        {"match": "circuit manager", "gesture": "right_click",
         "content": r"...\\Circuit Manager.pushbutton\\tutorial_content.py"},
        {"match": "circuit tools", "gesture": "shift_click",
         "content": r"...\\Circuit Tools.pulldown\\tutorial_content.py"},
    ])
"""

import os

import clr

for _asm in ("AdWindows", "PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Windows import ComponentManager, RibbonItem
from System import Uri
from System.Windows import (
    ContextMenuEventHandler,
    FrameworkElement,
    LogicalTreeHelper,
    ResourceDictionary,
)
from System.Windows.Controls import ContextMenu, Image, MenuItem
from System.Windows.Controls.Primitives import PlacementMode
from System.Windows.Input import Keyboard, ModifierKeys, MouseButtonEventHandler
from System.Windows.Media import VisualTreeHelper

from pyrevit import HOST_APP, script

HOOK_ENVVAR = "CED_TOUR_RIBBON_HOOK"
MENU_HEADER = "Tutorial"
MAX_IDLING_RETRIES = 40

GESTURE_RIGHT_CLICK = "right_click"
GESTURE_SHIFT_CLICK = "shift_click"
GESTURE_BOTH = "both"

_LOGGER = script.get_logger()

MASCOT_PATH = os.path.join(os.path.abspath(os.path.dirname(__file__)), "ColonelMascot.xaml")


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


def _item_identity(item):
    """(id, automation name, text) collapsed to lowercase for matching."""
    if item is None:
        return ("", "", "")
    return (
        _collapse(getattr(item, "Id", "")),
        _collapse(getattr(item, "AutomationName", "")),
        _collapse(getattr(item, "Text", "")),
    )


def _shift_held():
    try:
        return bool(int(Keyboard.Modifiers) & int(ModifierKeys.Shift))
    except Exception:
        try:
            return Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)
        except Exception:
            return False


# ---------------------------------------------------------------------------
# external event plumbing
# ---------------------------------------------------------------------------

class _TourLaunchHandler(IExternalEventHandler):
    """Runs the tour back inside a valid Revit API context."""

    def __init__(self):
        self.content_path = None

    def Execute(self, uiapp):
        path = self.content_path
        self.content_path = None
        if not path:
            return
        try:
            from UIClasses.tutorial import show_tour_for

            show_tour_for(path)
        except Exception as exc:
            _LOGGER.warning("Tour failed to open (%s): %s", path, exc)

    def GetName(self):
        return "CED Show Me Around Tour Launcher"


# ---------------------------------------------------------------------------
# the hook
# ---------------------------------------------------------------------------

class RibbonTourHook(object):
    def __init__(self):
        self.entries = []
        self._ribbon = None
        self._handler = None
        self._external_event = None
        self._glyph = None
        self._idling_handler = None
        self._idling_attempts = 0
        self._pending_right_click = None

        self._down_handler = MouseButtonEventHandler(self._preview_right_down)
        self._up_handler = MouseButtonEventHandler(self._preview_right_up)
        self._left_handler = MouseButtonEventHandler(self._preview_left_down)
        self._context_handler = ContextMenuEventHandler(self._context_menu_opening)

    # -- registry ---------------------------------------------------------

    def add_entries(self, entries):
        for raw in list(entries or []):
            entry = self._normalize(raw)
            if entry is None:
                continue
            self.entries = [e for e in self.entries if e["match"] != entry["match"]]
            self.entries.append(entry)

    @staticmethod
    def _normalize(raw):
        try:
            match = _collapse(raw.get("match"))
            content = raw.get("content")
        except Exception:
            return None
        if not match or not content:
            return None
        gesture = str(raw.get("gesture") or GESTURE_RIGHT_CLICK).strip().lower()
        if gesture not in (GESTURE_RIGHT_CLICK, GESTURE_SHIFT_CLICK, GESTURE_BOTH):
            gesture = GESTURE_RIGHT_CLICK
        return {
            "match": match,
            "gesture": gesture,
            "content": os.path.abspath(str(content)),
            "header": str(raw.get("header") or MENU_HEADER),
        }

    def _entry_for(self, item, gesture):
        identifier, automation, text = _item_identity(item)
        if not (identifier or automation or text):
            return None
        for entry in self.entries:
            if entry["gesture"] not in (gesture, GESTURE_BOTH):
                continue
            match = entry["match"]
            if automation == match or text == match or match in identifier:
                return entry
        return None

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
            ribbon.PreviewMouseLeftButtonDown -= self._left_handler
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
            self._handler = _TourLaunchHandler()
            self._external_event = ExternalEvent.Create(self._handler)
        except Exception as exc:
            # Not fatal - the click falls back to a direct call.
            _LOGGER.debug("Tour external event unavailable: %s", exc)
            self._handler = None
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
            ribbon.PreviewMouseLeftButtonDown += self._left_handler
            ribbon.AddHandler(
                FrameworkElement.ContextMenuOpeningEvent, self._context_handler, True
            )
        except Exception as exc:
            _LOGGER.warning("Tour ribbon hook could not attach: %s", exc)
            return False
        self._ribbon = ribbon
        _LOGGER.debug("Tour ribbon hook attached.")
        return True

    # -- deferred attach --------------------------------------------------

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

    # -- ribbon events ----------------------------------------------------

    def _entry_from_args(self, args, gesture):
        try:
            source = args.OriginalSource
        except Exception:
            return None
        return self._entry_for(_ribbon_item_for(source), gesture)

    def _preview_right_down(self, sender, args):
        entry = self._entry_from_args(args, GESTURE_RIGHT_CLICK)
        self._pending_right_click = entry
        if entry is None:
            return
        try:
            args.Handled = True
        except Exception:
            pass

    def _preview_right_up(self, sender, args):
        entry = self._pending_right_click
        self._pending_right_click = None
        if entry is None:
            return
        if self._entry_from_args(args, GESTURE_RIGHT_CLICK) is not entry:
            return
        try:
            args.Handled = True
        except Exception:
            pass
        self._show_menu(entry)

    def _preview_left_down(self, sender, args):
        if not _shift_held():
            return
        entry = self._entry_from_args(args, GESTURE_SHIFT_CLICK)
        if entry is None:
            return
        try:
            args.Handled = True
        except Exception:
            pass
        self._launch(entry)

    def _context_menu_opening(self, sender, args):
        # Suppress Revit's quick-access-toolbar menu on registered items only.
        if self._entry_from_args(args, GESTURE_RIGHT_CLICK) is not None:
            try:
                args.Handled = True
            except Exception:
                pass

    # -- our menu ---------------------------------------------------------

    def _glyph_image(self):
        if self._glyph is None:
            if os.path.exists(MASCOT_PATH):
                try:
                    dictionary = ResourceDictionary()
                    dictionary.Source = Uri(MASCOT_PATH)
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

    def _show_menu(self, entry):
        try:
            menu = ContextMenu()
            item = MenuItem()
            item.Header = entry["header"]
            item.ToolTip = "Open the guided walkthrough."
            icon = self._glyph_image()
            if icon is not None:
                item.Icon = icon
            item.Tag = entry["content"]
            item.Click += self._menu_item_clicked
            menu.Items.Add(item)
            menu.Placement = PlacementMode.MousePoint
            menu.PlacementTarget = self._ribbon
            menu.IsOpen = True
        except Exception as exc:
            _LOGGER.warning("Tour context menu failed: %s", exc)

    def _menu_item_clicked(self, sender, args):
        try:
            path = str(getattr(sender, "Tag", "") or "")
        except Exception:
            path = ""
        if path:
            self._launch({"content": path})

    def _launch(self, entry):
        path = entry.get("content")
        if not path:
            return
        if self._external_event is not None and self._handler is not None:
            try:
                self._handler.content_path = path
                self._external_event.Raise()
                return
            except Exception as exc:
                _LOGGER.debug("Tour external event raise failed: %s", exc)
        try:
            from UIClasses.tutorial import show_tour_for

            show_tour_for(path)
        except Exception as exc:
            _LOGGER.warning("Tour failed to open (%s): %s", path, exc)


# ---------------------------------------------------------------------------
# public API
# ---------------------------------------------------------------------------

def get_hook():
    """The one hook for this Revit session, surviving pyRevit reloads."""
    try:
        hook = script.get_envvar(HOOK_ENVVAR)
    except Exception:
        hook = None
    if hook is not None and hasattr(hook, "add_entries"):
        return hook
    hook = RibbonTourHook()
    try:
        script.set_envvar(HOOK_ENVVAR, hook)
    except Exception:
        pass
    return hook


def register(entries):
    """Register ribbon gestures and make sure the hook is attached."""
    hook = get_hook()
    hook.add_entries(entries)
    if hook._ribbon is None:
        hook.install()
    return hook


def reset():
    """Detach the current hook - used when a reload re-registers everything."""
    try:
        hook = script.get_envvar(HOOK_ENVVAR)
    except Exception:
        hook = None
    if hook is not None:
        try:
            hook.uninstall()
        except Exception:
            pass
    try:
        script.set_envvar(HOOK_ENVVAR, None)
    except Exception:
        pass
