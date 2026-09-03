# -*- coding: utf-8 -*-
"""Wire a tool's in-window Help button to its guided tour.

A tool opts in with two small changes and no dependency on the tour engine's
internals:

1. In its XAML, a button named ``TutorialHelpButton`` - conventionally a small
   "?" icon button in the window header:

       <Button Name="TutorialHelpButton"
               Style="{DynamicResource CED.Button.IconSmall}"
               Width="20" Height="20" Padding="0"
               ToolTip="Show me around">
         <TextBlock Text="?" FontSize="12" FontWeight="SemiBold"
                    HorizontalAlignment="Center" VerticalAlignment="Center"/>
       </Button>

2. In its script, one line once the window exists:

       from UIClasses.tutorial import wire_help
       wire_help(window, __file__)

``wire_help`` is deliberately forgiving: a missing button, a missing content
file, or an engine that fails to import all leave the tool working exactly as
it did before. A Help button that does nothing is a nuisance; a tool that will
not open because its Help button could not be wired is a bug.
"""

import os

from pyrevit import script

DEFAULT_BUTTON_NAME = "TutorialHelpButton"
DEFAULT_DOCS_BUTTON_NAME = "TutorialDocsButton"
DEFAULT_CONTENT_NAME = "tutorial_content.py"

_LOGGER = script.get_logger()


def _resolve_content(caller_file, content_path, content_name):
    if content_path:
        return os.path.abspath(content_path)
    directory = os.path.dirname(os.path.abspath(caller_file))
    candidate = os.path.join(directory, content_name)
    if os.path.exists(candidate):
        return candidate
    # A tool whose window class lives in lib/ passes its own __file__, so the
    # content sits one level up in the bundle folder.
    parent = os.path.join(os.path.dirname(directory), content_name)
    if os.path.exists(parent):
        return parent
    return candidate


def wire_help(
    owner,
    caller_file,
    button_name=DEFAULT_BUTTON_NAME,
    content_path=None,
    content_name=DEFAULT_CONTENT_NAME,
    hide_if_missing=True,
    modal=False,
):
    """Point a Help button at the tour whose content sits beside ``caller_file``.

    Returns True when the button was found and hooked up.
    """
    if owner is None:
        return False

    try:
        button = owner.FindName(button_name)
    except Exception:
        button = None
    if button is None:
        return False

    resolved = _resolve_content(caller_file, content_path, content_name)
    if not os.path.exists(resolved):
        _LOGGER.debug("Help button found but no tour content at %s", resolved)
        if hide_if_missing:
            try:
                from System.Windows import Visibility

                button.Visibility = Visibility.Collapsed
            except Exception:
                pass
        return False

    def _open_tour():
        try:
            from UIClasses.tutorial import show_tour_for

            show_tour_for(resolved, modal=modal)
        except Exception as exc:
            _LOGGER.warning("Help could not open the tour: %s", exc)

    def _open_docs():
        try:
            from UIClasses.tutorial import docs

            docs.show_docs(caller_file)
        except Exception as exc:
            _LOGGER.warning("Help could not open the documentation: %s", exc)

    def _clicked(sender, args):
        """Offer the walkthrough or the written reference."""
        try:
            menu = _build_menu(owner, sender, _open_tour, _open_docs, caller_file)
            if menu is not None:
                menu.IsOpen = True
                return
        except Exception as exc:
            _LOGGER.debug("Help menu failed, opening the tour directly: %s", exc)
        _open_tour()

    try:
        button.Click += _clicked
    except Exception as exc:
        _LOGGER.debug("Help button could not be hooked: %s", exc)
        return False

    # Keep references so the delegates are not collected while the window lives.
    try:
        owner._ced_help_handlers = (_clicked, _open_tour, _open_docs)
    except Exception:
        pass
    try:
        if not button.ToolTip:
            button.ToolTip = "Show me around, or open the documentation"
    except Exception:
        pass
    return True


def _menu_item(owner, header, action):
    from System.Windows.Controls import MenuItem

    item = MenuItem()
    item.Header = header
    try:
        style = owner.TryFindResource("CED.MenuItem.Base")
        if style is not None:
            item.Style = style
    except Exception:
        pass

    def _clicked(sender, args):
        action()

    item.Click += _clicked
    item._ced_action = _clicked  # keep the delegate alive
    return item


def _build_menu(owner, placement_target, open_tour, open_docs, caller_file):
    from System.Windows.Controls import ContextMenu

    menu = ContextMenu()
    menu.Items.Add(_menu_item(owner, "Show me around", open_tour))
    show_docs_item = True
    try:
        from UIClasses.tutorial import docs

        show_docs_item = docs.has_docs(caller_file)
    except Exception:
        pass
    if show_docs_item:
        menu.Items.Add(_menu_item(owner, "Open documentation", open_docs))
    try:
        menu.PlacementTarget = placement_target
    except Exception:
        pass
    return menu


def wire_docs(
    owner,
    caller_file,
    button_name=DEFAULT_DOCS_BUTTON_NAME,
    hide_if_missing=True,
):
    """Point a Docs button at the tool's written reference.

    Same forgiving contract as ``wire_help``: if there is nothing to show, the
    button hides itself rather than becoming a dead control.
    """
    if owner is None:
        return False
    try:
        button = owner.FindName(button_name)
    except Exception:
        button = None
    if button is None:
        return False

    available = True
    try:
        from UIClasses.tutorial import docs

        available = docs.has_docs(caller_file)
    except Exception as exc:
        _LOGGER.debug("Docs availability check failed: %s", exc)
    if not available:
        if hide_if_missing:
            try:
                from System.Windows import Visibility

                button.Visibility = Visibility.Collapsed
            except Exception:
                pass
        return False

    def _clicked(sender, args):
        try:
            from UIClasses.tutorial import docs

            docs.show_docs(caller_file)
        except Exception as exc:
            _LOGGER.warning("Docs button could not open the documentation: %s", exc)

    try:
        button.Click += _clicked
    except Exception as exc:
        _LOGGER.debug("Docs button could not be hooked: %s", exc)
        return False
    try:
        owner._ced_docs_handler = _clicked
    except Exception:
        pass
    try:
        if not button.ToolTip:
            button.ToolTip = "Open this tool's documentation"
    except Exception:
        pass
    return True


def wire_help_and_docs(owner, caller_file, modal=False):
    """Convenience: wire both buttons in one call. Returns (help, docs)."""
    return (
        wire_help(owner, caller_file, modal=modal),
        wire_docs(owner, caller_file),
    )
