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

    def _clicked(sender, args):
        try:
            from UIClasses.tutorial import show_tour_for

            show_tour_for(resolved, modal=modal)
        except Exception as exc:
            _LOGGER.warning("Help button could not open the tour: %s", exc)

    try:
        button.Click += _clicked
    except Exception as exc:
        _LOGGER.debug("Help button could not be hooked: %s", exc)
        return False

    # Keep a reference so the delegate is not collected while the window lives.
    try:
        owner._ced_help_handler = _clicked
    except Exception:
        pass
    try:
        if not button.ToolTip:
            button.ToolTip = "Show me around"
    except Exception:
        pass
    return True
