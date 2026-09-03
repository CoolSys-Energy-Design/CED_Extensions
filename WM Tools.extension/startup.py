# -*- coding: utf-8 -*-
"""Startup hooks for the WM Tools extension.

Registers the Shift+Click gesture on the WM Tools pulldown so it opens the
"What's In Here" tour. pyRevit runs config.py on Shift+Click for pushbuttons,
but a pulldown is a container (PullDownButtonGroup) and has no config script,
so the gesture has to come from the shared ribbon hook.
"""

import os

from pyrevit import script


def _extension_root():
    return os.path.abspath(os.path.dirname(__file__))


def _pulldown_content():
    return os.path.join(
        _extension_root(),
        "AE pyTools.Tab",
        "WM Tools.panel",
        "WM Tools.pulldown",
        "tutorial_content.py",
    )


def _register_tour_ribbon_gestures():
    logger = script.get_logger()
    content = _pulldown_content()
    if not os.path.exists(content):
        logger.warning("WM Tools tour content not found: %s", content)
        return
    try:
        from UIClasses.tutorial import ribbon_hooks

        ribbon_hooks.register([
            {
                "match": "wm tools",
                "gesture": "shift_click",
                "content": content,
            },
        ])
        logger.info("WM Tools tour ribbon gesture registered.")
    except Exception as exc:
        logger.warning("WM Tools tour ribbon gesture failed: %s", exc)


_register_tour_ribbon_gestures()
