# -*- coding: utf-8 -*-
"""Shift+Click on the Circuit Manager button opens the guided tour.

Right-clicking the button and picking "Tutorial" runs the same tour; that path
is wired up by the extension startup script via UIClasses.tutorial.ribbon_hooks.
"""

from UIClasses.tutorial import show_tour

show_tour(__file__)
