# -*- coding: utf-8 -*-
"""Guided "show me around" tour for the Circuit Manager dockable pane.

Two pieces work together:

* ``SpotlightOverlay`` injects a hit-test-transparent overlay as a second child
  of the pane's root ``DockFrameHost`` grid and paints a dimmed backdrop with a
  cut-out around the control the current step is talking about. Nothing in
  CircuitBrowserPanel.py or its XAML has to change, and because the overlay
  never takes input the user can keep driving the pane during the tour.
* ``TutorialGuideWindow`` is a small owned, frameless window that carries the
  mascot and the narration. It parks itself alongside the docked pane and
  follows it when the pane is moved or resized.

Entry point is ``show_tutorial()``. Call it from a valid Revit API context -
the ribbon right-click handler routes through an ExternalEvent for that reason,
and config.py (Shift+Click) is already in one.
"""

import os
import sys

import clr

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

from System import TimeSpan
from System.Collections.Generic import List
from System.Windows import (
    CornerRadius,
    Duration,
    HorizontalAlignment,
    Point,
    PresentationSource,
    Rect,
    SystemParameters,
    Thickness,
    UIElement,
    VerticalAlignment,
    Visibility,
    WindowState,
)
from System.Windows import Application
from System.Windows.Controls import Border, Grid, Panel
from System.Windows.Media import Color, FillRule, GeometryGroup, RectangleGeometry, SolidColorBrush
from System.Windows.Media.Animation import DoubleAnimation, RepeatBehavior
from System.Windows.Shapes import Path as ShapePath

from pyrevit import forms, script

THIS_DIR = os.path.abspath(os.path.dirname(__file__))

def _bootstrap_cedlib():
    """Put CEDLib.lib on sys.path without needing it importable first.

    startup.py normally seeds this, but the tour can also be launched from a
    fresh engine (Shift+Click) before anything else has touched CEDLib.
    """
    current = THIS_DIR
    while True:
        candidate = os.path.join(current, "CEDLib.lib")
        if os.path.isdir(candidate):
            if candidate not in sys.path:
                sys.path.insert(0, candidate)
            return candidate
        parent = os.path.dirname(current)
        if not parent or parent == current:
            return None
        current = parent


_bootstrap_cedlib()

from UIClasses import pathing as ui_pathing  # noqa: E402
from UIClasses import load_theme_state_from_config  # noqa: E402
from UIClasses import resource_loader  # noqa: E402

LIB_ROOT = ui_pathing.ensure_lib_root_on_syspath(THIS_DIR)
UI_RESOURCES_ROOT = ui_pathing.resolve_ui_resources_root(LIB_ROOT)

XAML_PATH = os.path.join(THIS_DIR, "TutorialGuideWindow.xaml")
MASCOT_PATH = os.path.join(THIS_DIR, "ColonelMascot.xaml")
CONTENT_PATH = os.path.join(THIS_DIR, "tutorial_content.py")

PANEL_MODULE_NAME = "ced_electools_circuit_manager_panel"
CONTENT_MODULE_NAME = "ced_electools_circuit_manager_tutorial_content"
PANEL_ID = "36c3fd8d-98c4-4cf4-92a4-4ac7f3f8c4f2"
_WINDOW_MARKER = "_ced_circuit_manager_tutorial_window_v1"

_LOGGER = script.get_logger()

# Friendly labels for the spotlight note, keyed by the element Name in
# CircuitBrowserPanel.xaml.
TARGET_LABELS = {
    "ActionsButton": "the Actions button",
    "BrowserOptionsButton": "the options (...) button",
    "CalcAllButton": "Calculate All",
    "CalcPreviewToggle": "the calculation preview toggle",
    "CalcSelectedButton": "Calculate Selected",
    "CalcSettingsButton": "the calculate settings gear",
    "CheckAllButton": "Check All",
    "CircuitList": "the circuit list",
    "ClearSelectionButton": "the clear-selection button",
    "ClearSearchButton": "the clear-search button",
    "DocumentNameText": "the document line",
    "FilterActiveMark": "the filter-active marker",
    "FilterButton": "the Filter button",
    "RefreshButton": "the Refresh button",
    "SearchBox": "the search box",
    "SelectCircuitsButton": "Select in Model - Circuit",
    "SelectDownstreamButton": "Select in Model - Device",
    "SelectEquipmentButton": "Select in Model - Panel",
    "ShowOneLineButton": "the one line diagram button",
    "StatusText": "the status line",
    "ToggleViewButton": "the list/card toggle",
    "UncheckAllButton": "Uncheck All",
}


# ---------------------------------------------------------------------------
# module + panel access
# ---------------------------------------------------------------------------

def _load_content_module():
    """Load the narration, re-reading it if the file changed on disk.

    Reword a step in tutorial_content.py and the next launch of the tour picks
    it up - no pyRevit reload needed. Code changes in this file still do.
    """
    import imp

    try:
        stamp = os.path.getmtime(CONTENT_PATH)
    except Exception:
        stamp = None

    module = sys.modules.get(CONTENT_MODULE_NAME)
    if module is not None:
        cached = getattr(module, "_ced_source_stamp", None)
        if stamp is None or cached == stamp:
            return module

    module = imp.load_source(CONTENT_MODULE_NAME, CONTENT_PATH)
    if stamp is not None:
        try:
            module._ced_source_stamp = stamp
        except Exception:
            pass
    return module


def _panel_instance():
    module = sys.modules.get(PANEL_MODULE_NAME)
    if module is None:
        return None
    panel_cls = getattr(module, "CircuitBrowserPanel", None)
    if panel_cls is None or not hasattr(panel_cls, "get_instance"):
        return None
    try:
        return panel_cls.get_instance()
    except Exception:
        return None


def _panel_host(panel=None):
    """Root Grid of the pane - the container the overlay is injected into."""
    if panel is None:
        panel = _panel_instance()
    if panel is None:
        return None
    host = getattr(panel, "_dock_frame_host", None)
    if host is None:
        try:
            host = panel.FindName("DockFrameHost")
        except Exception:
            host = None
    return host


def _find_named(name):
    if not name:
        return None
    panel = _panel_instance()
    if panel is not None:
        try:
            element = panel.FindName(name)
            if element is not None:
                return element
        except Exception:
            pass
    host = _panel_host(panel)
    if host is not None:
        try:
            return host.FindName(name)
        except Exception:
            return None
    return None


def _try_resource(owner, key, fallback=None):
    try:
        value = owner.TryFindResource(key)
        if value is not None:
            return value
    except Exception:
        pass
    return fallback


def _screen_rect(element):
    """Element bounds in device-independent screen coordinates, or None."""
    if element is None:
        return None
    try:
        if not bool(element.IsVisible):
            return None
        width = float(getattr(element, "ActualWidth", 0) or 0)
        height = float(getattr(element, "ActualHeight", 0) or 0)
        if width <= 0 or height <= 0:
            return None
        source = PresentationSource.FromVisual(element)
        if source is None or source.CompositionTarget is None:
            return None
        top_left = element.PointToScreen(Point(0, 0))
        bottom_right = element.PointToScreen(Point(width, height))
        matrix = source.CompositionTarget.TransformFromDevice
        first = matrix.Transform(top_left)
        second = matrix.Transform(bottom_right)
        return Rect(first.X, first.Y, second.X - first.X, second.Y - first.Y)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# spotlight overlay
# ---------------------------------------------------------------------------

class SpotlightOverlay(object):
    """Dim + cut-out highlight painted inside the Circuit Manager pane."""

    def __init__(self, on_layout=None):
        self._host = None
        self._overlay = None
        self._dim = None
        self._ring = None
        self._target_name = None
        self._layout_key = None
        self._on_layout = on_layout
        self._layout_hooked = False

    # -- lifecycle ---------------------------------------------------------

    def attach(self):
        host = _panel_host()
        if host is None:
            self.detach()
            return False
        if self._host is host and self._overlay is not None:
            return True
        self.detach()

        overlay = Grid()
        overlay.IsHitTestVisible = False
        overlay.Visibility = Visibility.Collapsed

        dim = ShapePath()
        dim.IsHitTestVisible = False
        dim.Fill = SolidColorBrush(Color.FromArgb(112, 0, 0, 0))
        overlay.Children.Add(dim)

        ring = Border()
        ring.IsHitTestVisible = False
        ring.BorderThickness = Thickness(2)
        ring.CornerRadius = CornerRadius(5)
        ring.BorderBrush = _try_resource(
            host,
            "CED.Brush.Accent",
            SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0xD1)),
        )
        ring.HorizontalAlignment = HorizontalAlignment.Left
        ring.VerticalAlignment = VerticalAlignment.Top
        overlay.Children.Add(ring)

        try:
            host.Children.Add(overlay)
            Panel.SetZIndex(overlay, 5000)
        except Exception as exc:
            _LOGGER.debug("Tutorial overlay could not be injected: %s", exc)
            return False

        self._host = host
        self._overlay = overlay
        self._dim = dim
        self._ring = ring
        self._layout_key = None
        self._start_pulse()

        try:
            host.LayoutUpdated += self._host_layout_updated
            self._layout_hooked = True
        except Exception:
            self._layout_hooked = False
        return True

    def detach(self):
        if self._host is not None and self._layout_hooked:
            try:
                self._host.LayoutUpdated -= self._host_layout_updated
            except Exception:
                pass
        self._layout_hooked = False
        self._stop_pulse()
        if self._host is not None and self._overlay is not None:
            try:
                self._host.Children.Remove(self._overlay)
            except Exception:
                pass
        self._host = None
        self._overlay = None
        self._dim = None
        self._ring = None
        self._layout_key = None

    # -- animation ---------------------------------------------------------

    def _start_pulse(self):
        if self._ring is None:
            return
        try:
            animation = DoubleAnimation(1.0, 0.35, Duration(TimeSpan.FromMilliseconds(900)))
            animation.AutoReverse = True
            animation.RepeatBehavior = RepeatBehavior.Forever
            self._ring.BeginAnimation(UIElement.OpacityProperty, animation)
        except Exception:
            pass

    def _stop_pulse(self):
        if self._ring is None:
            return
        try:
            self._ring.BeginAnimation(UIElement.OpacityProperty, None)
        except Exception:
            pass

    # -- painting ----------------------------------------------------------

    def show_target(self, target_name):
        self._target_name = target_name or None
        if not self.attach():
            return False
        self._layout_key = None
        return self._paint()

    def target_is_visible(self):
        return self._resolve_target() is not None

    def _resolve_target(self):
        if not self._target_name:
            return None
        element = _find_named(self._target_name)
        if element is None:
            return None
        try:
            if not bool(element.IsVisible):
                return None
            if float(element.ActualWidth or 0) <= 0 or float(element.ActualHeight or 0) <= 0:
                return None
        except Exception:
            return None
        return element

    def _host_layout_updated(self, sender, args):
        self._paint()
        if self._on_layout is not None:
            try:
                self._on_layout()
            except Exception:
                pass

    def _paint(self):
        host = self._host
        overlay = self._overlay
        if host is None or overlay is None:
            return False
        try:
            host_width = float(getattr(host, "ActualWidth", 0) or 0)
            host_height = float(getattr(host, "ActualHeight", 0) or 0)
        except Exception:
            return False
        if host_width <= 0 or host_height <= 0:
            return False

        target = self._resolve_target()
        rect = None
        if target is not None:
            try:
                origin = target.TransformToVisual(host).Transform(Point(0, 0))
                rect = Rect(
                    origin.X,
                    origin.Y,
                    float(target.ActualWidth),
                    float(target.ActualHeight),
                )
            except Exception:
                rect = None

        key = (
            round(host_width, 1),
            round(host_height, 1),
            None if rect is None else (
                round(rect.X, 1), round(rect.Y, 1), round(rect.Width, 1), round(rect.Height, 1)
            ),
        )
        if key == self._layout_key:
            return rect is not None
        self._layout_key = key

        if rect is None:
            overlay.Visibility = Visibility.Collapsed
            return False

        pad = 4.0
        left = max(rect.X - pad, 0.0)
        top = max(rect.Y - pad, 0.0)
        width = min(rect.Width + (pad * 2), host_width - left)
        height = min(rect.Height + (pad * 2), host_height - top)
        if width <= 0 or height <= 0:
            overlay.Visibility = Visibility.Collapsed
            return False
        hole = Rect(left, top, width, height)

        group = GeometryGroup()
        group.FillRule = FillRule.EvenOdd
        group.Children.Add(RectangleGeometry(Rect(0, 0, host_width, host_height)))
        group.Children.Add(RectangleGeometry(hole, 5, 5))
        self._dim.Data = group

        self._ring.Margin = Thickness(hole.X, hole.Y, 0, 0)
        self._ring.Width = hole.Width
        self._ring.Height = hole.Height
        overlay.Visibility = Visibility.Visible
        return True


# ---------------------------------------------------------------------------
# tutorial window
# ---------------------------------------------------------------------------

class TutorialGuideWindow(forms.WPFWindow):
    """Frameless narration window that walks the Circuit Manager pane."""

    def __init__(self, theme_mode, accent_mode, content):
        self._theme_mode = resource_loader.normalize_theme_mode(theme_mode, "light")
        self._accent_mode = resource_loader.normalize_accent_mode(accent_mode, "blue")
        self._steps = list(getattr(content, "STEPS", []) or [])
        self._chapters = list(content.chapters()) if hasattr(content, "chapters") else []
        self._index = 0
        self._suppress_jump = False
        self._last_pane_rect = None

        forms.WPFWindow.__init__(self, XAML_PATH)
        try:
            self.Tag = _WINDOW_MARKER
        except Exception:
            pass
        self.Opacity = 0.0

        self._merge_mascot_resources()
        self._apply_theme()

        self._title_bar = self.FindName("TitleBar")
        self._tour_title = self.FindName("TourTitleText")
        self._step_counter = self.FindName("StepCounterText")
        self._close_button = self.FindName("CloseButton")
        self._step_title = self.FindName("StepTitleText")
        self._step_body = self.FindName("StepBodyText")
        self._note_host = self.FindName("SpotlightNoteHost")
        self._note_text = self.FindName("SpotlightNoteText")
        self._bullets = self.FindName("BulletList")
        self._bullet_scroller = self.FindName("BulletScroller")
        self._chapter_text = self.FindName("ChapterText")
        self._progress_track = self.FindName("ProgressTrack")
        self._progress_fill = self.FindName("ProgressFill")
        self._jump_box = self.FindName("JumpBox")
        self._back_button = self.FindName("BackButton")
        self._next_button = self.FindName("NextButton")

        if self._tour_title is not None:
            self._tour_title.Text = getattr(content, "TOUR_TITLE", "Circuit Manager - Show Me Around")

        self._overlay = SpotlightOverlay(on_layout=self._pane_layout_changed)
        self._wire_events()
        self._fill_jump_box()
        self._render()

    # -- setup -------------------------------------------------------------

    def _merge_mascot_resources(self):
        if not os.path.exists(MASCOT_PATH):
            _LOGGER.debug("Colonel mascot resources missing: %s", MASCOT_PATH)
            return
        try:
            from System import Uri
            from System.Windows import ResourceDictionary

            dictionary = ResourceDictionary()
            dictionary.Source = Uri(MASCOT_PATH)
            self.Resources.MergedDictionaries.Add(dictionary)
        except Exception as exc:
            _LOGGER.debug("Colonel mascot resources failed to load: %s", exc)

    def _apply_theme(self):
        resource_loader.apply_theme(
            self,
            resources_root=UI_RESOURCES_ROOT,
            theme_mode=self._theme_mode,
            accent_mode=self._accent_mode,
        )

    def _wire_events(self):
        if self._close_button is not None:
            self._close_button.Click += self._close_clicked
        if self._back_button is not None:
            self._back_button.Click += self._back_clicked
        if self._next_button is not None:
            self._next_button.Click += self._next_clicked
        if self._title_bar is not None:
            self._title_bar.MouseLeftButtonDown += self._title_bar_mouse_down
        if self._jump_box is not None:
            self._jump_box.SelectionChanged += self._jump_changed
        if self._progress_track is not None:
            self._progress_track.SizeChanged += self._progress_track_resized
        self.Loaded += self._window_loaded
        self.SizeChanged += self._window_size_changed
        self.Closed += self._window_closed

    def _fill_jump_box(self):
        if self._jump_box is None:
            return
        self._suppress_jump = True
        try:
            self._jump_box.Items.Clear()
            for name, _start in self._chapters:
                self._jump_box.Items.Add(name)
        finally:
            self._suppress_jump = False

    # -- rendering ---------------------------------------------------------

    def _current_step(self):
        if not self._steps:
            return {}
        index = max(0, min(self._index, len(self._steps) - 1))
        return self._steps[index] or {}

    def _render(self):
        step = self._current_step()
        total = max(len(self._steps), 1)

        if self._step_title is not None:
            self._step_title.Text = str(step.get("title") or "")
        if self._step_body is not None:
            self._step_body.Text = str(step.get("body") or "")
        if self._step_counter is not None:
            self._step_counter.Text = "{0} / {1}".format(self._index + 1, total)
        if self._chapter_text is not None:
            self._chapter_text.Text = str(step.get("chapter") or "")

        if self._bullets is not None:
            bullets = List[str]()
            for line in list(step.get("bullets") or []):
                bullets.Add(str(line))
            self._bullets.ItemsSource = bullets
        if self._bullet_scroller is not None:
            try:
                self._bullet_scroller.ScrollToTop()
            except Exception:
                pass

        self._apply_spotlight(step)
        self._update_navigation()
        self._update_progress()
        self._sync_jump_box()
        self._reposition()

    def _apply_spotlight(self, step):
        target = step.get("target")
        shown = False
        if target:
            try:
                shown = self._overlay.show_target(target)
            except Exception as exc:
                _LOGGER.debug("Tutorial spotlight failed for %s: %s", target, exc)
                shown = False
        else:
            try:
                self._overlay.show_target(None)
            except Exception:
                pass

        if self._note_host is None or self._note_text is None:
            return
        if not target:
            self._note_host.Visibility = Visibility.Collapsed
            return
        label = TARGET_LABELS.get(target, target)
        if shown:
            self._note_text.Text = "Highlighted in the pane: {0}.".format(label)
        else:
            self._note_text.Text = (
                "{0} is not on screen right now - open the Circuit Manager pane to see it "
                "highlighted.".format(label[:1].upper() + label[1:])
            )
        self._note_host.Visibility = Visibility.Visible

    def _update_navigation(self):
        last = self._index >= len(self._steps) - 1
        if self._back_button is not None:
            self._back_button.IsEnabled = self._index > 0
        if self._next_button is not None:
            self._next_button.Content = "Finish" if last else "Next"

    def _update_progress(self):
        if self._progress_track is None or self._progress_fill is None:
            return
        try:
            track_width = float(self._progress_track.ActualWidth or 0)
        except Exception:
            track_width = 0.0
        if track_width <= 0:
            return
        total = max(len(self._steps), 1)
        fraction = float(self._index + 1) / float(total)
        self._progress_fill.Width = max(2.0, track_width * fraction)

    def _sync_jump_box(self):
        if self._jump_box is None or not self._chapters:
            return
        chapter_index = 0
        for position, (_name, start) in enumerate(self._chapters):
            if self._index >= start:
                chapter_index = position
        self._suppress_jump = True
        try:
            self._jump_box.SelectedIndex = chapter_index
        finally:
            self._suppress_jump = False

    # -- placement ---------------------------------------------------------

    def _reposition(self):
        try:
            width = float(self.ActualWidth or self.Width or 400)
            height = float(self.ActualHeight or 0)
        except Exception:
            return
        if height <= 0:
            height = 380.0

        screen_left = float(SystemParameters.VirtualScreenLeft)
        screen_top = float(SystemParameters.VirtualScreenTop)
        screen_right = screen_left + float(SystemParameters.VirtualScreenWidth)
        screen_bottom = screen_top + float(SystemParameters.VirtualScreenHeight)

        pane_rect = _screen_rect(_panel_host())
        gap = 6.0
        if pane_rect is None:
            left = screen_left + ((screen_right - screen_left) - width) / 2.0
            top = screen_top + ((screen_bottom - screen_top) - height) / 2.0
        else:
            room_left = pane_rect.X - screen_left
            room_right = screen_right - (pane_rect.X + pane_rect.Width)
            if room_right >= width + gap or room_right >= room_left:
                left = pane_rect.X + pane_rect.Width + gap
            else:
                left = pane_rect.X - width - gap
            top = pane_rect.Y

        left = max(screen_left, min(left, screen_right - width))
        top = max(screen_top, min(top, screen_bottom - height))
        try:
            self.Left = left
            self.Top = top
        except Exception:
            pass
        self._last_pane_rect = pane_rect

    def _pane_layout_changed(self):
        current = _screen_rect(_panel_host())
        previous = self._last_pane_rect
        if current is None and previous is None:
            return
        if current is not None and previous is not None:
            same = (
                abs(current.X - previous.X) < 1.0
                and abs(current.Y - previous.Y) < 1.0
                and abs(current.Width - previous.Width) < 1.0
                and abs(current.Height - previous.Height) < 1.0
            )
            if same:
                return
        self._reposition()

    # -- event handlers ----------------------------------------------------

    def _window_loaded(self, sender, args):
        self._update_progress()
        self._reposition()
        try:
            self.Opacity = 1.0
        except Exception:
            pass

    def _progress_track_resized(self, sender, args):
        # The fill is sized in pixels, so it has to be recomputed whenever the
        # track's own width changes (first layout pass, or a step that makes
        # the window taller/narrower).
        self._update_progress()

    def _window_size_changed(self, sender, args):
        self._update_progress()
        self._reposition()

    def _window_closed(self, sender, args):
        try:
            self._overlay.detach()
        except Exception:
            pass

    def _title_bar_mouse_down(self, sender, args):
        # DragMove throws if the button is already up by the time we get here.
        try:
            self.DragMove()
        except Exception:
            pass

    def _close_clicked(self, sender, args):
        self.Close()

    def _back_clicked(self, sender, args):
        if self._index > 0:
            self._index -= 1
            self._render()

    def _next_clicked(self, sender, args):
        if self._index >= len(self._steps) - 1:
            self.Close()
            return
        self._index += 1
        self._render()

    def _jump_changed(self, sender, args):
        if self._suppress_jump or not self._chapters:
            return
        try:
            position = int(self._jump_box.SelectedIndex)
        except Exception:
            return
        if position < 0 or position >= len(self._chapters):
            return
        self._index = self._chapters[position][1]
        self._render()


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------

def _find_existing_window():
    app = Application.Current
    if app is None:
        return None
    try:
        windows = list(app.Windows)
    except Exception:
        return None
    for window in windows:
        try:
            if str(getattr(window, "Tag", "") or "") == _WINDOW_MARKER:
                return window
        except Exception:
            continue
    return None


def _ensure_pane_open():
    """Open the dockable pane so there is something to spotlight."""
    try:
        forms.open_dockable_panel(PANEL_ID)
    except Exception as exc:
        _LOGGER.debug("Circuit Manager pane could not be opened for the tour: %s", exc)
        return False
    panel = _panel_instance()
    if panel is not None and hasattr(panel, "refresh_on_open"):
        try:
            panel.refresh_on_open()
        except Exception:
            pass
    return True


def show_tutorial():
    """Open (or focus) the guided tour. Safe to call repeatedly."""
    _ensure_pane_open()

    existing = _find_existing_window()
    if existing is not None:
        try:
            if getattr(existing, "WindowState", None) == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
            existing.Show()
            existing.Activate()
        except Exception:
            pass
        return existing

    content = _load_content_module()
    theme_mode, accent_mode = load_theme_state_from_config(
        default_theme="light",
        default_accent="blue",
    )
    window = TutorialGuideWindow(theme_mode, accent_mode, content)
    window.Show()
    try:
        window.Activate()
    except Exception:
        pass
    return window
