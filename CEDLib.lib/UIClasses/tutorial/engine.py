# -*- coding: utf-8 -*-
"""Shared "Show Me Around" tour engine for the AE pyTools ribbon.

One engine, many tours. A tool contributes only a ``tutorial_content.py`` next
to its bundle and a two-line ``config.py``; everything below is shared.

Content contract - see UIClasses/tutorial/__init__.py for the full docstring:

    TOUR_TITLE    str
    TOUR_KEY      str, unique - used for the window singleton
    HOST          dict or None, describes what to spotlight
    TARGET_LABELS dict of element name -> friendly phrase (optional)
    STEPS         list of {chapter, title, target, body, bullets}

A step whose ``target`` cannot be resolved degrades to narration-only rather
than failing, so a tour still reads correctly with the tool closed.
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
    Application,
    CornerRadius,
    Duration,
    HorizontalAlignment,
    Point,
    PresentationSource,
    Rect,
    ResourceDictionary,
    SystemParameters,
    Thickness,
    UIElement,
    VerticalAlignment,
    Visibility,
    Window,
    WindowState,
)
from System.Windows.Controls import Border, Grid, Panel
from System.Windows.Media import Color, FillRule, GeometryGroup, RectangleGeometry, SolidColorBrush
from System.Windows.Media.Animation import DoubleAnimation, RepeatBehavior
from System.Windows.Shapes import Path as ShapePath

from pyrevit import forms, script

from UIClasses import load_theme_state_from_config
from UIClasses import pathing as ui_pathing
from UIClasses import resource_loader

TUTORIAL_DIR = os.path.abspath(os.path.dirname(__file__))
XAML_PATH = os.path.join(TUTORIAL_DIR, "TourWindow.xaml")
MASCOT_PATH = os.path.join(TUTORIAL_DIR, "ColonelMascot.xaml")

LIB_ROOT = ui_pathing.resolve_lib_root(TUTORIAL_DIR) or os.path.abspath(
    os.path.join(TUTORIAL_DIR, "..", "..")
)
UI_RESOURCES_ROOT = ui_pathing.resolve_ui_resources_root(LIB_ROOT)

WINDOW_MARKER_PREFIX = "_ced_tour_window_v1::"

_LOGGER = script.get_logger()


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------

def load_module(module_name, path):
    """Load a module from an explicit path on any engine pyRevit ships.

    importlib first (IronPython 3.x / CPython), imp as the fallback
    (IronPython 2.7). ``imp`` was removed in Python 3.12, so it cannot be the
    only path if these tools ever run on the CPython engine.
    """
    cached = sys.modules.get(module_name)
    stamp = None
    try:
        stamp = os.path.getmtime(path)
    except Exception:
        stamp = None
    if cached is not None:
        if stamp is None or getattr(cached, "_ced_source_stamp", None) == stamp:
            return cached

    module = None
    try:
        import importlib.util

        spec = importlib.util.spec_from_file_location(module_name, path)
        if spec is None:
            raise ImportError("no spec for {0}".format(path))
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
    except (ImportError, AttributeError):
        import imp

        module = imp.load_source(module_name, path)

    if module is not None and stamp is not None:
        try:
            module._ced_source_stamp = stamp
        except Exception:
            pass
    return module


def _try_resource(owner, key, fallback=None):
    if owner is None or not key:
        return fallback
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


def _default_accent():
    """The house default accent, from the shared catalog when it exists."""
    try:
        from UIClasses import theme_manager

        return getattr(theme_manager, "DEFAULT_ACCENT_MODE", "blue")
    except Exception:
        return "blue"


def open_windows():
    """Every open WPF window in this process.

    Revit is not a WPF application, so ``Application.Current`` is null and
    ``Application.Current.Windows`` is unusable - pyRevit itself never touches
    it, going through HwndSource instead. Each top-level WPF window has its own
    PresentationSource whose RootVisual is the Window, and those we can see.
    """
    found = []
    try:
        for source in PresentationSource.CurrentSources:
            try:
                root = source.RootVisual
            except Exception:
                continue
            if isinstance(root, Window) and root not in found:
                found.append(root)
    except Exception as exc:
        _LOGGER.debug("Tour could not enumerate presentation sources: %s", exc)
    # If something in the process did create an Application, union it in.
    try:
        app = Application.Current
        if app is not None:
            for window in list(app.Windows):
                if window not in found:
                    found.append(window)
    except Exception:
        pass
    return found


def _visible_with_size(element):
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


# ---------------------------------------------------------------------------
# hosts : where a tour points its spotlight
# ---------------------------------------------------------------------------

class NoHost(object):
    """Narration-only tour. Nothing to open, nothing to highlight."""

    kind = "none"

    def ensure_open(self):
        return False

    def host_panel(self):
        return None

    def find_named(self, name):
        return None

    def diagnose(self, name):
        return None


class DockablePanelHost(object):
    """Spotlight a pyRevit dockable pane (Circuit Manager and friends)."""

    kind = "dockable_panel"

    def __init__(self, panel_module, panel_id, host_element="DockFrameHost"):
        self.panel_module = panel_module
        self.panel_id = panel_id
        self.host_element = host_element or "DockFrameHost"

    def _panel(self):
        module = sys.modules.get(self.panel_module)
        if module is None:
            return None
        for attribute in dir(module):
            candidate = getattr(module, attribute, None)
            if candidate is None or not hasattr(candidate, "get_instance"):
                continue
            try:
                instance = candidate.get_instance()
            except Exception:
                continue
            if instance is not None:
                return instance
        return None

    def ensure_open(self):
        if not self.panel_id:
            return False
        try:
            forms.open_dockable_panel(self.panel_id)
        except Exception as exc:
            _LOGGER.debug("Tour could not open dockable pane %s: %s", self.panel_id, exc)
            return False
        panel = self._panel()
        if panel is not None and hasattr(panel, "refresh_on_open"):
            try:
                panel.refresh_on_open()
            except Exception:
                pass
        return True

    def host_panel(self):
        panel = self._panel()
        if panel is None:
            return None
        host = getattr(panel, "_dock_frame_host", None)
        if host is None:
            try:
                host = panel.FindName(self.host_element)
            except Exception:
                host = None
        return host if isinstance(host, Panel) else None

    def find_named(self, name):
        if not name:
            return None
        panel = self._panel()
        if panel is None:
            return None
        try:
            return panel.FindName(name)
        except Exception:
            return None

    def diagnose(self, name):
        """Why the spotlight could not land. None means it should have."""
        if self._panel() is None:
            return "the pane is not loaded - try reloading pyRevit"
        if self.host_panel() is None:
            return "the pane's layout root ({0}) was not found".format(self.host_element)
        element = self.find_named(name)
        if element is None:
            return "there is no '{0}' in this build of the pane".format(name)
        if _visible_with_size(element) is None:
            return "it is hidden or collapsed in the pane right now"
        return None


class WindowHost(object):
    """Spotlight a modeless window the tool opens, matched by Tag or Title."""

    kind = "window"

    def __init__(self, tag=None, title=None, host_element=None):
        self.tag = tag or None
        self.title = title or None
        self.host_element = host_element or None

    def _window(self):
        for window in open_windows():
            try:
                if self.tag and str(getattr(window, "Tag", "") or "") == self.tag:
                    return window
                if self.title and str(getattr(window, "Title", "") or "") == self.title:
                    return window
            except Exception:
                continue
        return None

    def ensure_open(self):
        return self._window() is not None

    def host_panel(self):
        window = self._window()
        if window is None:
            return None
        if self.host_element:
            try:
                named = window.FindName(self.host_element)
                if isinstance(named, Panel):
                    return named
            except Exception:
                pass
        try:
            content = window.Content
        except Exception:
            return None
        return content if isinstance(content, Panel) else None

    def find_named(self, name):
        window = self._window()
        if window is None or not name:
            return None
        try:
            return window.FindName(name)
        except Exception:
            return None

    def diagnose(self, name):
        """Why the spotlight could not land. None means it should have."""
        window = self._window()
        if window is None:
            return "the {0} window is not open".format(self.title or "tool")
        if self.host_panel() is None:
            return "that window's layout root is not a container we can draw into"
        element = self.find_named(name)
        if element is None:
            return "there is no '{0}' in this build of the window".format(name)
        if _visible_with_size(element) is None:
            return "it is hidden or collapsed in that window right now"
        return None


def build_host(spec):
    """Turn a content module's HOST dict into a host object."""
    if not spec:
        return NoHost()
    kind = str(spec.get("kind") or "").strip().lower()
    try:
        if kind == "dockable_panel":
            return DockablePanelHost(
                panel_module=spec.get("panel_module"),
                panel_id=spec.get("panel_id"),
                host_element=spec.get("host_element"),
            )
        if kind == "window":
            return WindowHost(
                tag=spec.get("tag"),
                title=spec.get("title"),
                host_element=spec.get("host_element"),
            )
    except Exception as exc:
        _LOGGER.debug("Tour host spec rejected (%s): %s", spec, exc)
    return NoHost()


# ---------------------------------------------------------------------------
# spotlight overlay
# ---------------------------------------------------------------------------

class SpotlightOverlay(object):
    """Dim + cut-out highlight injected into the host's own visual tree.

    The overlay is added as one extra child of a Grid/Panel the tool already
    has, and it is never hit-test visible - so the user can keep using the tool
    while the tour is running, and nothing in the host has to change.
    """

    def __init__(self, host, on_layout=None):
        self._host = host or NoHost()
        self._panel = None
        self._overlay = None
        self._dim = None
        self._ring = None
        self._target_name = None
        self._layout_key = None
        self._on_layout = on_layout
        self._layout_hooked = False

    def attach(self):
        panel = None
        try:
            panel = self._host.host_panel()
        except Exception:
            panel = None
        if panel is None:
            self.detach()
            return False
        if self._panel is panel and self._overlay is not None:
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
            panel, "CED.Brush.Accent", SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0xD1))
        )
        ring.HorizontalAlignment = HorizontalAlignment.Left
        ring.VerticalAlignment = VerticalAlignment.Top
        overlay.Children.Add(ring)

        try:
            panel.Children.Add(overlay)
            Panel.SetZIndex(overlay, 5000)
            self._span_all_cells(panel, overlay)
        except Exception as exc:
            _LOGGER.debug("Tour overlay could not be injected: %s", exc)
            return False

        self._panel = panel
        self._overlay = overlay
        self._dim = dim
        self._ring = ring
        self._layout_key = None
        self._start_pulse()
        try:
            panel.LayoutUpdated += self._panel_layout_updated
            self._layout_hooked = True
        except Exception:
            self._layout_hooked = False
        return True

    @staticmethod
    def _span_all_cells(panel, overlay):
        """Cover the whole Grid, not just cell 0,0."""
        if not isinstance(panel, Grid):
            return
        try:
            rows = max(1, panel.RowDefinitions.Count)
            columns = max(1, panel.ColumnDefinitions.Count)
            Grid.SetRow(overlay, 0)
            Grid.SetColumn(overlay, 0)
            Grid.SetRowSpan(overlay, rows)
            Grid.SetColumnSpan(overlay, columns)
        except Exception:
            pass

    def detach(self):
        if self._panel is not None and self._layout_hooked:
            try:
                self._panel.LayoutUpdated -= self._panel_layout_updated
            except Exception:
                pass
        self._layout_hooked = False
        self._stop_pulse()
        if self._panel is not None and self._overlay is not None:
            try:
                self._panel.Children.Remove(self._overlay)
            except Exception:
                pass
        self._panel = None
        self._overlay = None
        self._dim = None
        self._ring = None
        self._layout_key = None

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

    def show_target(self, target_name):
        self._target_name = target_name or None
        if not self.attach():
            return False
        self._layout_key = None
        return self._paint()

    def _resolve_target(self):
        if not self._target_name:
            return None
        try:
            return _visible_with_size(self._host.find_named(self._target_name))
        except Exception:
            return None

    def _panel_layout_updated(self, sender, args):
        self._paint()
        if self._on_layout is not None:
            try:
                self._on_layout()
            except Exception:
                pass

    def _paint(self):
        panel = self._panel
        overlay = self._overlay
        if panel is None or overlay is None:
            return False
        try:
            panel_width = float(getattr(panel, "ActualWidth", 0) or 0)
            panel_height = float(getattr(panel, "ActualHeight", 0) or 0)
        except Exception:
            return False
        if panel_width <= 0 or panel_height <= 0:
            return False

        target = self._resolve_target()
        rect = None
        if target is not None:
            try:
                origin = target.TransformToVisual(panel).Transform(Point(0, 0))
                rect = Rect(
                    origin.X, origin.Y, float(target.ActualWidth), float(target.ActualHeight)
                )
            except Exception:
                rect = None

        key = (
            round(panel_width, 1),
            round(panel_height, 1),
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
        width = min(rect.Width + (pad * 2), panel_width - left)
        height = min(rect.Height + (pad * 2), panel_height - top)
        if width <= 0 or height <= 0:
            overlay.Visibility = Visibility.Collapsed
            return False
        hole = Rect(left, top, width, height)

        group = GeometryGroup()
        group.FillRule = FillRule.EvenOdd
        group.Children.Add(RectangleGeometry(Rect(0, 0, panel_width, panel_height)))
        group.Children.Add(RectangleGeometry(hole, 5, 5))
        self._dim.Data = group

        self._ring.Margin = Thickness(hole.X, hole.Y, 0, 0)
        self._ring.Width = hole.Width
        self._ring.Height = hole.Height
        overlay.Visibility = Visibility.Visible
        return True


# ---------------------------------------------------------------------------
# tour window
# ---------------------------------------------------------------------------

class TourWindow(forms.WPFWindow):
    """Frameless narration window carrying the Colonel."""

    def __init__(self, content, theme_mode=None, accent_mode=None):
        self._content = content
        self._steps = list(getattr(content, "STEPS", []) or [])
        self._labels = dict(getattr(content, "TARGET_LABELS", {}) or {})
        self._tour_key = str(getattr(content, "TOUR_KEY", "") or "tour")
        self._host = build_host(getattr(content, "HOST", None))
        self._chapters = self._build_chapters()
        self._index = 0
        self._suppress_jump = False
        self._last_host_rect = None
        self._user_moved = False

        if theme_mode is None or accent_mode is None:
            theme_mode, accent_mode = load_theme_state_from_config(
                default_theme="light", default_accent=_default_accent()
            )
        self._theme_mode = resource_loader.normalize_theme_mode(theme_mode, "light")
        self._accent_mode = resource_loader.normalize_accent_mode(accent_mode, "blue")

        forms.WPFWindow.__init__(self, XAML_PATH)
        try:
            self.Tag = WINDOW_MARKER_PREFIX + self._tour_key
        except Exception:
            pass
        self.Opacity = 0.0

        self._merge_mascot_resources()
        self._apply_theme()
        self._bind_elements()
        self._overlay = SpotlightOverlay(self._host, on_layout=self._host_layout_changed)
        self._wire_events()
        self._fill_jump_box()
        self._render()

    # -- setup ------------------------------------------------------------

    def _build_chapters(self):
        found = []
        for index, step in enumerate(self._steps):
            name = (step or {}).get("chapter") or ""
            if not found or found[-1][0] != name:
                found.append((name, index))
        return found

    def _merge_mascot_resources(self):
        if not os.path.exists(MASCOT_PATH):
            _LOGGER.debug("Colonel mascot resources missing: %s", MASCOT_PATH)
            return
        try:
            from System import Uri

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

    def _bind_elements(self):
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
            self._tour_title.Text = str(
                getattr(self._content, "TOUR_TITLE", "Show Me Around")
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
            if len(self._chapters) < 2:
                self._jump_box.Visibility = Visibility.Collapsed
        finally:
            self._suppress_jump = False

    # -- rendering --------------------------------------------------------

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

    def _apply_spotlight(self, step):
        target = step.get("target")
        shown = False
        try:
            shown = self._overlay.show_target(target)
        except Exception as exc:
            _LOGGER.debug("Tour spotlight failed for %s: %s", target, exc)
            shown = False

        if self._note_host is None or self._note_text is None:
            return
        if not target:
            self._note_host.Visibility = Visibility.Collapsed
            return
        label = self._labels.get(target, target)
        if shown:
            self._note_text.Text = "Highlighted for you: {0}.".format(label)
        else:
            reason = None
            try:
                reason = self._host.diagnose(target)
            except Exception:
                reason = None
            if reason:
                self._note_text.Text = "Can't highlight {0} - {1}.".format(label, reason)
            else:
                self._note_text.Text = (
                    "{0} is not on screen right now - open the tool to see it "
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
        self._progress_fill.Width = max(
            2.0, track_width * (float(self._index + 1) / float(total))
        )

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

    # -- placement --------------------------------------------------------

    def _reposition(self, force=False):
        """Park the window beside the host.

        Only ever called for the initial placement, and afterwards to follow the
        host pane if it is moved or resized. Once the user has dragged the
        window it stays exactly where they put it - stepping through the tour
        must never move it.
        """
        if self._user_moved and not force:
            return
        try:
            width = float(self.ActualWidth or self.Width or 400)
            height = float(self.ActualHeight or self.Height or 0)
        except Exception:
            return
        if height <= 0:
            height = 544.0

        screen_left = float(SystemParameters.VirtualScreenLeft)
        screen_top = float(SystemParameters.VirtualScreenTop)
        screen_right = screen_left + float(SystemParameters.VirtualScreenWidth)
        screen_bottom = screen_top + float(SystemParameters.VirtualScreenHeight)

        host_rect = None
        try:
            host_rect = _screen_rect(self._host.host_panel())
        except Exception:
            host_rect = None

        gap = 6.0
        if host_rect is None:
            left = screen_left + ((screen_right - screen_left) - width) / 2.0
            top = screen_top + ((screen_bottom - screen_top) - height) / 3.0
        else:
            room_left = host_rect.X - screen_left
            room_right = screen_right - (host_rect.X + host_rect.Width)
            if room_right >= width + gap or room_right >= room_left:
                left = host_rect.X + host_rect.Width + gap
            else:
                left = host_rect.X - width - gap
            top = host_rect.Y

        left = max(screen_left, min(left, screen_right - width))
        top = max(screen_top, min(top, screen_bottom - height))
        try:
            self.Left = left
            self.Top = top
        except Exception:
            pass
        self._last_host_rect = host_rect

    def _host_layout_changed(self):
        try:
            current = _screen_rect(self._host.host_panel())
        except Exception:
            return
        previous = self._last_host_rect
        if current is None and previous is None:
            return
        if current is not None and previous is not None:
            if (
                abs(current.X - previous.X) < 1.0
                and abs(current.Y - previous.Y) < 1.0
                and abs(current.Width - previous.Width) < 1.0
                and abs(current.Height - previous.Height) < 1.0
            ):
                return
        self._reposition()

    # -- event handlers ---------------------------------------------------

    def _window_loaded(self, sender, args):
        self._update_progress()
        self._reposition()
        try:
            self.Opacity = 1.0
        except Exception:
            pass

    def _window_size_changed(self, sender, args):
        self._update_progress()

    def _window_closed(self, sender, args):
        try:
            self._overlay.detach()
        except Exception:
            pass

    def _progress_track_resized(self, sender, args):
        self._update_progress()

    def _title_bar_mouse_down(self, sender, args):
        before = self._position()
        # DragMove blocks until the drag finishes, and throws if the button is
        # already up by the time we get here.
        try:
            self.DragMove()
        except Exception:
            return
        after = self._position()
        if before is None or after is None:
            self._user_moved = True
            return
        if abs(after[0] - before[0]) > 0.5 or abs(after[1] - before[1]) > 0.5:
            self._user_moved = True

    def _position(self):
        try:
            return (float(self.Left), float(self.Top))
        except Exception:
            return None

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
        if 0 <= position < len(self._chapters):
            self._index = self._chapters[position][1]
            self._render()


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------

def _find_existing_window(tour_key):
    marker = WINDOW_MARKER_PREFIX + str(tour_key or "tour")
    for window in open_windows():
        try:
            if str(getattr(window, "Tag", "") or "") == marker:
                return window
        except Exception:
            continue
    return None


def run_tour(content, modal=False):
    """Open (or focus) the tour described by a loaded content module.

    ``modal`` is for hosts that are themselves modal dialogs: a modeless window
    opened while a modal dialog is up cannot take input, so the tour has to be
    modal too or it appears frozen.
    """
    steps = list(getattr(content, "STEPS", []) or [])
    if not steps:
        forms.alert(
            "This tour has no content yet.",
            title=str(getattr(content, "TOUR_TITLE", "Show Me Around")),
        )
        return None

    host = build_host(getattr(content, "HOST", None))
    try:
        host.ensure_open()
    except Exception as exc:
        _LOGGER.debug("Tour host could not be opened: %s", exc)

    tour_key = str(getattr(content, "TOUR_KEY", "") or "tour")
    existing = _find_existing_window(tour_key)
    if existing is not None:
        try:
            if getattr(existing, "WindowState", None) == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
            existing.Show()
            existing.Activate()
        except Exception:
            pass
        return existing

    window = TourWindow(content)
    if modal:
        try:
            window.ShowDialog()
            return window
        except Exception as exc:
            _LOGGER.debug("Modal tour failed, falling back to modeless: %s", exc)
    window.Show()
    try:
        window.Activate()
    except Exception:
        pass
    return window
