# -*- coding: utf-8 -*-
"""Lightweight composited WPF control for :mod:`structured_search`.

``StructuredSearchBox`` is intentionally assembled from ordinary WPF
controls.  A host adds the control to its XAML-created panel/grid and registers
``SearchFilterDefinition`` objects; the host then consumes ``Query`` through a
query-changed handler.  No item model, Revit API type, or matching policy is
known here.

Example host wiring::

    definitions = [
        SearchFilterDefinition("status", "Status", matcher=matches_status),
    ]
    search = StructuredSearchBox(definitions, placeholder="Search items")
    search.add_query_changed_handler(on_query_changed)
    SearchHost.Children.Add(search)

The class uses the existing CED search and list resources when the host has
loaded the normal ``UIClasses/Resources`` dictionaries.  It still has modest
property fallbacks so it can be embedded in a host with a custom resource
selection.
"""

from __future__ import print_function

import clr
from System import Action

for _assembly in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass

from System.Windows import (
    CornerRadius,
    FrameworkElement,
    GridLength,
    GridUnitType,
    TextAlignment,
    Thickness,
    VerticalAlignment,
    Visibility,
)
from System.Windows.Controls import (
    Border,
    Button,
    ColumnDefinition,
    Dock,
    DockPanel,
    Grid,
    ListBox,
    ListBoxItem,
    Orientation,
    ScrollBarVisibility,
    ScrollViewer,
    StackPanel,
    TextBlock,
    TextBox,
    WrapPanel,
)
from System.Windows.Controls.Primitives import PlacementMode, Popup
from System.Windows.Input import Key, Keyboard
from System.Windows.Media import Brushes, Stretch, VisualTreeHelper
from System.Windows.Shapes import Path

from UIClasses import resource_loader
from UIClasses.structured_search import (
    SearchFilterDefinition,
    StructuredSearchState,
)


def _set_style(control, owner, key):
    lookup_owner = getattr(owner, "_resource_owner", None) or owner
    try:
        style = resource_loader.try_find_resource(lookup_owner, key)
        if style is not None:
            control.Style = style
            return
    except Exception:
        pass
    try:
        # Keep the same live lookup behavior as a XAML DynamicResource.  This
        # matters for controls created before they are attached to their host:
        # a concrete lookup at construction time can miss the host's merged
        # dictionaries, especially in newer Revit WPF hosts.
        control.SetResourceReference(FrameworkElement.StyleProperty, key)
        return
    except Exception:
        pass
    try:
        style = resource_loader.try_find_resource(lookup_owner, key)
        if style is not None:
            control.Style = style
    except Exception:
        pass


def _set_dynamic_brush(control, owner, property_name, resource_key):
    lookup_owner = getattr(owner, "_resource_owner", None) or owner
    try:
        brush = resource_loader.try_find_resource(lookup_owner, resource_key)
        if brush is not None:
            setattr(control, property_name, brush)
            return
    except Exception:
        pass
    try:
        dependency_property = getattr(control, property_name + "Property")
        control.SetResourceReference(dependency_property, resource_key)
        return
    except Exception:
        pass
    try:
        brush = resource_loader.try_find_resource(lookup_owner, resource_key)
        if brush is not None:
            setattr(control, property_name, brush)
    except Exception:
        pass


def _find_visual_ancestor(start, target_type):
    current = start
    while current is not None:
        if isinstance(current, target_type):
            return current
        try:
            current = VisualTreeHelper.GetParent(current)
        except Exception:
            return None
    return None


class StructuredSearchBox(Grid):
    """A reusable WPF search input with slash-command filter tokens.

    Public properties:

    * ``Query`` / ``query``: effective :class:`SearchQuery`.
    * ``FilterDefinitions``: cached host-provided command definitions.
    * ``IsCommandMode`` and ``CommandSuggestions``: current picker state.

    Public events are exposed through ``add_*_handler`` methods because this
    control is created from Python in pyRevit rather than declared as a WPF
    compiled custom control.  Handlers receive ``(sender, event_args)``.
    """

    def __init__(
        self,
        filter_definitions=None,
        placeholder="Search",
        input_min_width=72.0,
        state=None,
        active_placeholder="Type / to add a search field",
    ):
        Grid.__init__(self)
        self.Focusable = True
        # A transparent panel is visually inert but keeps the entire control
        # bounds hit-testable, including the empty area after the input text.
        self.Background = Brushes.Transparent
        self.MinHeight = 24.0
        self._placeholder_text = str(placeholder or "Search")
        self._active_placeholder_text = str(
            active_placeholder or "Type / to add a search field"
        )
        self._input_min_width = float(input_min_width or 0.0)
        self._suppress_input_events = False
        self._suggestion_index = 0
        self._token_editor = None
        self._resource_owner = None
        self._query_handlers = []
        self._command_handlers = []
        self._state = state or StructuredSearchState(filter_definitions or ())
        if not isinstance(self._state, StructuredSearchState):
            raise TypeError("state must be StructuredSearchState")
        if state is not None and filter_definitions is not None:
            self._state.set_filter_definitions(filter_definitions)

        self._build_visual_tree()
        self.PreviewMouseLeftButtonDown += self._search_surface_mouse_down
        self.PreviewKeyDown += self._search_surface_preview_key_down
        self._state.add_query_changed_handler(self._state_query_changed)
        self._state.add_command_changed_handler(self._state_command_changed)
        self._state.add_interaction_changed_handler(self._state_interaction_changed)
        try:
            self.Loaded += self._loaded
        except Exception:
            pass
        self._apply_resources()
        self._render_tokens()
        self._update_chrome()

    @property
    def state(self):
        return self._state

    @property
    def query(self):
        return self._state.query

    @property
    def Query(self):
        return self.query

    @property
    def filter_definitions(self):
        return self._state.filter_definitions

    @property
    def FilterDefinitions(self):
        return self.filter_definitions

    @property
    def tokens(self):
        return self._state.tokens

    @property
    def IsCommandMode(self):
        return self._state.is_command_mode

    @property
    def CommandSuggestions(self):
        return self._state.command_suggestions

    @property
    def Text(self):
        """Compatibility view of the ordinary free-text portion."""
        return self._state.free_text

    @Text.setter
    def Text(self, value):
        self.set_free_text(value)

    def add_query_changed_handler(self, handler):
        if callable(handler) and handler not in self._query_handlers:
            self._query_handlers.append(handler)
        return handler

    def remove_query_changed_handler(self, handler):
        if handler in self._query_handlers:
            self._query_handlers.remove(handler)

    def add_command_changed_handler(self, handler):
        if callable(handler) and handler not in self._command_handlers:
            self._command_handlers.append(handler)
        return handler

    def remove_command_changed_handler(self, handler):
        if handler in self._command_handlers:
            self._command_handlers.remove(handler)

    def set_filter_definitions(self, filter_definitions):
        self._state.set_filter_definitions(filter_definitions)
        self._render_tokens()
        self._update_command_popup()

    def set_free_text(self, value):
        self._state.set_free_text(value)
        self._sync_input_text()
        self._update_chrome()

    def clear(self):
        self._state.clear()
        self._sync_input_text()
        self._render_tokens()
        self._update_command_popup()
        self._update_chrome()

    def focus_input(self):
        try:
            self._input.Focus()
            try:
                Keyboard.Focus(self._input)
            except Exception:
                pass
            return True
        except Exception:
            return False

    def select_all(self):
        """Focus and select the ordinary free-text portion of the control."""
        if not self.focus_input():
            return False
        try:
            self._input.SelectAll()
            return True
        except Exception:
            return False

    def refresh_resources(self, resource_owner=None):
        """Reapply theme resources after a host changes its active theme."""
        if resource_owner is not None:
            self._resource_owner = resource_owner
        self._apply_resources()
        self._update_command_popup()
        self._update_chrome()

    def _invoke_later(self, callback):
        if callback is None:
            return False
        try:
            dispatcher = getattr(self, "Dispatcher", None)
            if dispatcher is not None:
                dispatcher.BeginInvoke(Action(callback))
                return True
        except Exception:
            pass
        try:
            callback()
            return True
        except Exception:
            return False

    def _focus_input_deferred(self, caret_index=None):
        def _focus():
            if not self.focus_input():
                return
            if caret_index is not None:
                try:
                    text_length = len(self._input.Text or "")
                    self._input.CaretIndex = max(
                        0, min(int(caret_index), text_length)
                    )
                except Exception:
                    pass

        self._invoke_later(_focus)

    def _focus_active_editor_deferred(self):
        self._invoke_later(self._focus_active_editor)

    def _build_visual_tree(self):
        self._chrome = Border()
        self._content = DockPanel()
        self._chrome.Child = self._content
        self.Children.Add(self._chrome)

        self._search_icon = Path()
        self._search_icon.Width = 14.0
        self._search_icon.Height = 14.0
        self._search_icon.Margin = Thickness(8, 0, 6, 0)
        self._search_icon.VerticalAlignment = VerticalAlignment.Center
        self._search_icon.Stretch = Stretch.Uniform
        DockPanel.SetDock(self._search_icon, Dock.Left)
        self._content.Children.Add(self._search_icon)

        self._clear_button = Button()
        self._clear_button.ToolTip = "Clear search"
        self._clear_button.Click += self._clear_clicked
        _set_style(self._clear_button, self, "CED.SearchBox.ClearButton")
        DockPanel.SetDock(self._clear_button, Dock.Right)
        self._clear_icon = Path()
        self._clear_icon.Width = 10.0
        self._clear_icon.Height = 10.0
        self._clear_icon.Stretch = Stretch.Uniform
        self._clear_button.Content = self._clear_icon
        self._content.Children.Add(self._clear_button)

        self._input_surface = Grid()
        self._input_surface.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(self._input_surface, Dock.Left)
        self._content.Children.Add(self._input_surface)

        self._scroll = ScrollViewer()
        self._scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        self._scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
        self._scroll.VerticalAlignment = VerticalAlignment.Center
        self._tokens_panel = WrapPanel()
        self._tokens_panel.Orientation = Orientation.Horizontal
        self._tokens_panel.VerticalAlignment = VerticalAlignment.Center
        self._scroll.Content = self._tokens_panel
        self._input_surface.Children.Add(self._scroll)

        self._placeholder = TextBlock()
        self._placeholder.Text = self._placeholder_text
        self._placeholder.Margin = Thickness(2, 0, 0, 0)
        self._placeholder.VerticalAlignment = VerticalAlignment.Center
        self._placeholder.IsHitTestVisible = False
        _set_style(self._placeholder, self, "CED.SearchBox.Placeholder")
        self._input_surface.Children.Add(self._placeholder)

        self._input = TextBox()
        self._input.MinWidth = self._input_min_width
        self._input.Height = 22.0
        self._input.VerticalAlignment = VerticalAlignment.Center
        self._input.VerticalContentAlignment = VerticalAlignment.Center
        self._input.TextChanged += self._input_text_changed
        self._input.GotFocus += self._input_focus_changed
        self._input.LostFocus += self._input_focus_changed
        _set_style(self._input, self, "CED.SearchBox.Input")

        self._popup = Popup()
        self._popup.PlacementTarget = self
        self._popup.Placement = PlacementMode.Bottom
        self._popup.StaysOpen = False
        self._popup.AllowsTransparency = True
        self._popup_border = Border()
        self._popup_border.Padding = Thickness(0)
        self._popup_border.BorderThickness = Thickness(1)
        self._popup_border.CornerRadius = CornerRadius(5.0)
        self._popup_border.MinWidth = 220.0
        self._popup_list = ListBox()
        self._popup_list.MaxHeight = 220.0
        self._popup_list.MinWidth = 220.0
        self._popup_list.MouseLeftButtonUp += self._suggestion_clicked
        self._popup_list.PreviewKeyDown += self._popup_preview_key_down
        self._popup.Closed += self._popup_closed
        _set_style(self._popup_list, self, "CED.ListBox")
        # The outer Border is the popup's only visible frame.  Keeping the
        # ListBox transparent prevents its template from creating a second
        # inset frame around the suggestions.
        self._popup_list.Background = Brushes.Transparent
        self._popup_list.BorderBrush = Brushes.Transparent
        self._popup_list.BorderThickness = Thickness(0)
        _set_dynamic_brush(
            self._popup_border,
            self,
            "Background",
            "CED.Brush.ListBackground",
        )
        _set_dynamic_brush(
            self._popup_border,
            self,
            "BorderBrush",
            "CED.Brush.ListBorder",
        )
        self._popup_footer = Border()
        self._popup_footer.Padding = Thickness(6, 2, 6, 3)
        self._popup_footer.Height = 20.0
        self._popup_footer.MinHeight = 20.0
        self._popup_footer.CornerRadius = CornerRadius(0, 0, 5.0, 5.0)
        self._popup_footer.BorderThickness = Thickness(0, 1, 0, 0)
        _set_dynamic_brush(
            self._popup_footer,
            self,
            "Background",
            "CED.Brush.Background",
        )
        _set_dynamic_brush(
            self._popup_footer,
            self,
            "BorderBrush",
            "CED.Brush.ListBorder",
        )
        footer_text = TextBlock()
        footer_text.Text = "Press Enter to add"
        footer_text.FontSize = 9.0
        footer_text.Opacity = 0.75
        footer_text.VerticalAlignment = VerticalAlignment.Center
        _set_dynamic_brush(
            footer_text,
            self,
            "Foreground",
            "CED.Brush.SecondaryText",
        )
        self._popup_footer_text = footer_text
        self._popup_footer.Child = footer_text
        self._popup_content = StackPanel()
        self._popup_content.Children.Add(self._popup_list)
        self._popup_content.Children.Add(self._popup_footer)
        self._popup_border.Child = self._popup_content
        self._popup.Child = self._popup_border

    def _loaded(self, sender, args):
        self._apply_resources()
        self._render_tokens()
        self._update_chrome()

    def _apply_resources(self):
        # The host is the reliable resource scope for a dynamically inserted
        # control.  ``_set_style`` and ``_set_dynamic_brush`` use this owner
        # when present, then retain dynamic-reference fallbacks for standalone
        # use.
        _set_style(self._chrome, self, "CED.SearchBox.Chrome")
        _set_style(self._input, self, "CED.SearchBox.Input")
        _set_style(self._placeholder, self, "CED.SearchBox.Placeholder")
        _set_style(self._clear_button, self, "CED.SearchBox.ClearButton")
        _set_style(self._popup_list, self, "CED.ListBox")
        # The popup shell owns the frame and list background.  Reassert these
        # local values after the shared ListBox style is refreshed so a theme
        # reload cannot bring back an inset frame.
        self._popup_list.Background = Brushes.Transparent
        self._popup_list.BorderBrush = Brushes.Transparent
        self._popup_list.BorderThickness = Thickness(0)
        _set_dynamic_brush(
            self._popup_border,
            self,
            "Background",
            "CED.Brush.ListBackground",
        )
        _set_dynamic_brush(
            self._popup_border,
            self,
            "BorderBrush",
            "CED.Brush.ListBorder",
        )
        if hasattr(self, "_popup_footer"):
            _set_dynamic_brush(
                self._popup_footer,
                self,
                "Background",
                "CED.Brush.ListBackground",
            )
            _set_dynamic_brush(
                self._popup_footer,
                self,
                "BorderBrush",
                "CED.Brush.ListBorder",
            )
            _set_dynamic_brush(
                self._popup_footer_text,
                self,
                "Foreground",
                "CED.Brush.SecondaryText",
            )
        _set_dynamic_brush(self._search_icon, self, "Fill", "CED.Brush.IconMuted")
        resource_owner = getattr(self, "_resource_owner", None) or self
        close_geometry = resource_loader.try_find_resource(resource_owner, "CED.Icon.Close")
        search_geometry = resource_loader.try_find_resource(resource_owner, "CED.Icon.Search")
        if close_geometry is not None:
            self._clear_icon.Data = close_geometry
        else:
            try:
                self._clear_icon.SetResourceReference(Path.DataProperty, "CED.Icon.Close")
            except Exception:
                pass
        if search_geometry is not None:
            self._search_icon.Data = search_geometry
        else:
            try:
                self._search_icon.SetResourceReference(Path.DataProperty, "CED.Icon.Search")
            except Exception:
                pass
        _set_dynamic_brush(self._clear_icon, self, "Fill", "CED.Brush.IconMuted")

    def _state_query_changed(self, sender, args):
        for handler in list(self._query_handlers):
            handler(self, args)
        self._update_chrome()

    def _state_command_changed(self, sender, args):
        self._suggestion_index = 0
        self._sync_input_text()
        self._update_command_popup()
        self._update_chrome()
        for handler in list(self._command_handlers):
            handler(self, args)

    def _state_interaction_changed(self, sender, args):
        reason = str(getattr(args, "reason", "") or "")
        if reason in (
            "filter_selected",
            "filter_reactivated",
            "token_selected",
            "token_deselected",
            "token_editing_started",
            "token_editing_finished",
            "filter_removed",
            "cleared",
            "filter_definitions_changed",
        ):
            self._render_tokens()
            if reason in ("filter_selected", "filter_reactivated", "token_editing_started"):
                self._focus_active_editor_deferred()
        self._update_chrome()

    def _set_input_text(self, value):
        text = str(value or "")
        self._suppress_input_events = True
        try:
            if self._input.Text != text:
                self._input.Text = text
            self._input.CaretIndex = len(text)
        finally:
            self._suppress_input_events = False

    def _sync_input_text(self):
        if self._state.is_command_mode:
            self._set_input_text(self._state.command_input_text)
        else:
            self._set_input_text(self._state.free_text)

    def _input_text_changed(self, sender, args):
        if self._suppress_input_events:
            return
        self._state.set_input_text(sender.Text or "")
        self._sync_input_text()
        self._update_command_popup()
        self._update_chrome()

    def _input_focus_changed(self, sender, args):
        self._update_chrome()

    def _search_surface_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if _find_visual_ancestor(source, Button) is not None:
            return
        if _find_visual_ancestor(source, TextBox) is not None:
            return
        self.focus_input()
        args.Handled = True

    def _search_surface_preview_key_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if _find_visual_ancestor(source, TextBox) is self._input:
            self._input_preview_key_down(self._input, args)

    def _input_preview_key_down(self, sender, args):
        key = getattr(args, "Key", None)
        if key == Key.Escape:
            if self._state.is_command_mode:
                self._state.cancel_command_as_literal()
                self._sync_input_text()
                self.focus_input()
                args.Handled = True
                return
            if self._state.active_token_index is not None:
                self._state.commit_active_token()
                self._render_tokens()
                self._focus_input_deferred()
                args.Handled = True
                return
            if self._state.selected_token_index is not None:
                self._state.clear_token_selection()
                args.Handled = True
                return

        if self._state.is_command_mode:
            if key == Key.Down:
                self._move_suggestion(1)
                args.Handled = True
                return
            if key == Key.Up:
                self._move_suggestion(-1)
                args.Handled = True
                return
            if key == Key.Enter:
                if self._select_current_suggestion():
                    args.Handled = True
                return

        if key in (Key.Left, Key.Right):
            if self._handle_input_arrow(key):
                args.Handled = True
                return

        if key == Key.Enter:
            selected = self._state.selected_token_index
            if selected is not None and self._state.edit_token(selected):
                args.Handled = True
                return
            if self._state.active_token_index is not None:
                self._state.commit_active_token()
                self._render_tokens()
                self._focus_input_deferred()
                args.Handled = True
                return

        if key in (Key.Back, Key.Delete):
            if self._state.selected_token_index is not None:
                if key == Key.Back:
                    changed = self._state.backspace_at_input_start()
                else:
                    changed = self._state.delete_at_input_start()
                if changed:
                    args.Handled = True
                    self._render_tokens()
                    self._focus_input_deferred(0)
                return
            try:
                at_input_edge = self._input.SelectionLength == 0 and (
                    self._input.CaretIndex == 0 or not (self._input.Text or "")
                )
            except Exception:
                at_input_edge = not bool(self._input.Text or "")
            if at_input_edge:
                if key == Key.Back:
                    changed = self._state.backspace_at_input_start()
                else:
                    changed = self._state.delete_at_input_start()
                if changed:
                    args.Handled = True
                    self._render_tokens()
                    self._focus_input_deferred(0)

    def _input_boundary(self, key):
        try:
            if self._input.SelectionLength != 0:
                return False
            text = self._input.Text or ""
            caret = int(self._input.CaretIndex)
        except Exception:
            return False
        if key == Key.Left:
            return caret <= 0
        if key == Key.Right:
            return caret >= len(text)
        return False

    def _handle_input_arrow(self, key):
        selected = self._state.selected_token_index
        token_count = len(self._state.tokens)
        if selected is not None:
            if key == Key.Left:
                if selected > 0:
                    self._state.select_token(selected - 1)
                else:
                    self._state.clear_token_selection()
                self._focus_input_deferred(0)
                return True
            if key == Key.Right:
                if selected < token_count - 1:
                    self._state.select_token(selected + 1)
                else:
                    self._state.clear_token_selection()
                self._focus_input_deferred(0)
                return True

        if not self._input_boundary(key):
            return False
        if key == Key.Left and token_count:
            self._state.select_token(token_count - 1)
            self._focus_input_deferred(0)
            return True
        return False

    def _build_token(self, index, token):
        is_active = index == self._state.active_token_index
        is_selected = index == self._state.selected_token_index
        definition = token.definition
        root = Border()
        root.Tag = index
        _set_style(
            root,
            self,
            definition.selected_token_style_key
            if (is_active or is_selected)
            else definition.token_style_key,
        )
        root.MouseEnter += self._token_mouse_enter
        root.MouseLeave += self._token_mouse_leave
        root.PreviewMouseDown += self._token_preview_mouse_down
        root.MouseUp += self._token_mouse_up

        content = Grid()
        root.Child = content
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        # Keep the chip's existing visual padding while letting the hover
        # overlay cover the entire chip interior, including that padding.
        root.Padding = Thickness(0)
        row.Margin = Thickness(4, 0, 4, 0)
        content.Children.Add(row)

        if is_active:
            label = TextBlock()
            label.Text = "{}:".format(token.display_name)
            label.Margin = Thickness(0, 0, 3, 0)
            _set_style(label, self, "CED.SearchBox.TokenText")
            _set_dynamic_brush(
                label,
                self,
                "Foreground",
                definition.token_text_brush_key,
            )
            row.Children.Add(label)

            value_box = TextBox()
            value_box.Text = token.value
            value_box.Tag = index
            value_box.ToolTip = token.definition.placeholder
            value_box.MinWidth = 18.0
            value_box.Width = self._token_editor_width(token.value)
            value_box.TextChanged += self._token_value_changed
            value_box.PreviewKeyDown += self._token_preview_key_down
            value_box.LostFocus += self._token_editor_lost_focus
            _set_style(value_box, self, "CED.SearchBox.TokenInput")
            _set_dynamic_brush(
                value_box,
                self,
                "Foreground",
                definition.token_text_brush_key,
            )
            _set_dynamic_brush(
                value_box,
                self,
                "CaretBrush",
                definition.token_text_brush_key,
            )
            row.Children.Add(value_box)
            self._token_editor = value_box
        else:
            token_button = Button()
            token_button.Tag = index
            token_button.ToolTip = "Press Enter to edit {}".format(token.display_name)
            token_button.Click += self._token_clicked
            token_button.MouseDoubleClick += self._token_double_clicked
            display = TextBlock()
            display.Text = "{}: {}".format(token.display_name, token.display_value)
            _set_style(display, self, "CED.SearchBox.TokenText")
            _set_dynamic_brush(
                display,
                self,
                "Foreground",
                definition.token_text_brush_key,
            )
            token_button.Content = display
            _set_style(token_button, self, "CED.SearchBox.TokenButton")
            _set_dynamic_brush(
                token_button,
                self,
                "Foreground",
                definition.token_text_brush_key,
            )
            row.Children.Add(token_button)

        remove_button = Button()
        remove_button.Tag = index
        remove_button.Content = "×"
        remove_button.ToolTip = "Remove {} filter".format(token.display_name)
        remove_button.Click += self._token_remove_clicked
        _set_style(remove_button, self, "CED.SearchBox.TokenButton")
        _set_dynamic_brush(
            remove_button,
            self,
            "Foreground",
            definition.token_text_brush_key,
        )
        remove_button.Margin = Thickness(4, 0, 0, 0)
        row.Children.Add(remove_button)

        overlay = Border()
        overlay.IsHitTestVisible = False
        overlay.Visibility = Visibility.Visible if is_selected else Visibility.Collapsed
        try:
            overlay.CornerRadius = root.CornerRadius
        except Exception:
            pass
        if is_selected:
            _set_dynamic_brush(
                overlay,
                self,
                "Background",
                "CED.Brush.ButtonStateOverlayPressed",
            )
        content.Children.Add(overlay)
        return root

    @staticmethod
    def _token_overlay(sender):
        try:
            return sender.Child.Children[1]
        except Exception:
            return None

    def _show_token_overlay(self, sender, resource_key):
        overlay = self._token_overlay(sender)
        if overlay is None:
            return
        _set_dynamic_brush(overlay, self, "Background", resource_key)
        overlay.Visibility = Visibility.Visible

    def _hide_token_overlay(self, sender, args=None):
        overlay = self._token_overlay(sender)
        if overlay is not None:
            try:
                index = int(sender.Tag)
            except Exception:
                index = None
            if index is not None and index == self._state.selected_token_index:
                overlay.Visibility = Visibility.Visible
            else:
                overlay.Visibility = Visibility.Collapsed

    def _token_mouse_enter(self, sender, args):
        self._show_token_overlay(sender, "CED.Brush.ButtonStateOverlayHover")

    def _token_mouse_leave(self, sender, args):
        self._hide_token_overlay(sender, args)

    def _token_preview_mouse_down(self, sender, args):
        self._show_token_overlay(sender, "CED.Brush.ButtonStateOverlayPressed")

    def _token_mouse_up(self, sender, args):
        self._show_token_overlay(sender, "CED.Brush.ButtonStateOverlayHover")

    def _token_double_clicked(self, sender, args):
        try:
            index = int(sender.Tag)
        except Exception:
            return
        if self._state.edit_token(index):
            self._focus_active_editor_deferred()
            args.Handled = True

    def _token_editor_width(self, value):
        text = str(value or "")
        return max(18.0, min(220.0, 8.0 + (len(text) * 7.0)))

    def _render_tokens(self):
        if not hasattr(self, "_tokens_panel"):
            return
        self._token_editor = None
        self._tokens_panel.Children.Clear()
        previous_token = None
        for index, token in enumerate(self._state.tokens):
            if (
                previous_token is not None
                and previous_token.key.lower() == token.key.lower()
                and token.definition.combine_mode == "or"
            ):
                operator = TextBlock()
                operator.Text = "OR"
                operator.Width = 18.0
                operator.TextAlignment = TextAlignment.Center
                operator.Margin = Thickness(-2, 2, 2, 2)
                _set_style(operator, self, "CED.SearchBox.TokenOperator")
                self._tokens_panel.Children.Add(operator)
            self._tokens_panel.Children.Add(self._build_token(index, token))
            previous_token = token
        self._tokens_panel.Children.Add(self._input)
        self._sync_input_text()

    def _focus_active_editor(self):
        if self._token_editor is None:
            return
        try:
            self._token_editor.Focus()
            try:
                Keyboard.Focus(self._token_editor)
            except Exception:
                pass
            self._token_editor.CaretIndex = len(self._token_editor.Text or "")
        except Exception:
            pass

    def _token_value_changed(self, sender, args):
        if self._token_editor is not sender:
            return
        try:
            sender.Width = self._token_editor_width(sender.Text)
            index = int(sender.Tag)
        except Exception:
            return
        self._state.set_active_token_value(sender.Text or "")
        if index != self._state.active_token_index:
            return
        self._update_chrome()

    def _token_editor_lost_focus(self, sender, args):
        if self._state.active_token_index is None:
            return
        self._state.commit_active_token()

    def _token_preview_key_down(self, sender, args):
        key = getattr(args, "Key", None)
        if key in (Key.Left, Key.Right):
            if self._handle_token_editor_arrow(sender, key):
                args.Handled = True
                return
        if key in (Key.Back, Key.Delete):
            try:
                empty_at_edge = not (sender.Text or "") and sender.CaretIndex == 0
                index = int(sender.Tag)
            except Exception:
                empty_at_edge = False
                index = None
            if empty_at_edge and index is not None:
                self._state.select_token(index)
                self._render_tokens()
                self._focus_input_deferred()
                args.Handled = True
                return
        if key == Key.Escape:
            self._state.commit_active_token()
            self._render_tokens()
            self._focus_input_deferred()
            args.Handled = True
        elif key == Key.Enter:
            self._state.commit_active_token()
            self._render_tokens()
            self._focus_input_deferred()
            args.Handled = True

    def _handle_token_editor_arrow(self, sender, key):
        try:
            if sender.SelectionLength != 0:
                return False
            caret = int(sender.CaretIndex)
            text_length = len(sender.Text or "")
            index = int(sender.Tag)
        except Exception:
            return False

        token_count = len(self._state.tokens)
        if key == Key.Left and caret <= 0:
            if index <= 0:
                return False
            self._state.commit_active_token()
            self._state.select_token(index - 1)
            self._focus_input_deferred(0)
            return True
        if key == Key.Right and caret >= text_length:
            self._state.commit_active_token()
            if index < token_count - 1:
                self._state.select_token(index + 1)
                self._focus_input_deferred(0)
            else:
                self._focus_input_deferred(0)
            return True
        return False

    def _token_clicked(self, sender, args):
        try:
            index = int(sender.Tag)
        except Exception:
            return
        self._state.select_token(index)
        self._focus_input_deferred(0)
        args.Handled = True

    def _token_remove_clicked(self, sender, args):
        try:
            index = int(sender.Tag)
        except Exception:
            return
        self._state.remove_token(index)
        self._focus_input_deferred(0)
        args.Handled = True

    def _clear_clicked(self, sender, args):
        self.clear()
        self._focus_input_deferred()
        args.Handled = True

    def _update_chrome(self):
        has_input = bool(self._state.free_text or self._state.command_text)
        has_content = bool(has_input or self._state.tokens)
        try:
            focused = bool(self._input.IsKeyboardFocused or self.IsKeyboardFocusWithin)
        except Exception:
            focused = False
        if not has_content and not self._state.is_command_mode:
            self._placeholder.Text = (
                self._active_placeholder_text if focused else self._placeholder_text
            )
            self._placeholder.Visibility = Visibility.Visible
        else:
            self._placeholder.Visibility = Visibility.Collapsed
        self._clear_button.Visibility = Visibility.Visible if has_content else Visibility.Collapsed

    def _update_command_popup(self):
        if not self._state.is_command_mode:
            self._popup.IsOpen = False
            return
        suggestions = self._state.command_suggestions
        self._popup_list.Items.Clear()
        for definition in suggestions:
            item = ListBoxItem()
            item.Tag = definition
            item.Padding = Thickness(6, 1, 6, 1)
            item.MinHeight = 0.0
            _set_style(item, self, "CED.SearchBox.CommandItem")
            _set_dynamic_brush(
                item,
                self,
                "Background",
                "CED.Brush.ListItemBackground",
            )
            _set_dynamic_brush(
                item,
                self,
                "Foreground",
                "CED.Brush.PrimaryText",
            )
            row = Grid()
            row.MinHeight = 18.0
            name_column = ColumnDefinition()
            name_column.Width = GridLength(82.0, GridUnitType.Pixel)
            command_column = ColumnDefinition()
            command_column.Width = GridLength(58.0, GridUnitType.Pixel)
            hint_column = ColumnDefinition()
            hint_column.Width = GridLength(1.0, GridUnitType.Star)
            row.ColumnDefinitions.Add(name_column)
            row.ColumnDefinitions.Add(command_column)
            row.ColumnDefinitions.Add(hint_column)
            name = TextBlock()
            name.Text = definition.display_name
            name.VerticalAlignment = VerticalAlignment.Center
            _set_style(name, self, "CED.SearchBox.TokenText")
            _set_dynamic_brush(
                name,
                self,
                "Foreground",
                definition.token_text_brush_key,
            )
            Grid.SetColumn(name, 0)
            row.Children.Add(name)
            key_text = TextBlock()
            key_text.Text = "/{}".format(definition.key)
            key_text.VerticalAlignment = VerticalAlignment.Center
            _set_dynamic_brush(key_text, self, "Foreground", "CED.Brush.SecondaryText")
            Grid.SetColumn(key_text, 1)
            row.Children.Add(key_text)
            if definition.value_hint:
                hint = TextBlock()
                hint.Text = definition.value_hint
                hint.FontSize = 10.0
                hint.VerticalAlignment = VerticalAlignment.Center
                _set_dynamic_brush(hint, self, "Foreground", "CED.Brush.SecondaryText")
                Grid.SetColumn(hint, 2)
                row.Children.Add(hint)
            item.Content = row
            self._popup_list.Items.Add(item)
        if suggestions:
            self._suggestion_index = max(0, min(self._suggestion_index, len(suggestions) - 1))
            self._popup_list.SelectedIndex = self._suggestion_index
            self._popup.IsOpen = True
        else:
            self._popup.IsOpen = False

    def _move_suggestion(self, delta):
        count = self._popup_list.Items.Count
        if not count:
            return
        self._suggestion_index = (self._suggestion_index + int(delta)) % count
        self._popup_list.SelectedIndex = self._suggestion_index

    def _selected_suggestion(self):
        try:
            item = self._popup_list.SelectedItem
            definition = getattr(item, "Tag", None)
            if isinstance(definition, SearchFilterDefinition):
                return definition
        except Exception:
            pass
        suggestions = self._state.command_suggestions
        if suggestions and 0 <= self._suggestion_index < len(suggestions):
            return suggestions[self._suggestion_index]
        return None

    def _select_current_suggestion(self):
        definition = self._selected_suggestion()
        if definition is None:
            return False
        token = self._state.select_filter(definition)
        if token is None:
            return False
        return True

    def _suggestion_clicked(self, sender, args):
        if self._select_current_suggestion():
            args.Handled = True

    def _popup_preview_key_down(self, sender, args):
        key = getattr(args, "Key", None)
        if key == Key.Enter:
            if self._select_current_suggestion():
                args.Handled = True
        elif key == Key.Escape:
            self._state.cancel_command_as_literal()
            self.focus_input()
            args.Handled = True

    def _popup_closed(self, sender, args):
        # Popup closure can happen before the suggestion click has finished
        # routing.  Defer the literal fallback so selecting a field wins over
        # focus loss, while clicking outside still commits the slash as text.
        def _cancel_if_still_provisional():
            if self._state.is_command_mode:
                self._state.cancel_command_as_literal()

        self._invoke_later(_cancel_if_still_provisional)


__all__ = ["StructuredSearchBox"]
