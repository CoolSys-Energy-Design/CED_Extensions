# -*- coding: utf-8 -*-
"""Reusable editable/filterable WPF ComboBox behavior for CED tools."""

try:
    from System import Action
    from System.Windows import DataTemplate, DependencyProperty, FrameworkElementFactory
    from System.Windows.Controls import ComboBox, ComboBoxItem, ScrollViewer, TextBlock, TextBox
    from System.Windows.Data import Binding, CollectionViewSource
    from System.Windows.Documents import Run
    from System.Windows.Input import (
        FocusNavigationDirection,
        Key,
        Keyboard,
        TraversalRequest,
    )
    from System.Windows.Media import Brushes, VisualTreeHelper
    from System.Windows.Threading import DispatcherPriority
except Exception:
    # Keep the matching and state logic importable by CPython tests.  WPF is
    # only available when the behavior is constructed inside pyRevit.
    Action = None
    Binding = None
    Brushes = None
    ComboBox = None
    ComboBoxItem = None
    ScrollViewer = None
    CollectionViewSource = None
    DataTemplate = None
    DependencyProperty = None
    DispatcherPriority = None
    FocusNavigationDirection = None
    FrameworkElementFactory = None
    Key = None
    Keyboard = None
    Run = None
    TextBlock = None
    TextBox = None
    TraversalRequest = None
    VisualTreeHelper = None


FALLBACK_BLANK = "blank"
FALLBACK_LAST_VALID = "last_valid"
FALLBACK_DEFAULT_ITEM = "default_item"

try:
    _STRING_TYPES = (basestring,)
except NameError:
    _STRING_TYPES = (str,)


def _text_value(value):
    """Return text without changing user-entered whitespace."""
    if value is None:
        return ""
    try:
        return str(value)
    except Exception:
        return ""


def _safe_text(value):
    """Return normalized text for matching or display comparisons."""
    return _text_value(value).strip()


def _normalized_query(value):
    return _safe_text(value).lower()


def item_text(item, text_getter=None, display_member_path=None):
    """Return the text users should search for and see in an item."""
    if ComboBoxItem is not None and isinstance(item, ComboBoxItem):
        item = getattr(item, "Content", item)
    if text_getter is not None:
        try:
            return _safe_text(text_getter(item))
        except Exception:
            return ""

    path = _safe_text(display_member_path)
    if path:
        value = item
        for part in path.split("."):
            try:
                value = getattr(value, part)
            except Exception:
                value = None
                break
        if value is not None:
            return _safe_text(value)
    return _safe_text(item)


def filter_items(items, query, text_getter=None, display_member_path=None):
    """Return items whose display text contains *query*, case-insensitively."""
    values = list(items) if items is not None else []
    needle = _normalized_query(query)
    if not needle:
        return values
    return [
        value
        for value in values
        if needle
        in item_text(
            value,
            text_getter=text_getter,
            display_member_path=display_member_path,
        ).lower()
    ]


def highlight_segments(value, query):
    """Split display text into ``(text, is_match)`` segments."""
    text = _text_value(value)
    needle = _safe_text(query)
    if not text:
        return []
    if not needle:
        return [(text, False)]

    lower_text = text.lower()
    lower_needle = needle.lower()
    segments = []
    cursor = 0
    while cursor < len(text):
        match_start = lower_text.find(lower_needle, cursor)
        if match_start < 0:
            if cursor < len(text):
                segments.append((text[cursor:], False))
            break
        if match_start > cursor:
            segments.append((text[cursor:match_start], False))
        match_end = match_start + len(needle)
        segments.append((text[match_start:match_end], True))
        cursor = match_end
    return segments


def navigation_index(current_index, item_count, direction):
    """Return a clamped visible-item index for an Up/Down key press."""
    count = int(item_count or 0)
    if count <= 0:
        return -1
    current = int(current_index)
    step = 1 if int(direction or 0) > 0 else -1
    if current < 0:
        return 0 if step > 0 else count - 1
    return max(0, min(count - 1, current + step))


def _same_item(left, right):
    if left is right:
        return True
    try:
        return bool(left == right)
    except Exception:
        return False


def find_exact_item(items, value, text_getter=None, display_member_path=None):
    """Find a list item whose display text exactly matches *value*."""
    needle = _normalized_query(value)
    if not needle:
        return None
    for candidate in list(items) if items is not None else []:
        candidate_text = item_text(
            candidate,
            text_getter=text_getter,
            display_member_path=display_member_path,
        )
        if candidate_text.lower() == needle:
            return candidate
    return None


def _find_reference_item(
    items,
    reference,
    text_getter=None,
    display_member_path=None,
):
    if reference is None:
        return None
    values = list(items) if items is not None else []
    for candidate in values:
        if _same_item(candidate, reference):
            return candidate
    if isinstance(reference, _STRING_TYPES):
        return find_exact_item(
            values,
            reference,
            text_getter=text_getter,
            display_member_path=display_member_path,
        )
    return None


def resolve_fallback_item(
    items,
    fallback=FALLBACK_LAST_VALID,
    last_valid_item=None,
    default_item=None,
    text_getter=None,
    display_member_path=None,
):
    """Resolve an enforced combo's configured fallback item."""
    policy = _safe_text(fallback).lower() or FALLBACK_LAST_VALID
    values = list(items) if items is not None else []
    if policy == FALLBACK_LAST_VALID:
        return _find_reference_item(
            values,
            last_valid_item,
            text_getter=text_getter,
            display_member_path=display_member_path,
        )
    if policy in (FALLBACK_DEFAULT_ITEM, "default"):
        return _find_reference_item(
            values,
            default_item,
            text_getter=text_getter,
            display_member_path=display_member_path,
        )
    return None


class FilterableComboBox(object):
    """Add live filtering and predictable editing to one WPF ComboBox.

    The behavior keeps three values separate: the filter query, the current
    candidate row, and the committed item.  Enforced mode displays the row
    being navigated and commits it on Enter, a row click, or blur resolution.
    Custom mode leaves typed text visible until a row is intentionally chosen.
    """

    def __init__(
        self,
        combo,
        text_getter=None,
        on_filter_changed=None,
        allow_custom_values=True,
        fallback=FALLBACK_LAST_VALID,
        fallback_item=None,
        on_value_committed=None,
        navigation_requires_open=False,
        clear_focus_on_commit=True,
        commit_selection_while_open=False,
        commit_programmatic_selection=True,
        navigation_advances_implicit_candidate=False,
    ):
        if ComboBox is not None and not isinstance(combo, ComboBox):
            raise TypeError("FilterableComboBox requires a WPF ComboBox")
        if combo is None:
            raise ValueError("FilterableComboBox requires a ComboBox")

        self.combo = combo
        self._text_getter = text_getter
        self._display_member_path = getattr(combo, "DisplayMemberPath", "")
        self._on_filter_changed_callback = on_filter_changed
        self._on_value_committed_callback = on_value_committed
        self.allow_custom_values = bool(allow_custom_values)
        self.fallback = fallback or FALLBACK_LAST_VALID
        self.fallback_item = fallback_item
        self.navigation_requires_open = bool(navigation_requires_open)
        self.clear_focus_on_commit = bool(clear_focus_on_commit)
        self.commit_selection_while_open = bool(commit_selection_while_open)
        self.commit_programmatic_selection = bool(commit_programmatic_selection)
        self.navigation_advances_implicit_candidate = bool(
            navigation_advances_implicit_candidate
        )

        self._textbox = None
        self._query_text = ""
        self._candidate_item = None
        self._candidate_index = -1
        # A candidate can be highlighted without having been intentionally
        # navigated to.  Enforced mode uses that distinction so opening or
        # filtering highlights row zero, while the first Up/Down still shows
        # row zero instead of skipping it.
        self._candidate_engaged = False
        self._candidate_backgrounds = {}
        self._committed_item = getattr(combo, "SelectedItem", None)
        self._last_valid_item = self._committed_item
        self._editing = False
        self._suppress_events = 0
        self._popup = None
        self._direct_items_snapshot = []
        self._attached = False
        self._highlight_refresh_pending = False

        self._capture_items_snapshot()
        initial_text = self._read_text()
        if self._committed_item is None:
            self._query_text = initial_text
            self._editing = bool(initial_text)

        self._attach_combo_events()
        self._install_item_template()
        self._attach_textbox()

    @property
    def query(self):
        return self._query_text

    @property
    def value(self):
        if (
            self._candidate_item is not None
            and self._candidate_engaged
            and not self.allow_custom_values
        ):
            return item_text(
                self._candidate_item,
                text_getter=self._text_getter,
                display_member_path=self._display_member_path,
            )
        if self._committed_item is not None:
            return item_text(
                self._committed_item,
                text_getter=self._text_getter,
                display_member_path=self._display_member_path,
            )
        return self._read_text()

    @property
    def selected_item(self):
        return self._committed_item

    @property
    def has_valid_selection(self):
        return self._committed_item is not None

    def dispose(self):
        if not self._attached:
            return
        try:
            self.combo.Loaded -= self._on_loaded
            self.combo.DropDownOpened -= self._on_dropdown_opened
            self.combo.DropDownClosed -= self._on_dropdown_closed
            self.combo.SelectionChanged -= self._on_selection_changed
            self.combo.GotKeyboardFocus -= self._on_got_keyboard_focus
            self.combo.LostKeyboardFocus -= self._on_combo_lost_focus
            self.combo.PreviewKeyDown -= self._on_key_down
            self.combo.PreviewMouseLeftButtonDown -= self._on_combo_mouse_down
            self.combo.PreviewMouseWheel -= self._on_combo_mouse_wheel
        except Exception:
            pass
        self._detach_textbox()
        self._detach_popup()
        self._clear_filter(notify=False)
        self._attached = False

    def _attach_combo_events(self):
        try:
            self.combo.Loaded += self._on_loaded
            self.combo.DropDownOpened += self._on_dropdown_opened
            self.combo.DropDownClosed += self._on_dropdown_closed
            self.combo.SelectionChanged += self._on_selection_changed
            self.combo.GotKeyboardFocus += self._on_got_keyboard_focus
            self.combo.LostKeyboardFocus += self._on_combo_lost_focus
            # Catch keys from the toggle; textbox and popup row handlers catch
            # keys after focus enters those surfaces.
            self.combo.PreviewKeyDown += self._on_key_down
            self.combo.PreviewMouseLeftButtonDown += self._on_combo_mouse_down
            self.combo.PreviewMouseWheel += self._on_combo_mouse_wheel
            self._attached = True
        except Exception:
            self._attached = False

    def _on_loaded(self, sender, args):
        self._install_item_template()
        self._attach_textbox()
        self._attach_popup()
        # Recycled grid editors can load many times while the user scrolls.
        # With no active query there are no highlights to rebuild, so avoid a
        # full item-container walk on every virtualization cycle.
        if self._query_text or bool(getattr(self.combo, "IsDropDownOpen", False)):
            self._queue_highlight_refresh()

    def _on_got_keyboard_focus(self, sender, args):
        self._attach_textbox()

    def _on_dropdown_opened(self, sender, args):
        self._install_item_template()
        self._attach_textbox()
        self._attach_popup()
        # TextChanged filters before it opens the popup. Do not refresh the
        # same view a second time during that synchronous open event.
        if not self._editing:
            self._apply_filter("", False)
        if self.allow_custom_values:
            self._clear_candidate()
        else:
            # First row is the enforced default candidate, but does not
            # replace an empty/user query until navigation or Enter.
            self._set_candidate(0, replace_display=False, engaged=False)
        # The arrow can leave focus on the ToggleButton.  Return it to the
        # editable part so keyboard handling is deterministic.
        self._focus_textbox()

    def _on_dropdown_closed(self, sender, args):
        if not self._editing and self._candidate_item is not None:
            self._clear_candidate()
        # A Popup has its own focus scope.  Losing focus to a popup row is not
        # an external blur; resolve only after the popup has actually closed.
        self._request_blur_resolution()

    def _on_selection_changed(self, sender, args):
        if self._suppress_events:
            return
        selected = getattr(self.combo, "SelectedItem", None)
        if selected is None:
            if not self.commit_programmatic_selection and not self._editing:
                self._adopt_selected_item(None)
            return
        if (
            self._committed_item is not None
            and _same_item(selected, self._committed_item)
            and not self._editing
        ):
            return

        if bool(getattr(self.combo, "IsDropDownOpen", False)):
            if self.commit_selection_while_open:
                self._commit_item(selected)
                return
            index = self._visible_index(selected)
            if index >= 0:
                self._set_candidate(
                    index,
                    replace_display=not self.allow_custom_values,
                    engaged=True,
                )
                return
        if not self.commit_programmatic_selection and not self._editing:
            self._adopt_selected_item(selected)
            return
        self._commit_item(selected)

    def _adopt_selected_item(self, item):
        """Synchronize a binding-driven selection without reporting a commit."""
        self._clear_candidate()
        self._committed_item = item
        if item is not None:
            self._last_valid_item = item
        self._editing = False
        self._query_text = ""
        value = item_text(
            item,
            text_getter=self._text_getter,
            display_member_path=self._display_member_path,
        ) if item is not None else ""
        self._suppress_events += 1
        try:
            self._write_text(value)
            self._set_caret_end()
        finally:
            self._suppress_events -= 1

    def _editable_textbox(self, combo):
        if combo is None:
            return None
        try:
            combo.ApplyTemplate()
        except Exception:
            pass
        try:
            template = getattr(combo, "Template", None)
            if template is not None:
                textbox = template.FindName("PART_EditableTextBox", combo)
                if textbox is not None:
                    return textbox
        except Exception:
            pass
        if VisualTreeHelper is None or TextBox is None:
            return None
        queue = [combo]
        while queue:
            current = queue.pop(0)
            try:
                count = int(VisualTreeHelper.GetChildrenCount(current) or 0)
            except Exception:
                count = 0
            for index in range(count):
                try:
                    child = VisualTreeHelper.GetChild(current, index)
                except Exception:
                    continue
                if isinstance(child, TextBox):
                    return child
                queue.append(child)
        return None

    def _install_item_template(self):
        try:
            if getattr(self.combo, "ItemTemplate", None) is not None:
                return
        except Exception:
            pass
        if self._text_getter is None and not _safe_text(self._display_member_path):
            try:
                template = self.combo.TryFindResource(
                    "CED.Input.ComboBox.FilterableItemTemplate"
                )
                if template is not None:
                    self.combo.ItemTemplate = template
                    return
            except Exception:
                pass
        if DataTemplate is None or FrameworkElementFactory is None or Binding is None:
            return
        try:
            factory = FrameworkElementFactory(TextBlock)
            factory.SetBinding(
                TextBlock.TextProperty,
                Binding(_safe_text(self._display_member_path) or "."),
            )
            factory.SetResourceReference(
                TextBlock.StyleProperty,
                "CED.Input.ComboBox.FilterableItemText",
            )
            template = DataTemplate()
            template.VisualTree = factory
            self.combo.ItemTemplate = template
        except Exception:
            pass

    def _attach_textbox(self):
        textbox = self._editable_textbox(self.combo)
        if textbox is None or textbox is self._textbox:
            return
        self._detach_textbox()
        try:
            textbox.TextChanged += self._on_text_changed
            textbox.GotKeyboardFocus += self._on_textbox_got_focus
            textbox.LostKeyboardFocus += self._on_textbox_lost_focus
            textbox.PreviewKeyDown += self._on_key_down
            textbox.PreviewMouseDoubleClick += self._on_textbox_double_click
            self._textbox = textbox
        except Exception:
            self._textbox = None

    def _editable_popup(self):
        try:
            self.combo.ApplyTemplate()
            template = getattr(self.combo, "Template", None)
            if template is not None:
                return template.FindName("PART_Popup", self.combo)
        except Exception:
            pass
        return None

    def _attach_popup(self):
        """Attach at the Popup root so routed row input cannot escape us.

        Popup content does not route keyboard events through the parent
        ComboBox.  Attaching once at PART_Popup is both more reliable and
        simpler than chasing generated/recycled ComboBoxItem containers.
        """
        popup = self._editable_popup()
        if popup is None or popup is self._popup:
            return
        self._detach_popup()
        try:
            popup.PreviewKeyDown += self._on_key_down
            popup.PreviewMouseLeftButtonDown += self._on_popup_mouse_down
            popup.PreviewMouseWheel += self._on_popup_mouse_wheel
            self._popup = popup
        except Exception:
            self._popup = None

    def _detach_popup(self):
        if self._popup is None:
            return
        try:
            self._popup.PreviewKeyDown -= self._on_key_down
            self._popup.PreviewMouseLeftButtonDown -= self._on_popup_mouse_down
            self._popup.PreviewMouseWheel -= self._on_popup_mouse_wheel
        except Exception:
            pass
        self._popup = None

    def _detach_textbox(self):
        if self._textbox is None:
            return
        try:
            self._textbox.TextChanged -= self._on_text_changed
            self._textbox.GotKeyboardFocus -= self._on_textbox_got_focus
            self._textbox.LostKeyboardFocus -= self._on_textbox_lost_focus
            self._textbox.PreviewKeyDown -= self._on_key_down
            self._textbox.PreviewMouseDoubleClick -= self._on_textbox_double_click
        except Exception:
            pass
        self._textbox = None

    def _read_text(self):
        if self._textbox is not None:
            return _text_value(getattr(self._textbox, "Text", ""))
        return _text_value(getattr(self.combo, "Text", ""))

    def _write_text(self, value):
        text = _text_value(value)
        try:
            self.combo.Text = text
        except Exception:
            pass
        if self._textbox is not None:
            try:
                self._textbox.Text = text
            except Exception:
                pass

    def _set_caret_end(self):
        if self._textbox is None:
            return
        text_length = len(self._read_text())
        try:
            self._textbox.Select(text_length, 0)
            self._textbox.CaretIndex = text_length
        except Exception:
            pass

    def _select_all(self):
        if self._textbox is None:
            return
        try:
            self._textbox.SelectAll()
        except Exception:
            pass

    def _focus_textbox(self):
        if self._textbox is None:
            return
        # Typing can open the dropdown synchronously. Refocusing the editor at
        # that point raises GotKeyboardFocus again and can select the first
        # character the user just entered.
        if bool(getattr(self._textbox, "IsKeyboardFocused", False)):
            return
        try:
            if Keyboard is not None:
                Keyboard.Focus(self._textbox)
            else:
                self._textbox.Focus()
        except Exception:
            pass

    def _on_textbox_got_focus(self, sender, args):
        if self._committed_item is not None and not self._editing:
            self._select_all()

    def _on_textbox_double_click(self, sender, args):
        self._select_all()
        try:
            args.Handled = True
        except Exception:
            pass

    def _on_text_changed(self, sender, args):
        if self._suppress_events:
            return
        # DataGrid virtualization and binding refreshes can update the editor
        # while it is off-screen or unfocused. Those are synchronization
        # changes, not user queries, and must never open the Popup.
        if not bool(getattr(self.combo, "IsKeyboardFocusWithin", False)) and not bool(
            getattr(sender, "IsKeyboardFocused", False)
        ):
            selected = getattr(self.combo, "SelectedItem", None)
            self._clear_candidate()
            self._query_text = ""
            self._editing = False
            self._committed_item = selected
            if selected is not None:
                self._last_valid_item = selected
            self._clear_filter(notify=False)
            return
        text = self._read_text()
        self._clear_candidate()
        self._editing = True
        self._committed_item = None
        if getattr(self.combo, "SelectedItem", None) is not None:
            self._set_internal_text(text, clear_selection=True)
        else:
            # Editable WPF templates usually synchronize this automatically,
            # but keeping ComboBox.Text explicit makes the committed value
            # reliable for bindings and custom-entry consumers as well.
            self._suppress_events += 1
            try:
                self._write_text(text)
            finally:
                self._suppress_events -= 1
        self._apply_filter(text, True)
        if not self.allow_custom_values:
            self._set_candidate(0, replace_display=False, engaged=False)

    def _set_internal_text(self, value, clear_selection=False):
        self._suppress_events += 1
        try:
            if clear_selection:
                try:
                    self.combo.SelectedIndex = -1
                except Exception:
                    pass
            self._write_text(value)
        finally:
            self._suppress_events -= 1
        self._set_caret_end()

    def _key_is(self, args, key_name):
        key = getattr(args, "Key", None)
        key_text = str(key)
        if key_text == key_name:
            return True
        if key_name == "Enter" and key_text == "Return":
            return True
        return Key is not None and key == getattr(Key, key_name, None)

    def _on_key_down(self, sender, args):
        if self._suppress_events:
            return
        if self._key_is(args, "Escape"):
            if bool(getattr(self.combo, "IsDropDownOpen", False)):
                if self._commit_keyboard_value():
                    args.Handled = True
                return
            if self._editing:
                self._cancel_editing()
                args.Handled = True
            return
        if self._key_is(args, "Down"):
            if self.navigation_requires_open and not bool(
                getattr(self.combo, "IsDropDownOpen", False)
            ):
                args.Handled = True
                return
            if self._move_candidate(1):
                args.Handled = True
            return
        if self._key_is(args, "Up"):
            if self.navigation_requires_open and not bool(
                getattr(self.combo, "IsDropDownOpen", False)
            ):
                args.Handled = True
                return
            if self._move_candidate(-1):
                args.Handled = True
            return
        if self._key_is(args, "Enter") and self._commit_keyboard_value():
            args.Handled = True

    def handle_host_preview_key_down(self, args):
        """Let an ancestor control give an open picker first refusal on keys.

        Controls such as DataGrid process preview arrow keys before descendant
        editors receive them. Hosts can call this from their own preview route
        without duplicating the picker state machine.
        """
        if not bool(getattr(self.combo, "IsDropDownOpen", False)):
            return False
        self._on_key_down(self.combo, args)
        return bool(getattr(args, "Handled", False))

    def _move_candidate(self, direction):
        query = self._query_text
        previous_item = self._candidate_item
        previous_engaged = self._candidate_engaged
        # The current query was already applied by TextChanged. Re-filtering
        # here made every arrow press rebuild the popup and could invalidate
        # the candidate before navigation completed.
        if not bool(getattr(self.combo, "IsDropDownOpen", False)):
            self._apply_filter(query, True)
        visible = self._visible_items()
        if not visible:
            return False

        current_index = -1
        if previous_item is not None:
            current_index = self._visible_index(previous_item, visible)
        if (
            previous_item is not None
            and not previous_engaged
            and not self.navigation_advances_implicit_candidate
        ):
            target_index = current_index
        else:
            target_index = navigation_index(current_index, len(visible), direction)
        if target_index < 0:
            return False
        self._set_candidate(
            target_index,
            replace_display=not self.allow_custom_values,
            engaged=True,
        )
        return True

    def _commit_keyboard_value(self):
        if self._candidate_item is not None:
            return self._commit_item(self._candidate_item)
        if self.allow_custom_values:
            return self._commit_custom_text()
        return self._resolve_enforced_value()

    def _set_candidate(self, index, replace_display=False, engaged=True):
        self._clear_candidate()
        visible = self._visible_items()
        if index < 0 or index >= len(visible):
            return None
        item = visible[index]
        self._candidate_item = item
        self._candidate_index = index
        self._candidate_engaged = bool(engaged)
        self._set_container_highlight(index, True)
        if replace_display:
            value = item_text(
                item,
                text_getter=self._text_getter,
                display_member_path=self._display_member_path,
            )
            self._committed_item = None
            self._editing = True
            # Let ComboBox render a navigated row through its native
            # selection.  Clearing SelectedItem while forcing editable text
            # makes WPF resynchronize PART_EditableTextBox back to blank on
            # the next layout pass, which destroys both the display and the
            # saved filter query in a live Popup.
            self._suppress_events += 1
            try:
                self.combo.SelectedItem = item
                self._write_text(value)
            finally:
                self._suppress_events -= 1
            self._set_caret_end()
        return item

    def _clear_candidate(self):
        if self._candidate_index >= 0:
            self._set_container_highlight(self._candidate_index, False)
        self._candidate_item = None
        self._candidate_index = -1
        self._candidate_engaged = False

    def _on_textbox_lost_focus(self, sender, args):
        self._request_blur_resolution()

    def _on_combo_lost_focus(self, sender, args):
        self._request_blur_resolution()

    def _request_blur_resolution(self):
        # Exactly one dispatcher pass lets WPF finish moving focus.  If focus
        # moved into the Popup we stop; DropDownClosed requests a fresh check.
        self._begin_invoke(self._resolve_if_focus_outside, priority="ContextIdle")

    def _resolve_if_focus_outside(self):
        try:
            if bool(self.combo.IsKeyboardFocusWithin):
                return
            if bool(self.combo.IsDropDownOpen):
                return
            if self._popup is not None and bool(self._popup.IsKeyboardFocusWithin):
                return
        except Exception:
            pass
        self._resolve_on_blur()

    def _resolve_on_blur(self):
        if self._suppress_events:
            return
        if not self._editing:
            self._clear_filter(notify=True)
            return
        if self.allow_custom_values:
            self._commit_custom_text()
            return
        self._resolve_enforced_value()

    def _cancel_editing(self):
        """Close the Popup and restore the last committed list value."""
        self._clear_candidate()
        self._query_text = ""
        self._editing = False
        restore_item = self._last_valid_item
        view = self._items_view()
        self._suppress_events += 1
        try:
            if view is not None:
                view.Filter = None
                try:
                    view.Refresh()
                except Exception:
                    pass
            self._committed_item = restore_item
            if restore_item is None:
                self.combo.SelectedIndex = -1
                self._write_text("")
            else:
                self.combo.SelectedItem = restore_item
                self._write_text(
                    item_text(
                        restore_item,
                        text_getter=self._text_getter,
                        display_member_path=self._display_member_path,
                    )
                )
            self.combo.IsDropDownOpen = False
        finally:
            self._suppress_events -= 1
        if view is not None:
            self._notify_filter_changed(view)
        self._queue_highlight_refresh()
        self._focus_textbox()

    def _commit_custom_text(self):
        value = self._read_text()
        was_dirty = self._editing
        self._clear_candidate()
        if was_dirty:
            self._committed_item = None
        self._editing = False
        self._query_text = ""
        self._suppress_events += 1
        try:
            self._write_text(value)
            self._set_caret_end()
        finally:
            self._suppress_events -= 1
        self._clear_filter(notify=True)
        self._close_and_unfocus()
        if was_dirty:
            self._notify_value_committed(value, None)
        return True

    def _resolve_enforced_value(self):
        exact = find_exact_item(
            self._all_items(),
            self._read_text(),
            text_getter=self._text_getter,
            display_member_path=self._display_member_path,
        )
        if exact is not None:
            return self._commit_item(exact)
        fallback_item = resolve_fallback_item(
            self._all_items(),
            fallback=self.fallback,
            last_valid_item=self._last_valid_item,
            default_item=self.fallback_item,
            text_getter=self._text_getter,
            display_member_path=self._display_member_path,
        )
        if fallback_item is not None:
            return self._commit_item(fallback_item)
        return self._commit_blank()

    def _commit_item(self, item):
        if item is None:
            return self._commit_blank()
        value = item_text(
            item,
            text_getter=self._text_getter,
            display_member_path=self._display_member_path,
        )
        self._clear_candidate()
        self._committed_item = item
        self._last_valid_item = item
        self._editing = False
        self._query_text = ""

        view = self._items_view()
        self._suppress_events += 1
        try:
            if view is not None:
                view.Filter = None
                try:
                    view.Refresh()
                except Exception:
                    pass
            self.combo.SelectedItem = item
            self._write_text(value)
            self._set_caret_end()
            self.combo.IsDropDownOpen = False
        except Exception:
            return False
        finally:
            self._suppress_events -= 1

        if view is not None:
            self._notify_filter_changed(view)
        self._notify_value_committed(value, item)
        self._close_and_unfocus()
        return True

    def _commit_blank(self):
        self._clear_candidate()
        self._committed_item = None
        self._editing = False
        self._query_text = ""
        view = self._items_view()
        self._suppress_events += 1
        try:
            if view is not None:
                view.Filter = None
                try:
                    view.Refresh()
                except Exception:
                    pass
            self.combo.SelectedIndex = -1
            self._write_text("")
            self._set_caret_end()
            self.combo.IsDropDownOpen = False
        except Exception:
            return False
        finally:
            self._suppress_events -= 1

        if view is not None:
            self._notify_filter_changed(view)
        self._notify_value_committed("", None)
        self._close_and_unfocus()
        return True

    def _close_and_unfocus(self):
        try:
            self.combo.IsDropDownOpen = False
        except Exception:
            pass
        if not self.clear_focus_on_commit:
            self._focus_textbox()
            return
        if Keyboard is not None:
            try:
                Keyboard.ClearFocus()
                return
            except Exception:
                pass
        if TraversalRequest is None or FocusNavigationDirection is None:
            return
        try:
            self.combo.MoveFocus(
                TraversalRequest(FocusNavigationDirection.Next)
            )
        except Exception:
            pass

    def _items_view(self):
        """Return the view that drives the ComboBox's visible item list."""
        try:
            items = self.combo.Items
        except Exception:
            return None

        # ItemCollection already implements ICollectionView for ComboBoxes
        # populated directly.  Keeping this fast path also makes the helper
        # usable by lightweight test doubles.
        try:
            if hasattr(items, "Filter") and hasattr(items, "Refresh"):
                return items
        except Exception:
            pass

        if CollectionViewSource is None:
            return None
        try:
            source = getattr(self.combo, "ItemsSource", None)
            return CollectionViewSource.GetDefaultCollectionView(source or items)
        except Exception:
            return None

    def _capture_items_snapshot(self):
        try:
            if getattr(self.combo, "ItemsSource", None) is not None:
                return
            self._direct_items_snapshot = list(self.combo.Items)
        except Exception:
            self._direct_items_snapshot = []

    def _all_items(self):
        try:
            source = getattr(self.combo, "ItemsSource", None)
            if source is not None:
                return list(source)
        except Exception:
            pass
        return list(self._direct_items_snapshot)

    def _visible_items(self):
        try:
            return list(self.combo.Items)
        except Exception:
            return []

    def _visible_index(self, target, values=None):
        items = self._visible_items() if values is None else values
        for index, item in enumerate(items):
            if _same_item(item, target):
                return index
        return -1

    def _apply_filter(self, text, open_dropdown):
        self._query_text = _text_value(text)
        self._clear_candidate()
        view = self._items_view()
        if view is not None:
            self._suppress_events += 1
            try:
                view.Filter = (
                    (lambda value: self._matches(value))
                    if _normalized_query(self._query_text)
                    else None
                )
                try:
                    view.Refresh()
                except Exception:
                    pass
            finally:
                self._suppress_events -= 1
            self._notify_filter_changed(view)
        self._queue_highlight_refresh()
        if open_dropdown:
            try:
                self.combo.IsDropDownOpen = True
            except Exception:
                pass
        self._attach_popup()

    def _clear_filter(self, notify=True):
        self._query_text = ""
        self._clear_candidate()
        view = self._items_view()
        if view is None:
            return
        self._suppress_events += 1
        try:
            view.Filter = None
            try:
                view.Refresh()
            except Exception:
                pass
        finally:
            self._suppress_events -= 1
        if notify:
            self._notify_filter_changed(view)
        self._queue_highlight_refresh()

    def _set_container_highlight(self, index, is_highlighted):
        if ComboBoxItem is None:
            return
        try:
            container = self.combo.ItemContainerGenerator.ContainerFromIndex(index)
            if container is None:
                return
            container.SetValue(ComboBoxItem.IsHighlightedProperty, bool(is_highlighted))
        except Exception:
            # IsHighlighted is read-only in standard WPF.  Use a reversible
            # local Background resource as the visual candidate marker while
            # keeping IsSelected/SelectedItem reserved for committed values.
            try:
                background_property = ComboBoxItem.BackgroundProperty
                if is_highlighted:
                    if container not in self._candidate_backgrounds:
                        self._candidate_backgrounds[container] = container.ReadLocalValue(
                            background_property
                        )
                    container.SetResourceReference(
                        background_property,
                        "CED.Brush.CircuitItemSelectedBackground",
                    )
                else:
                    original = self._candidate_backgrounds.pop(container, None)
                    if original is None or (
                        DependencyProperty is not None
                        and original is DependencyProperty.UnsetValue
                    ):
                        container.ClearValue(background_property)
                    else:
                        container.SetValue(background_property, original)
            except Exception:
                try:
                    container.IsHighlighted = bool(is_highlighted)
                except Exception:
                    pass

    def _refresh_item_highlights(self):
        if TextBlock is None or VisualTreeHelper is None:
            return
        try:
            self.combo.UpdateLayout()
            count = int(self.combo.Items.Count)
            generator = self.combo.ItemContainerGenerator
        except Exception:
            return
        for index in range(count):
            try:
                container = generator.ContainerFromIndex(index)
                item = self.combo.Items[index]
            except Exception:
                continue
            if container is None:
                continue
            text_blocks = self._visual_descendants(container, TextBlock)
            if not text_blocks:
                continue
            self._set_item_highlight(
                text_blocks[0],
                item_text(
                    item,
                    text_getter=self._text_getter,
                    display_member_path=self._display_member_path,
                ),
            )
        if self._candidate_index >= 0:
            self._set_container_highlight(self._candidate_index, True)

    def _queue_highlight_refresh(self):
        # Fast typing can otherwise enqueue one layout/container walk per
        # character. One pending pass always reads the latest query.
        if self._highlight_refresh_pending:
            return
        self._highlight_refresh_pending = True

        def refresh_latest():
            self._highlight_refresh_pending = False
            if not self._attached:
                return
            self._refresh_item_highlights()

        self._begin_invoke(refresh_latest, priority="Background")

    def _set_item_highlight(self, text_block, value):
        if Run is None:
            return
        try:
            text_block.Text = ""
            text_block.Inlines.Clear()
            normal = self._resource_brush("CED.Brush.PrimaryText", Brushes.Black)
            match_background = self._resource_brush(
                "CED.Brush.SearchHighlightBackground",
                Brushes.Yellow,
            )
            match_foreground = self._resource_brush(
                "CED.Brush.SearchHighlightForeground",
                Brushes.Black,
            )
            for segment, matched in highlight_segments(value, self._query_text):
                run = Run(segment)
                if normal is not None:
                    run.Foreground = normal
                if matched:
                    if match_background is not None:
                        run.Background = match_background
                    if match_foreground is not None:
                        run.Foreground = match_foreground
                text_block.Inlines.Add(run)
        except Exception:
            pass

    def _resource_brush(self, key, fallback):
        try:
            value = self.combo.TryFindResource(key)
            return value if value is not None else fallback
        except Exception:
            return fallback

    def _visual_descendants(self, node, target_type):
        results = []
        queue = [node]
        while queue:
            current = queue.pop(0)
            try:
                count = int(VisualTreeHelper.GetChildrenCount(current) or 0)
            except Exception:
                count = 0
            for index in range(count):
                try:
                    child = VisualTreeHelper.GetChild(current, index)
                except Exception:
                    continue
                if isinstance(child, target_type):
                    results.append(child)
                queue.append(child)
        return results

    def _find_combo_item(self, node):
        if ComboBoxItem is None or VisualTreeHelper is None:
            return None
        current = node
        while current is not None:
            try:
                if isinstance(current, ComboBoxItem):
                    return current
                current = VisualTreeHelper.GetParent(current)
            except Exception:
                return None
        return None

    def _item_for_container(self, container):
        if container is None:
            return None
        try:
            index = self.combo.ItemContainerGenerator.IndexFromContainer(container)
            if index >= 0:
                return self.combo.Items[index]
        except Exception:
            pass
        try:
            return container.DataContext
        except Exception:
            return None

    def _on_combo_mouse_down(self, sender, args):
        source = getattr(args, "OriginalSource", None)
        if (
            self._textbox is not None
            and self._is_descendant_or_self(source, self._textbox)
            and self._committed_item is not None
            and not self._editing
            and not bool(getattr(self._textbox, "IsKeyboardFocused", False))
        ):
            self._focus_textbox()
            self._select_all()
            try:
                args.Handled = True
            except Exception:
                pass
            return
        container = self._find_combo_item(source)
        if container is None:
            return
        item = self._item_for_container(container)
        if item is None:
            return
        self._commit_item(item)
        try:
            args.Handled = True
        except Exception:
            pass

    def _parent_scroll_viewer(self):
        if ScrollViewer is None or VisualTreeHelper is None:
            return None
        current = self.combo
        while current is not None:
            try:
                current = VisualTreeHelper.GetParent(current)
            except Exception:
                return None
            if isinstance(current, ScrollViewer):
                return current
        return None

    def _on_combo_mouse_wheel(self, sender, args):
        """Keep wheel input from changing a closed editor's value.

        Popup content lives in a separate visual tree, so wheel events over
        its list never reach this handler and retain native list scrolling.
        Wheel events over the editor are redirected to the parent scroller.
        """
        if bool(getattr(self.combo, "IsDropDownOpen", False)):
            self._on_popup_mouse_wheel(sender, args)
            # Never let an open picker wheel fall through to the host grid,
            # even if its ScrollViewer has not been realized yet.
            try:
                args.Handled = True
            except Exception:
                pass
            return
        scroll_viewer = self._parent_scroll_viewer()
        if scroll_viewer is None:
            return
        try:
            delta = int(getattr(args, "Delta", 0) or 0)
            steps = max(1, int(abs(delta) // 120))
            for _unused in range(steps):
                if delta > 0:
                    scroll_viewer.LineUp()
                else:
                    scroll_viewer.LineDown()
            args.Handled = True
        except Exception:
            pass

    def _popup_scroll_viewer(self, source=None):
        if ScrollViewer is None or VisualTreeHelper is None:
            return None
        current = source
        while current is not None:
            if isinstance(current, ScrollViewer):
                return current
            if current is self._popup:
                break
            try:
                current = VisualTreeHelper.GetParent(current)
            except Exception:
                break
        for node in self._visual_descendants(self._popup, ScrollViewer):
            return node
        return None

    def _on_popup_mouse_wheel(self, sender, args):
        """Scroll the open popup list and stop the wheel reaching its host grid."""
        scroll_viewer = self._popup_scroll_viewer(getattr(args, "OriginalSource", None))
        if scroll_viewer is None:
            return
        try:
            delta = int(getattr(args, "Delta", 0) or 0)
            steps = max(1, int(abs(delta) // 120))
            for _unused in range(steps):
                if delta > 0:
                    scroll_viewer.LineUp()
                else:
                    scroll_viewer.LineDown()
            args.Handled = True
        except Exception:
            pass

    def focus_editor(self, select_all=False):
        """Focus the editable surface for host-controlled grid navigation."""
        self._attach_textbox()
        self._focus_textbox()
        if select_all:
            self._select_all()

    def _on_popup_mouse_down(self, sender, args):
        # Routed mouse events inside PART_Popup do not reach the ComboBox.
        self._on_combo_mouse_down(sender, args)

    def _is_descendant_or_self(self, node, ancestor):
        current = node
        while current is not None:
            if current is ancestor:
                return True
            if VisualTreeHelper is None:
                return False
            try:
                current = VisualTreeHelper.GetParent(current)
            except Exception:
                return False
        return False

    def _notify_filter_changed(self, view):
        callback = self._on_filter_changed_callback
        if callback is None:
            return
        try:
            count = int(view.Count)
        except Exception:
            count = len(self._visible_items())
        try:
            callback(self, self._query_text, count)
        except Exception:
            pass

    def _notify_value_committed(self, value, item):
        callback = self._on_value_committed_callback
        if callback is None:
            return
        try:
            callback(self, value, item)
        except Exception:
            pass

    def _begin_invoke(self, callback, priority="Background"):
        if Action is None or DispatcherPriority is None:
            callback()
            return
        try:
            dispatcher_priority = getattr(
                DispatcherPriority,
                priority,
                DispatcherPriority.Background,
            )
            self.combo.Dispatcher.BeginInvoke(
                dispatcher_priority,
                Action(callback),
            )
        except Exception:
            callback()

    def _matches(self, value):
        needle = _normalized_query(self._query_text)
        if not needle:
            return True
        return needle in item_text(
            value,
            text_getter=self._text_getter,
            display_member_path=self._display_member_path,
        ).lower()
