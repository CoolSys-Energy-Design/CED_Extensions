# -*- coding: utf-8 -*-
"""Generic structured-search state and query objects.

This module intentionally has no WPF or Revit dependencies.  It contains the
small query model used by :mod:`UIClasses.structured_search_box` and can also be
used by a host that wants to provide its own presentation.

The host owns the item model and matching rules:

* ``SearchFilterDefinition.matcher`` receives ``(item, value)``.
* ``SearchQuery.matches`` receives a free-text matcher when the host wants the
  query object to evaluate the ordinary search portion.

The state object keeps incomplete filter tokens out of the effective query.
That is important for live search: choosing ``/panel`` creates a visual token,
but does not change results until the token has a meaningful value.
"""

from __future__ import print_function


try:
    _string_types = (basestring,)
except NameError:
    _string_types = (str,)


def _text(value):
    """Return a safe text representation on IronPython 2 and Python 3."""
    if value is None:
        return ""
    if isinstance(value, _string_types):
        return value
    try:
        return unicode(value)  # noqa: F821 - defined on IronPython 2
    except NameError:
        return str(value)
    except Exception:
        return str(value)


def _normalized_text(value):
    return _text(value).strip()


def _normalized_key(value):
    return _normalized_text(value).lower()


def _contains_text(value, query_text):
    """Case-insensitive substring helper for host-defined free-text fields."""
    needle = _normalized_text(query_text).lower()
    if not needle:
        return True
    return needle in _text(value).lower()


def make_contains_matcher(value_getter):
    """Build a free-text matcher from a host-owned item value getter.

    ``value_getter`` may return one field or a host-composed string containing
    all fields that should participate in ordinary search.
    """
    if not callable(value_getter):
        raise TypeError("value_getter must be callable")

    def _matcher(item, query_text):
        return _contains_text(value_getter(item), query_text)

    return _matcher


class SearchFilterDefinition(object):
    """Definition supplied by a searchable host.

    Args:
        key: Stable identifier used by the host.  It is never displayed unless
            the host chooses to display it.
        display_name: Human-readable command/token label.
        matcher: Optional callable ``matcher(item, value) -> bool``.
        placeholder: Optional value-entry hint shown by the WPF control.
        value_formatter: Optional callable ``formatter(value) -> text`` used
            only for token display.  Matching still receives the raw text.
        value_hint: Optional accepted-value guidance shown in the command
            picker.
        allow_multiple: Whether selecting this command appends another token
            instead of reactivating the existing token with the same key.
        combine_mode: ``"and"`` for each value or ``"or"`` for values of the
            same key.  When omitted, multi-value definitions default to OR.
        token_style_key: Optional resource key for the normal token border.
        selected_token_style_key: Optional resource key for the selected or
            active token border.
        token_text_brush_key: Optional resource key for token text and caret.
    """

    def __init__(
        self,
        key,
        display_name,
        matcher=None,
        placeholder=None,
        value_formatter=None,
        value_hint=None,
        allow_multiple=False,
        combine_mode=None,
        token_style_key=None,
        selected_token_style_key=None,
        token_text_brush_key=None,
    ):
        normalized_key = _normalized_text(key)
        normalized_name = _normalized_text(display_name)
        if not normalized_key:
            raise ValueError("A search filter definition requires a key")
        if not normalized_name:
            raise ValueError("A search filter definition requires a display name")
        if matcher is not None and not callable(matcher):
            raise TypeError("matcher must be callable")
        if value_formatter is not None and not callable(value_formatter):
            raise TypeError("value_formatter must be callable")

        self.key = normalized_key
        self.display_name = normalized_name
        self.matcher = matcher
        self.placeholder = _normalized_text(placeholder) or "Enter {}".format(normalized_name)
        self.value_formatter = value_formatter
        self.value_hint = _normalized_text(value_hint)
        self.allow_multiple = bool(allow_multiple)
        requested_combine_mode = _normalized_key(combine_mode)
        if not requested_combine_mode:
            requested_combine_mode = "or" if self.allow_multiple else "and"
        if requested_combine_mode not in ("and", "or"):
            raise ValueError("combine_mode must be 'and' or 'or'")
        self.combine_mode = requested_combine_mode
        self.token_style_key = (
            _normalized_text(token_style_key) or "CED.SearchBox.Token"
        )
        self.selected_token_style_key = (
            _normalized_text(selected_token_style_key)
            or "CED.SearchBox.Token.Selected"
        )
        self.token_text_brush_key = (
            _normalized_text(token_text_brush_key)
            or "CED.Brush.BadgeInfoText"
        )

    @property
    def Key(self):
        return self.key

    @property
    def DisplayName(self):
        return self.display_name

    @property
    def Placeholder(self):
        return self.placeholder

    @property
    def ValueHint(self):
        return self.value_hint

    @property
    def AllowMultiple(self):
        return self.allow_multiple

    @property
    def CombineMode(self):
        return self.combine_mode

    @property
    def TokenStyleKey(self):
        return self.token_style_key

    @property
    def SelectedTokenStyleKey(self):
        return self.selected_token_style_key

    @property
    def TokenTextBrushKey(self):
        return self.token_text_brush_key

    def format_value(self, value):
        if self.value_formatter is not None:
            return _text(self.value_formatter(value))
        return _text(value)

    def matches(self, item, value):
        """Evaluate this definition against a host item.

        The framework does not invent matching behavior.  A definition without
        a matcher is therefore deliberately non-matching when this convenience
        method is used; hosts that only consume ``SearchQuery.filters`` can
        ignore this method and evaluate their own query.
        """
        if self.matcher is None:
            return False
        return bool(self.matcher(item, value))

    def __repr__(self):
        return "SearchFilterDefinition({!r}, {!r})".format(self.key, self.display_name)


class SearchFilter(object):
    """A meaningful, applied filter value in a :class:`SearchQuery`."""

    def __init__(self, definition, value):
        if not isinstance(definition, SearchFilterDefinition):
            raise TypeError("definition must be SearchFilterDefinition")
        normalized_value = _normalized_text(value)
        if not normalized_value:
            raise ValueError("A SearchFilter requires a meaningful value")
        self.definition = definition
        self.key = definition.key
        self.display_name = definition.display_name
        self.value = normalized_value

    @property
    def Key(self):
        return self.key

    @property
    def DisplayName(self):
        return self.display_name

    @property
    def Value(self):
        return self.value

    def matches(self, item):
        return self.definition.matches(item, self.value)

    def __eq__(self, other):
        return (
            isinstance(other, SearchFilter)
            and self.key == other.key
            and self.value == other.value
        )

    def __ne__(self, other):
        return not self == other

    def __repr__(self):
        return "SearchFilter({!r}, {!r})".format(self.key, self.value)


class SearchFilterToken(object):
    """Draft token used by the interaction state before it enters a query."""

    def __init__(self, definition, value=""):
        if not isinstance(definition, SearchFilterDefinition):
            raise TypeError("definition must be SearchFilterDefinition")
        self.definition = definition
        self.value = _text(value)

    @property
    def key(self):
        return self.definition.key

    @property
    def display_name(self):
        return self.definition.display_name

    @property
    def meaningful(self):
        return bool(_normalized_text(self.value))

    @property
    def display_value(self):
        return self.definition.format_value(self.value)

    def to_filter(self):
        if not self.meaningful:
            return None
        return SearchFilter(self.definition, self.value)

    def __repr__(self):
        return "SearchFilterToken({!r}, {!r})".format(self.key, self.value)


class SearchQuery(object):
    """Effective live-search query.

    ``filters`` contains only meaningful ``SearchFilter`` instances.  The
    query therefore never exposes a half-entered command or empty token.
    Conditions are ANDed by default; definitions configured with
    ``combine_mode="or"`` are ORed with other values of the same key.
    """

    def __init__(self, free_text="", filters=None):
        self.free_text = _normalized_text(free_text)
        self.filters = tuple(filters or ())
        for item in self.filters:
            if not isinstance(item, SearchFilter):
                raise TypeError("SearchQuery.filters must contain SearchFilter objects")

    @property
    def FreeText(self):
        """PascalCase compatibility alias for WPF/host code."""
        return self.free_text

    @property
    def Filters(self):
        """PascalCase compatibility alias for WPF/host code."""
        return self.filters

    @property
    def is_empty(self):
        return not self.free_text and not self.filters

    @property
    def has_filters(self):
        return bool(self.filters)

    def filter_for_key(self, key):
        normalized = _normalized_key(key)
        for item in self.filters:
            if item.key.lower() == normalized:
                return item
        return None

    def matches(self, item, free_text_matcher=None):
        """Evaluate this query with host-supplied free-text semantics.

        ``free_text_matcher`` receives ``(item, free_text)``.  Structured
        filter definitions receive ``(item, value)`` through their own
        callbacks.  Conditions are ANDed across keys, while values in an OR
        group for the same key are alternatives.
        """
        if self.free_text:
            if free_text_matcher is None:
                raise ValueError("free_text_matcher is required for a non-empty free-text query")
            if not bool(free_text_matcher(item, self.free_text)):
                return False
        or_groups = {}
        for item_filter in self.filters:
            if item_filter.definition.combine_mode == "or":
                key = item_filter.key.lower()
                or_groups.setdefault(key, []).append(item_filter)
            elif not item_filter.matches(item):
                return False
        for filters in or_groups.values():
            if not any(item_filter.matches(item) for item_filter in filters):
                return False
        return True

    def __eq__(self, other):
        return (
            isinstance(other, SearchQuery)
            and self.free_text == other.free_text
            and self.filters == other.filters
        )

    def __ne__(self, other):
        return not self == other

    def __repr__(self):
        return "SearchQuery(free_text={!r}, filters={!r})".format(self.free_text, self.filters)


class SearchQueryChangedEventArgs(object):
    """Event payload emitted when the effective query changes."""

    def __init__(self, query, reason="changed"):
        self.query = query
        self.reason = _text(reason) or "changed"

    @property
    def Query(self):
        return self.query


class SearchCommandChangedEventArgs(object):
    """Event payload emitted while the slash-command picker is active."""

    def __init__(self, command_text, suggestions):
        self.command_text = _text(command_text)
        self.suggestions = tuple(suggestions or ())

    @property
    def CommandText(self):
        return self.command_text

    @property
    def Suggestions(self):
        return self.suggestions


class SearchInteractionChangedEventArgs(object):
    """Event payload for token selection/editing changes that do not alter results."""

    def __init__(self, state, reason="changed"):
        self.tokens = state.tokens
        self.selected_token_index = state.selected_token_index
        self.active_token_index = state.active_token_index
        self.reason = _text(reason) or "changed"


class StructuredSearchState(object):
    """Small, presentation-independent state machine for structured search.

    The WPF control is one consumer.  Keeping this state separate makes the
    slash command behavior, incomplete-token semantics, and atomic deletion
    testable without Revit or a desktop UI.
    """

    MODE_TEXT = "text"
    MODE_COMMAND = "command"

    def __init__(self, filter_definitions=None):
        self._filter_definitions = tuple(filter_definitions or ())
        for definition in self._filter_definitions:
            if not isinstance(definition, SearchFilterDefinition):
                raise TypeError("filter_definitions must contain SearchFilterDefinition objects")
        self._free_text = ""
        self._tokens = []
        self._mode = self.MODE_TEXT
        self._command_text = ""
        self._command_prefix = ""
        self._literal_slash_context = ""
        self._literal_slash_index = None
        self._selected_token_index = None
        self._active_token_index = None
        self._last_query = SearchQuery()
        self._query_handlers = []
        self._command_handlers = []
        self._interaction_handlers = []

    @property
    def filter_definitions(self):
        return self._filter_definitions

    @property
    def tokens(self):
        return tuple(self._tokens)

    @property
    def free_text(self):
        return self._free_text

    @property
    def mode(self):
        return self._mode

    @property
    def is_command_mode(self):
        return self._mode == self.MODE_COMMAND

    @property
    def command_text(self):
        return self._command_text

    @property
    def command_input_text(self):
        """The full text shown while the provisional slash picker is active."""
        if not self.is_command_mode:
            return self._free_text
        return self._command_prefix + "/" + self._command_text

    @property
    def CommandInputText(self):
        return self.command_input_text

    @property
    def selected_token_index(self):
        return self._selected_token_index

    @property
    def active_token_index(self):
        return self._active_token_index

    @property
    def command_suggestions(self):
        needle = _normalized_key(self._command_text)
        if not needle:
            return self._filter_definitions
        return tuple(
            definition
            for definition in self._filter_definitions
            if needle in definition.display_name.lower() or needle in definition.key.lower()
        )

    @property
    def query(self):
        return self._build_query()

    @property
    def has_content(self):
        return bool(
            self._free_text
            or self._tokens
            or self._command_text
            or self._command_prefix
            or self.is_command_mode
        )

    def add_query_changed_handler(self, handler):
        if callable(handler) and handler not in self._query_handlers:
            self._query_handlers.append(handler)
        return handler

    def set_filter_definitions(self, filter_definitions):
        """Replace the cached command definitions supplied by the host."""
        definitions = tuple(filter_definitions or ())
        for definition in definitions:
            if not isinstance(definition, SearchFilterDefinition):
                raise TypeError("filter_definitions must contain SearchFilterDefinition objects")
        self._filter_definitions = definitions
        if self.is_command_mode:
            self._emit_command_changed()
        self._emit_interaction_changed("filter_definitions_changed")

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

    def add_interaction_changed_handler(self, handler):
        if callable(handler) and handler not in self._interaction_handlers:
            self._interaction_handlers.append(handler)
        return handler

    def remove_interaction_changed_handler(self, handler):
        if handler in self._interaction_handlers:
            self._interaction_handlers.remove(handler)

    def _emit_query_if_changed(self, reason):
        query = self._build_query()
        if query == self._last_query:
            return False
        self._last_query = query
        args = SearchQueryChangedEventArgs(query, reason=reason)
        for handler in list(self._query_handlers):
            handler(self, args)
        return True

    def _emit_command_changed(self):
        args = SearchCommandChangedEventArgs(
            self._command_text,
            self.command_suggestions,
        )
        for handler in list(self._command_handlers):
            handler(self, args)

    def _emit_interaction_changed(self, reason):
        args = SearchInteractionChangedEventArgs(self, reason=reason)
        for handler in list(self._interaction_handlers):
            handler(self, args)

    def _build_query(self):
        filters = []
        for token in self._tokens:
            item_filter = token.to_filter()
            if item_filter is not None:
                filters.append(item_filter)
        return SearchQuery(self._free_text, filters)

    def _definition_for(self, definition_or_key):
        if isinstance(definition_or_key, SearchFilterDefinition):
            for definition in self._filter_definitions:
                if definition is definition_or_key or definition.key.lower() == definition_or_key.key.lower():
                    return definition
            return None
        key = _normalized_key(definition_or_key)
        for definition in self._filter_definitions:
            if definition.key.lower() == key:
                return definition
        return None

    def _token_group_order(self, token):
        for index, definition in enumerate(self._filter_definitions):
            if definition.key.lower() == token.key.lower():
                return index
        return len(self._filter_definitions)

    def _organize_tokens(self):
        """Keep the visual/query order grouped by registered field type."""
        self._tokens = sorted(self._tokens, key=self._token_group_order)

    def _command_start(self, text):
        raw = _text(text)
        for slash_index in range(len(raw) - 1, -1, -1):
            if raw[slash_index] != "/":
                continue
            if slash_index == 0 or raw[slash_index - 1].isspace():
                return slash_index
        return None

    def set_input_text(self, text):
        """Accept text from the active input portion of a search control.

        A slash starts command mode only at the beginning of the input or when
        preceded by whitespace.  Any text before that slash remains ordinary
        free text.  If a slash fragment is dismissed or has no matching
        command, it is remembered as literal text while the user continues
        typing so the same slash does not reopen the picker on every change.
        """
        raw = _text(text)
        if self.is_command_mode:
            command_prefix = self._command_prefix + "/"
            if raw.startswith(command_prefix):
                self.set_command_text(raw[len(command_prefix):])
            else:
                # Editing/removing the provisional slash turns the entire
                # input back into ordinary free text.
                self.set_free_text(raw)
            return

        if self._literal_slash_context:
            slash_index = self._command_start(raw)
            if (
                len(raw) >= len(self._literal_slash_context)
                and raw.startswith(self._literal_slash_context)
                and (slash_index is None or slash_index <= self._literal_slash_index)
            ):
                self.set_free_text(raw)
                self._literal_slash_context = raw
                return
            self._literal_slash_context = ""
            self._literal_slash_index = None

        command_start = self._command_start(raw)
        if command_start is not None:
            prefix = raw[:command_start]
            self.set_free_text(prefix)
            self.begin_command(raw[command_start + 1:], prefix=prefix)
            if not self.command_suggestions:
                self.cancel_command_as_literal()
            return
        self.set_free_text(raw)

    def set_free_text(self, text):
        # Keep the live editor value verbatim so spaces can be entered at the
        # caret.  SearchQuery normalizes the value for matching/equality, and
        # host matchers are expected to trim their own needle as needed.
        value = _text(text)
        changed = value != self._free_text
        self._free_text = value
        self._literal_slash_context = ""
        self._literal_slash_index = None
        self._selected_token_index = None
        if self.is_command_mode:
            self._mode = self.MODE_TEXT
            self._command_text = ""
            self._command_prefix = ""
            self._emit_command_changed()
        if changed:
            self._emit_query_if_changed("free_text_changed")
        self._emit_interaction_changed("free_text_changed")

    def begin_command(self, command_text="", prefix=None):
        self._mode = self.MODE_COMMAND
        self._command_text = _text(command_text)
        self._command_prefix = _text(self._free_text if prefix is None else prefix)
        self._literal_slash_context = ""
        self._literal_slash_index = None
        self._selected_token_index = None
        self._active_token_index = None
        self._emit_command_changed()
        self._emit_interaction_changed("command_started")

    def set_command_text(self, command_text):
        if not self.is_command_mode:
            self.begin_command(command_text)
            return
        value = _text(command_text)
        if value == self._command_text:
            return
        self._command_text = value
        # A command keyword cannot contain whitespace.  Once the user types
        # it, the slash fragment is ordinary free text again; this also makes
        # spaces after an unmatched slash behave like normal search input.
        if any(character.isspace() for character in value):
            self.cancel_command_as_literal()
            return
        self._emit_command_changed()
        if not self.command_suggestions:
            self.cancel_command_as_literal()

    def cancel_command(self):
        if not self.is_command_mode:
            return False
        self._mode = self.MODE_TEXT
        self._command_text = ""
        self._command_prefix = ""
        self._literal_slash_context = ""
        self._literal_slash_index = None
        self._emit_command_changed()
        self._emit_interaction_changed("command_cancelled")
        return True

    def cancel_command_as_literal(self):
        """Leave slash mode and keep its complete input as free text."""
        if not self.is_command_mode:
            return False
        literal_prefix = self._command_prefix
        literal = self._command_prefix + "/" + self._command_text
        self._mode = self.MODE_TEXT
        self._command_text = ""
        self._command_prefix = ""
        self._free_text = _text(literal)
        self._literal_slash_context = literal
        self._literal_slash_index = len(literal_prefix)
        self._selected_token_index = None
        self._emit_command_changed()
        self._emit_query_if_changed("free_text_changed")
        self._emit_interaction_changed("command_cancelled_as_text")
        return True

    def select_filter(self, definition_or_key):
        definition = self._definition_for(definition_or_key)
        if definition is None:
            return None
        if not definition.allow_multiple:
            for index, existing in enumerate(self._tokens):
                if existing.key.lower() == definition.key.lower():
                    self._active_token_index = index
                    self._selected_token_index = None
                    self._mode = self.MODE_TEXT
                    self._command_text = ""
                    self._command_prefix = ""
                    self._literal_slash_context = ""
                    self._literal_slash_index = None
                    self._emit_command_changed()
                    self._emit_interaction_changed("filter_reactivated")
                    return existing
        new_token = SearchFilterToken(definition)
        self._tokens.append(new_token)
        self._organize_tokens()
        self._active_token_index = self._tokens.index(new_token)
        self._selected_token_index = None
        self._mode = self.MODE_TEXT
        self._command_text = ""
        self._command_prefix = ""
        self._literal_slash_context = ""
        self._literal_slash_index = None
        self._emit_command_changed()
        self._emit_interaction_changed("filter_selected")
        # An empty token is intentionally omitted from the query.
        return new_token

    def set_active_token_value(self, value):
        index = self._active_token_index
        if index is None or index < 0 or index >= len(self._tokens):
            return False
        token = self._tokens[index]
        new_value = _text(value)
        if token.value == new_value:
            return False
        token.value = new_value
        self._selected_token_index = None
        self._emit_query_if_changed("filter_value_changed")
        self._emit_interaction_changed("filter_value_changed")
        return True

    def edit_token(self, index):
        if index is None or index < 0 or index >= len(self._tokens):
            return False
        self._active_token_index = int(index)
        self._selected_token_index = None
        self._emit_interaction_changed("token_editing_started")
        return True

    def commit_active_token(self):
        if self._active_token_index is None:
            return False
        self._active_token_index = None
        self._emit_interaction_changed("token_editing_finished")
        return True

    def select_token(self, index):
        if index is None or index < 0 or index >= len(self._tokens):
            return False
        self._selected_token_index = int(index)
        self._active_token_index = None
        self._emit_interaction_changed("token_selected")
        return True

    def clear_token_selection(self):
        if self._selected_token_index is None:
            return False
        self._selected_token_index = None
        self._emit_interaction_changed("token_deselected")
        return True

    def remove_token(self, index):
        if index is None or index < 0 or index >= len(self._tokens):
            return False
        del self._tokens[index]
        if self._active_token_index == index:
            self._active_token_index = None
        elif self._active_token_index is not None and self._active_token_index > index:
            self._active_token_index -= 1
        if self._selected_token_index == index:
            self._selected_token_index = None
        elif self._selected_token_index is not None and self._selected_token_index > index:
            self._selected_token_index -= 1
        self._emit_query_if_changed("filter_removed")
        self._emit_interaction_changed("filter_removed")
        return True

    def remove_selected_token(self):
        if self._selected_token_index is None:
            return False
        return self.remove_token(self._selected_token_index)

    def backspace_at_input_start(self):
        """Select the adjacent token, then remove it on a second Backspace."""
        if self._selected_token_index is not None:
            return self.remove_selected_token()
        if not self._tokens:
            return False
        return self.select_token(len(self._tokens) - 1)

    def delete_at_input_start(self):
        """Delete behaves like Backspace for the token beside the input edge."""
        if self._selected_token_index is not None:
            return self.remove_selected_token()
        if not self._tokens:
            return False
        return self.select_token(len(self._tokens) - 1)

    def clear(self):
        had_content = self.has_content
        self._free_text = ""
        self._tokens = []
        self._mode = self.MODE_TEXT
        self._command_text = ""
        self._command_prefix = ""
        self._literal_slash_context = ""
        self._literal_slash_index = None
        self._selected_token_index = None
        self._active_token_index = None
        self._emit_command_changed()
        self._emit_query_if_changed("cleared")
        if had_content:
            self._emit_interaction_changed("cleared")
        return had_content


__all__ = [
    "SearchFilterDefinition",
    "SearchFilter",
    "SearchFilterToken",
    "SearchQuery",
    "SearchQueryChangedEventArgs",
    "SearchCommandChangedEventArgs",
    "SearchInteractionChangedEventArgs",
    "StructuredSearchState",
    "make_contains_matcher",
]
