# -*- coding: utf-8 -*-
"""One reusable settings window, so a tool asks everything up front.

Several CED tools were built as a chain of modal prompts - a CommandSwitchWindow
for the mode, a SelectFromList for the type, an ask_for_string for a number -
which means the user cannot see their choices together, cannot go back, and
cannot tell how many questions are left. This replaces that chain with a single
themed window built from a field spec.

    from UIClasses import settings_form as sf

    result = sf.prompt(
        "Space Coils",
        [
            sf.choice("mode", "Placement", [("wall", "Along a wall"),
                                            ("center", "Centred grid")]),
            sf.number("offset", "Offset from wall (inches)", default=15,
                      minimum=0),
            sf.boolean("mixed", "Grid mixes coil sizes", default=False),
        ],
        caller_file=__file__,
        subtitle="Coils must be selected before you run this.",
    )
    if result is None:
        return              # cancelled - write nothing
    mode = result["mode"]

``prompt`` returns a dict of field key -> value, or None when cancelled. It
never raises for user input: validation is reported in the window and blocks
Continue instead.

The window carries the shared Help and Docs buttons, so every tool converted to
this shell gets its guided tour and its reference for free.
"""

import os

import clr

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

from System.Windows import TextWrapping, Thickness, Visibility
from System.Windows.Controls import (
    Border,
    CheckBox,
    ComboBox,
    ComboBoxItem,
    Grid,
    Orientation,
    ScrollBarVisibility,
    ScrollViewer,
    StackPanel,
    TextBlock,
    TextBox,
)
from System.Windows.Controls import Button as WpfButton
from System.Windows.Controls import ColumnDefinition

from pyrevit import forms, script

from UIClasses import pathing as ui_pathing
from UIClasses import resource_loader
from UIClasses.ui_bases import CEDWindowBase, load_theme_state_from_config

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
XAML_PATH = os.path.join(THIS_DIR, "SettingsForm.xaml")

_LOGGER = script.get_logger()

KIND_CHOICE = "choice"
KIND_MULTICHOICE = "multichoice"
KIND_TEXT = "text"
KIND_MULTILINE = "multiline"
KIND_INT = "int"
KIND_FLOAT = "float"
KIND_BOOL = "bool"
KIND_OPEN_FILE = "open_file"
KIND_SAVE_FILE = "save_file"
KIND_INFO = "info"


# ---------------------------------------------------------------------------
# field spec
# ---------------------------------------------------------------------------

class Field(object):
    """One input in a settings form. Prefer the helpers below."""

    def __init__(self, key, label, kind, default=None, options=None, hint=None,
                 required=True, minimum=None, maximum=None, file_filter=None,
                 max_length=None, height=None, visible_when=None,
                 file_ext=None, options_from=None):
        self.key = key
        self.label = label
        self.kind = kind
        self.default = default
        self.options = list(options or [])
        self.hint = hint
        self.required = bool(required)
        self.minimum = minimum
        self.maximum = maximum
        self.file_filter = file_filter
        # pyRevit's pick_file/save_file want a bare extension ("xlsx") in
        # file_ext; files_filter is a full "Description|*.ext" pair. Passing a
        # glob as the filter silently gives you a dialog that will not open.
        self.file_ext = file_ext
        # ("other_field_key", provider) - provider(value) returns this field's
        # options, recomputed whenever that other field changes.
        self.options_from = options_from
        self.max_length = max_length
        self.height = height
        # ("other_field_key", [values]) - this field only shows while that
        # field holds one of those values.
        self.visible_when = visible_when
        self.control = None          # set when the form is built
        self.container = None        # the row, so it can be shown/hidden
        self._extra = None           # per-kind helper state


def _normalize_options(options):
    """Accept ['a','b'] or [('a','Label A'), ...] and return pairs."""
    pairs = []
    for option in list(options or []):
        if isinstance(option, (tuple, list)) and len(option) >= 2:
            pairs.append((option[0], str(option[1])))
        else:
            pairs.append((option, str(option)))
    return pairs


def choice(key, label, options, default=None, hint=None, required=True,
           visible_when=None, options_from=None):
    """Dropdown. ``options_from=("other_key", fn)`` makes it depend on another
    field: fn(that field's value) returns this one's options."""
    return Field(key, label, KIND_CHOICE, default=default,
                 options=_normalize_options(options), hint=hint, required=required,
                 visible_when=visible_when, options_from=options_from)


def multichoice(key, label, options, default=None, hint=None, required=True,
                      visible_when=None):
    """Checklist. ``default`` is a list of selected values (None = all)."""
    return Field(key, label, KIND_MULTICHOICE, default=default,
                 options=_normalize_options(options), hint=hint, required=required,
                 visible_when=visible_when)


def text(key, label, default="", hint=None, required=True, max_length=None,
         visible_when=None):
    return Field(key, label, KIND_TEXT, default=default, hint=hint,
                 required=required, max_length=max_length,
                 visible_when=visible_when)


def multiline(key, label, default="", hint=None, required=True, height=120,
              visible_when=None):
    return Field(key, label, KIND_MULTILINE, default=default, hint=hint,
                 required=required, height=height,
                 visible_when=visible_when)


def number(key, label, default=None, minimum=None, maximum=None, hint=None,
           required=True, decimals=False, visible_when=None):
    return Field(key, label, KIND_FLOAT if decimals else KIND_INT,
                 default=default, minimum=minimum, maximum=maximum,
                 hint=hint, required=required, visible_when=visible_when)


def boolean(key, label, default=False, hint=None, visible_when=None):
    return Field(key, label, KIND_BOOL, default=bool(default), hint=hint,
                 required=False, visible_when=visible_when)


def open_file(key, label, file_ext=None, file_filter=None, default="", hint=None,
              required=True, visible_when=None):
    """Browse-for-a-file row. ``file_ext`` is a bare extension, e.g. "xlsx"."""
    return Field(key, label, KIND_OPEN_FILE, default=default, hint=hint,
                 required=required, file_filter=file_filter, file_ext=file_ext,
                 visible_when=visible_when)


def save_file(key, label, file_ext=None, file_filter=None, default="", hint=None,
              required=True, visible_when=None):
    """Save-as row. ``file_ext`` is a bare extension, e.g. "xlsx"."""
    return Field(key, label, KIND_SAVE_FILE, default=default, hint=hint,
                 required=required, file_filter=file_filter, file_ext=file_ext,
                 visible_when=visible_when)


def info(label, hint=None):
    return Field(None, label, KIND_INFO, hint=hint, required=False)


# ---------------------------------------------------------------------------
# window
# ---------------------------------------------------------------------------

class SettingsFormWindow(CEDWindowBase):
    theme_aware = True

    def __init__(self, title, fields, subtitle=None, ok_label="Continue",
                 caller_file=None, help_modal=True, validate=None):
        self._fields = [f for f in fields if f is not None]
        self._caller_file = caller_file
        self._validate = validate
        self.result = None

        theme_mode, accent_mode = load_theme_state_from_config(
            default_theme="light", default_accent="neutral")
        CEDWindowBase.__init__(self, xaml_source=XAML_PATH,
                               theme_mode=theme_mode, accent_mode=accent_mode,
                               theme_aware=True, handle_esc=True)
        self.Title = str(title or "Settings")

        self._title_text = self.FindName("FormTitleText")
        self._subtitle_text = self.FindName("FormSubtitleText")
        self._host = self.FindName("FieldHost")
        self._validation_host = self.FindName("ValidationHost")
        self._validation_text = self.FindName("ValidationText")
        self._ok = self.FindName("OkButton")
        self._cancel = self.FindName("CancelButton")
        self._reset = self.FindName("ResetButton")

        if self._title_text is not None:
            self._title_text.Text = str(title or "Settings")
        if subtitle and self._subtitle_text is not None:
            self._subtitle_text.Text = str(subtitle)
            self._subtitle_text.Visibility = Visibility.Visible
        if self._ok is not None:
            self._ok.Content = str(ok_label or "Continue")
            self._ok.Click += self._ok_clicked
        if self._cancel is not None:
            self._cancel.Click += self._cancel_clicked
        if self._reset is not None:
            self._reset.Click += self._reset_clicked

        self._suppress_dependents = False
        self._build_fields()
        self._refresh_dependent_options()
        self._apply_visibility()

        # Help and Docs. This window is modal, so the tour has to be too.
        if caller_file:
            try:
                from UIClasses.tutorial import wire_docs, wire_help

                wire_help(self, caller_file, modal=help_modal)
                wire_docs(self, caller_file)
            except Exception as exc:
                _LOGGER.debug("Settings form help/docs not wired: %s", exc)

    # -- construction ------------------------------------------------------

    def _style(self, key):
        try:
            return self.TryFindResource(key)
        except Exception:
            return None

    def _label_block(self, field):
        block = TextBlock()
        block.Text = str(field.label or "")
        style = self._style("CED.Text.CaptionStrong")
        if style is not None:
            block.Style = style
        block.TextWrapping = TextWrapping.Wrap
        block.Margin = Thickness(0, 0, 0, 3)
        return block

    def _hint_block(self, field):
        if not field.hint:
            return None
        block = TextBlock()
        block.Text = str(field.hint)
        style = self._style("CED.Text.Caption")
        if style is not None:
            block.Style = style
        block.TextWrapping = TextWrapping.Wrap
        block.Margin = Thickness(0, 2, 0, 0)
        return block

    def _build_fields(self):
        if self._host is None:
            return
        self._host.Children.Clear()
        for field in self._fields:
            container = StackPanel()
            container.Orientation = Orientation.Vertical
            container.Margin = Thickness(0, 0, 0, 12)

            if field.kind == KIND_INFO:
                border = Border()
                border.Background = self._style("CED.Brush.InfoPanelBackground")
                border.BorderBrush = self._style("CED.Brush.InfoPanelBorder")
                border.BorderThickness = Thickness(1)
                radius = self._style("CED.CornerRadius.Control")
                if radius is not None:
                    border.CornerRadius = radius
                border.Padding = Thickness(8, 6, 8, 6)
                inner = TextBlock()
                inner.Text = str(field.label or "")
                inner.TextWrapping = TextWrapping.Wrap
                style = self._style("CED.Text.Secondary")
                if style is not None:
                    inner.Style = style
                border.Child = inner
                container.Children.Add(border)
                self._host.Children.Add(container)
                continue

            if field.kind != KIND_BOOL:
                container.Children.Add(self._label_block(field))

            control = self._build_control(field)
            if control is not None:
                container.Children.Add(control)
            if field.kind == KIND_CHOICE and control is not None:
                control.SelectionChanged += self._choice_changed
            hint = self._hint_block(field)
            if hint is not None:
                container.Children.Add(hint)
            field.container = container
            self._host.Children.Add(container)

    def _build_control(self, field):
        kind = field.kind
        if kind == KIND_CHOICE:
            combo = ComboBox()
            combo.Height = 26
            for value, label in field.options:
                item = ComboBoxItem()
                item.Content = label
                item.Tag = value
                combo.Items.Add(item)
            combo.SelectedIndex = self._default_index(field)
            field.control = combo
            return combo

        if kind == KIND_MULTICHOICE:
            border = Border()
            border.BorderBrush = self._style("CED.Brush.ListBorder")
            border.BorderThickness = Thickness(1)
            border.Background = self._style("CED.Brush.ListBackground")
            border.Padding = Thickness(6, 4, 6, 4)
            scroller = ScrollViewer()
            scroller.MaxHeight = 150
            scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            panel = StackPanel()
            selected = field.default
            boxes = []
            for value, label in field.options:
                box = CheckBox()
                box.Content = label
                box.Tag = value
                box.Margin = Thickness(0, 2, 0, 2)
                box.IsChecked = True if selected is None else (value in selected)
                panel.Children.Add(box)
                boxes.append(box)
            scroller.Content = panel
            border.Child = scroller
            field.control = border
            field._extra = boxes
            return border

        if kind in (KIND_TEXT, KIND_INT, KIND_FLOAT, KIND_MULTILINE):
            box = TextBox()
            if kind == KIND_MULTILINE:
                box.AcceptsReturn = True
                box.TextWrapping = TextWrapping.Wrap
                box.Height = float(field.height or 120)
                box.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            else:
                box.Height = 26
            if field.max_length:
                box.MaxLength = int(field.max_length)
            box.Text = "" if field.default is None else str(field.default)
            field.control = box
            return box

        if kind == KIND_BOOL:
            box = CheckBox()
            box.Content = str(field.label or "")
            box.IsChecked = bool(field.default)
            field.control = box
            return box

        if kind in (KIND_OPEN_FILE, KIND_SAVE_FILE):
            grid = Grid()
            grid.ColumnDefinitions.Add(ColumnDefinition())
            auto = ColumnDefinition()
            auto.Width = self._auto_length()
            grid.ColumnDefinitions.Add(auto)
            box = TextBox()
            box.Height = 26
            box.Text = "" if field.default is None else str(field.default)
            Grid.SetColumn(box, 0)
            browse = WpfButton()
            browse.Content = "Browse..."
            browse.Height = 26
            browse.Width = 86
            browse.Margin = Thickness(6, 0, 0, 0)
            style = self._style("CED.Button.Base")
            if style is not None:
                browse.Style = style
            Grid.SetColumn(browse, 1)

            def _browse(sender, args, field=field, box=box):
                self._pick_path(field, box)

            browse.Click += _browse
            grid.Children.Add(box)
            grid.Children.Add(browse)
            field.control = box
            field._extra = browse
            return grid

        return None

    @staticmethod
    def _auto_length():
        from System.Windows import GridLength, GridUnitType

        return GridLength(1, GridUnitType.Auto)

    def _default_index(self, field):
        if not field.options:
            return -1
        if field.default is not None:
            for index, (value, _label) in enumerate(field.options):
                if value == field.default:
                    return index
        return 0

    def _pick_path(self, field, box):
        kwargs = {"title": str(field.label or "Select a file")}
        if field.file_filter:
            kwargs["files_filter"] = field.file_filter
        elif field.file_ext:
            kwargs["file_ext"] = str(field.file_ext).lstrip("*.")
        try:
            if field.kind == KIND_OPEN_FILE:
                picked = forms.pick_file(**kwargs)
            else:
                current = str(box.Text or "").strip()
                if current:
                    kwargs["default_name"] = os.path.basename(current)
                picked = forms.save_file(**kwargs)
        except Exception as exc:
            _LOGGER.warning("File picker failed: %s", exc)
            picked = None
        if picked:
            box.Text = str(picked)

    # -- values and validation --------------------------------------------

    def _choice_changed(self, sender, args):
        if getattr(self, "_suppress_dependents", False):
            return
        self._refresh_dependent_options()
        self._apply_visibility()

    def _refresh_dependent_options(self):
        """Rebuild any dropdown whose options depend on another field.

        Keeps the current pick when it survives the new list, so changing the
        source field does not silently reset a choice the user already made.
        """
        self._suppress_dependents = True
        try:
            for field in self._fields:
                rule = getattr(field, "options_from", None)
                if not rule or field.control is None:
                    continue
                try:
                    source_key, provider = rule[0], rule[1]
                except Exception:
                    continue
                source_value = None
                for other in self._fields:
                    if other.key == source_key:
                        source_value = self._value_of(other)
                        break
                try:
                    options = _normalize_options(provider(source_value))
                except Exception as exc:
                    _LOGGER.debug("Options provider failed for %s: %s", field.key, exc)
                    options = []
                previous = self._value_of(field)
                field.options = options
                combo = field.control
                combo.Items.Clear()
                for value, label in options:
                    item = ComboBoxItem()
                    item.Content = label
                    item.Tag = value
                    combo.Items.Add(item)
                index = -1
                for position, (value, _label) in enumerate(options):
                    if value == previous:
                        index = position
                        break
                if index < 0 and options:
                    index = self._default_index(field)
                combo.SelectedIndex = index
        finally:
            self._suppress_dependents = False

    def _is_visible(self, field):
        rule = field.visible_when
        if not rule:
            return True
        try:
            other_key, wanted = rule[0], rule[1]
        except Exception:
            return True
        if not isinstance(wanted, (list, tuple, set)):
            wanted = [wanted]
        for other in self._fields:
            if other.key == other_key:
                return self._value_of(other) in wanted
        return True

    def _apply_visibility(self):
        """Hide fields whose visible_when rule is not satisfied."""
        for field in self._fields:
            if field.container is None:
                continue
            try:
                field.container.Visibility = (
                    Visibility.Visible if self._is_visible(field)
                    else Visibility.Collapsed)
            except Exception:
                pass

    def _value_of(self, field):
        control = field.control
        kind = field.kind
        if control is None:
            return None
        if kind == KIND_CHOICE:
            item = control.SelectedItem
            return None if item is None else item.Tag
        if kind == KIND_MULTICHOICE:
            return [b.Tag for b in (field._extra or []) if bool(b.IsChecked)]
        if kind == KIND_BOOL:
            return bool(control.IsChecked)
        return control.Text

    def _problems(self):
        found = []
        for field in self._fields:
            if field.kind == KIND_INFO or not self._is_visible(field):
                continue
            value = self._value_of(field)

            if field.kind in (KIND_INT, KIND_FLOAT):
                raw = str(value or "").strip()
                if not raw:
                    if field.required:
                        found.append("%s is required." % field.label)
                    continue
                try:
                    parsed = float(raw) if field.kind == KIND_FLOAT else int(float(raw))
                except Exception:
                    found.append("%s must be a number." % field.label)
                    continue
                if field.minimum is not None and parsed < field.minimum:
                    found.append("%s must be %s or more." % (field.label, field.minimum))
                if field.maximum is not None and parsed > field.maximum:
                    found.append("%s must be %s or less." % (field.label, field.maximum))
                continue

            if not field.required:
                continue
            if field.kind == KIND_MULTICHOICE:
                if not value:
                    found.append("Select at least one %s." % field.label.lower())
            elif field.kind == KIND_CHOICE:
                if value is None:
                    found.append("%s is required." % field.label)
            elif field.kind == KIND_BOOL:
                pass
            else:
                if not str(value or "").strip():
                    found.append("%s is required." % field.label)
        return found

    def _show_problems(self, problems):
        if self._validation_host is None or self._validation_text is None:
            return
        if not problems:
            self._validation_host.Visibility = Visibility.Collapsed
            return
        self._validation_text.Text = "\n".join(problems)
        self._validation_host.Visibility = Visibility.Visible

    def collect(self):
        """Typed values for every field. Assumes validation passed."""
        values = {}
        for field in self._fields:
            if field.kind == KIND_INFO or not field.key:
                continue
            if not self._is_visible(field):
                values[field.key] = None
                continue
            value = self._value_of(field)
            if field.kind == KIND_INT:
                raw = str(value or "").strip()
                values[field.key] = int(float(raw)) if raw else None
            elif field.kind == KIND_FLOAT:
                raw = str(value or "").strip()
                values[field.key] = float(raw) if raw else None
            elif field.kind in (KIND_TEXT, KIND_MULTILINE, KIND_OPEN_FILE, KIND_SAVE_FILE):
                values[field.key] = str(value or "")
            else:
                values[field.key] = value
        return values

    # -- handlers ----------------------------------------------------------

    def _ok_clicked(self, sender, args):
        problems = self._problems()
        if problems:
            self._show_problems(problems)
            return
        values = self.collect()
        # Cross-field rules the per-field checks cannot see, e.g. "this tag type
        # does not suit that equipment category".
        if self._validate is not None:
            try:
                extra = self._validate(values) or []
            except Exception as exc:
                _LOGGER.debug("Settings form validate callback failed: %s", exc)
                extra = []
            if isinstance(extra, str):
                extra = [extra]
            if extra:
                self._show_problems(list(extra))
                return
        self._show_problems([])
        self.result = values
        self.Close()

    def _cancel_clicked(self, sender, args):
        self.result = None
        self.Close()

    def _reset_clicked(self, sender, args):
        for field in self._fields:
            control = field.control
            if control is None:
                continue
            if field.kind == KIND_CHOICE:
                control.SelectedIndex = self._default_index(field)
            elif field.kind == KIND_MULTICHOICE:
                selected = field.default
                for box in (field._extra or []):
                    box.IsChecked = True if selected is None else (box.Tag in selected)
            elif field.kind == KIND_BOOL:
                control.IsChecked = bool(field.default)
            else:
                control.Text = "" if field.default is None else str(field.default)
        self._show_problems([])


def prompt(title, fields, subtitle=None, ok_label="Continue", caller_file=None,
           validate=None):
    """Show the form and return {key: value}, or None if cancelled.

    ``validate`` is an optional callable taking the collected values and
    returning a problem string, a list of them, or nothing.
    """
    try:
        window = SettingsFormWindow(
            title, fields, subtitle=subtitle, ok_label=ok_label,
            caller_file=caller_file, validate=validate)
    except Exception as exc:
        _LOGGER.warning("Settings form could not be built: %s", exc)
        raise
    window.ShowDialog()
    return window.result
