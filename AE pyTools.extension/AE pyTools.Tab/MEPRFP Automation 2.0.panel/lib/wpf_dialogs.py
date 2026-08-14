# -*- coding: utf-8 -*-
"""
Lightweight WPF-based prompts: single-line text input and list picker.

Separated from ``forms_compat`` so a module-load failure in the WPF
stack only affects scripts that genuinely need WPF (the Stage 1
authoring tools), not the Stage 3 Import/Export scripts that only
touch Windows.Forms.

API:
    prompt_for_string(prompt, title="", default="")  ->  str | None
    pick_from_list(options, title="Pick one", prompt="Choose:",
                   display_func=None)               ->  T   | None
    multi_select_from_list(options, ...)             ->  [T] | None
    assign_choices(rows, choices, ...)               ->  [(row, choice)] | None
"""

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.IO import StringReader  # noqa: E402
from System.Windows import (  # noqa: E402
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    TextWrapping,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (  # noqa: E402
    Button,
    CheckBox,
    ColumnDefinition,
    Grid,
    TextBlock,
)
from System.Windows.Markup import XamlReader  # noqa: E402
from System.Xml import XmlReader  # noqa: E402


def _load_xaml(text):
    return XamlReader.Load(XmlReader.Create(StringReader(text)))


_STRING_PROMPT_XAML = """\
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="460" Height="170"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="PromptText" TextWrapping="Wrap" Margin="0,0,0,8"/>
    <TextBox  Grid.Row="1" x:Name="InputBox" Margin="0,0,0,8"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="OkButton" Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"""

_LIST_PICKER_XAML = """\
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="520" Height="500"
        WindowStartupLocation="CenterScreen">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="PromptText" Margin="0,0,0,8" TextWrapping="Wrap"/>
    <TextBox   Grid.Row="1" x:Name="SearchBox" Margin="0,0,0,8" Padding="2"
               ToolTip="Filter the list (case-insensitive substring match)"/>
    <ListBox   Grid.Row="2" x:Name="Options"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
      <Button x:Name="OkButton" Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"""


_MULTI_SELECT_XAML = """\
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="640" Height="540"
        WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <Style x:Key="CheckListBoxItem" TargetType="ListBoxItem">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border Background="Transparent" Padding="2">
              <CheckBox IsChecked="{Binding IsSelected,
                                            RelativeSource={RelativeSource TemplatedParent},
                                            Mode=TwoWay}"
                        VerticalContentAlignment="Center">
                <ContentPresenter VerticalAlignment="Center"/>
              </CheckBox>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="PromptText" Margin="0,0,0,8" TextWrapping="Wrap"/>
    <ListBox   Grid.Row="1" x:Name="Options"
               SelectionMode="Multiple"
               ItemContainerStyle="{StaticResource CheckListBoxItem}"/>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
      <Button x:Name="CheckAllButton" Content="Check all" Width="100" Margin="0,0,8,0"/>
      <Button x:Name="UncheckAllButton" Content="Uncheck all" Width="100" Margin="0,0,16,0"/>
      <Button x:Name="OkButton" Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"""


_ASSIGN_CHOICES_XAML = """\
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="900" Height="560"
        WindowStartupLocation="CenterScreen">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" x:Name="PromptText" Margin="0,0,0,8" TextWrapping="Wrap"/>
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
      <StackPanel x:Name="RowsPanel"/>
    </ScrollViewer>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
      <Button x:Name="OkButton" Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
"""

_PICK_PLACEHOLDER = "(pick a profile…)"


class _AssignRow(object):
    """Per-row UI state for ``_AssignChoicesDialog``."""

    __slots__ = ("row", "checkbox", "button", "choice_index")

    def __init__(self, row, checkbox, button, choice_index):
        self.row = row
        self.checkbox = checkbox
        self.button = button
        self.choice_index = choice_index


class _AssignChoicesDialog(object):
    """One checkbox + choice-button line per row. The button opens the
    searchable ``_ListPickerDialog`` over ``choices``, so a long choice
    list stays navigable. Picking a choice auto-checks the row."""

    def __init__(self, rows, choices, title, prompt,
                 row_display_func, choice_display_func,
                 preselected, picker_title, picker_prompt):
        self.window = _load_xaml(_ASSIGN_CHOICES_XAML)
        self.window.Title = title or ""
        self.window.FindName("PromptText").Text = prompt or ""
        self._rows_panel = self.window.FindName("RowsPanel")
        self._choices = list(choices or [])
        self._choice_display_func = choice_display_func
        self._picker_title = picker_title
        self._picker_prompt = picker_prompt
        self._row_display_func = row_display_func
        # Retained handlers (pythonnet GC defence).
        self._handlers = []
        self._row_states = []
        preselected = preselected or {}
        for i, row in enumerate(rows or []):
            self._add_row(row, preselected.get(i))
        self._h_ok = lambda s, e: self._on_ok(s, e)
        self._h_cancel = lambda s, e: self._on_cancel(s, e)
        self.window.FindName("OkButton").Click += self._h_ok
        self.window.FindName("CancelButton").Click += self._h_cancel
        self._result = None

    def _choice_label(self, choice_index):
        if choice_index is None:
            return _PICK_PLACEHOLDER
        choice = self._choices[choice_index]
        if self._choice_display_func:
            return self._choice_display_func(choice)
        return str(choice)

    def _add_row(self, row, choice_index):
        grid = Grid()
        for width in (28, None, None):
            col = ColumnDefinition()
            if width is not None:
                col.Width = GridLength(width)
            else:
                col.Width = GridLength(1.0, GridUnitType.Star)
            grid.ColumnDefinitions.Add(col)

        checkbox = CheckBox()
        checkbox.IsChecked = choice_index is not None
        checkbox.Margin = Thickness(4, 2, 0, 2)
        checkbox.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(checkbox, 0)
        grid.Children.Add(checkbox)

        label = TextBlock()
        label.Text = (
            self._row_display_func(row) if self._row_display_func else str(row)
        )
        label.Margin = Thickness(0, 4, 8, 4)
        label.VerticalAlignment = VerticalAlignment.Center
        label.TextWrapping = TextWrapping.Wrap
        Grid.SetColumn(label, 1)
        grid.Children.Add(label)

        button = Button()
        button.Content = self._choice_label(choice_index)
        button.Margin = Thickness(0, 2, 4, 2)
        button.Padding = Thickness(6, 2, 6, 2)
        button.HorizontalContentAlignment = HorizontalAlignment.Left
        Grid.SetColumn(button, 2)
        grid.Children.Add(button)

        state = _AssignRow(row, checkbox, button, choice_index)
        handler = self._make_pick_handler(state)
        self._handlers.append(handler)
        button.Click += handler

        self._rows_panel.Children.Add(grid)
        self._row_states.append(state)

    def _make_pick_handler(self, state):
        def _on_pick(sender, e):
            chosen = pick_from_list(
                list(range(len(self._choices))),
                title=self._picker_title,
                prompt=self._picker_prompt,
                display_func=self._choice_label,
            )
            if chosen is None:
                return
            state.choice_index = chosen
            state.button.Content = self._choice_label(chosen)
            state.checkbox.IsChecked = True
        return _on_pick

    def _on_ok(self, sender, e):
        chosen = []
        for state in self._row_states:
            if state.checkbox.IsChecked and state.choice_index is not None:
                chosen.append((state.row, self._choices[state.choice_index]))
        self._result = chosen
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._result = None
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self._result


def assign_choices(rows, choices, title="Assign", prompt="",
                   row_display_func=None, choice_display_func=None,
                   preselected=None, picker_title="Pick one",
                   picker_prompt="Choose:"):
    """Row-by-row assignment: each ``rows`` entry gets a checkbox and a
    button that opens the searchable list picker over ``choices``.

    ``preselected`` maps row index -> choice index; preselected rows
    start checked, the rest start unchecked with no choice. Picking a
    choice on any row auto-checks it.

    Returns ``[(row, choice), ...]`` for every checked row that has a
    choice (possibly empty), or ``None`` if cancelled.
    """
    return _AssignChoicesDialog(
        rows, choices, title, prompt,
        row_display_func, choice_display_func,
        preselected, picker_title, picker_prompt,
    ).show()


class _StringPromptDialog(object):
    def __init__(self, prompt, title, default):
        self.window = _load_xaml(_STRING_PROMPT_XAML)
        self.window.Title = title or ""
        self.window.FindName("PromptText").Text = prompt or ""
        self._input = self.window.FindName("InputBox")
        self._input.Text = default or ""
        self._input.SelectAll()
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("CancelButton").Click += self._on_cancel
        self._result = None

    def _on_ok(self, sender, e):
        self._result = self._input.Text or ""
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._result = None
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self._result


class _ListPickerDialog(object):
    def __init__(self, options, title, prompt, display_func):
        self.window = _load_xaml(_LIST_PICKER_XAML)
        self.window.Title = title or ""
        self.window.FindName("PromptText").Text = prompt or ""
        self._listbox = self.window.FindName("Options")
        self._search = self.window.FindName("SearchBox")
        # ``_pairs`` is the canonical list (label, option) sorted in
        # caller-supplied order. ``_visible_options`` mirrors the labels
        # currently shown in the listbox 1:1 so SelectedIndex maps back
        # to the right object even after filtering.
        self._pairs = []
        for opt in (options or []):
            label = display_func(opt) if display_func else str(opt)
            self._pairs.append((label, opt))
        self._visible_options = []
        self._render("")
        self.window.FindName("OkButton").Click += self._on_ok
        self.window.FindName("CancelButton").Click += self._on_cancel
        self._search.TextChanged += self._on_search_changed
        # Focus the search box so the user can start typing immediately.
        try:
            self._search.Focus()
        except Exception:
            pass
        self._result = None

    def _render(self, needle):
        n = (needle or "").strip().lower()
        self._listbox.Items.Clear()
        self._visible_options = []
        for label, opt in self._pairs:
            if n and n not in label.lower():
                continue
            self._listbox.Items.Add(label)
            self._visible_options.append(opt)

    def _on_search_changed(self, sender, e):
        self._render(self._search.Text)

    def _on_ok(self, sender, e):
        idx = self._listbox.SelectedIndex
        if 0 <= idx < len(self._visible_options):
            self._result = self._visible_options[idx]
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._result = None
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
        return self._result


def prompt_for_string(prompt, title="", default=""):
    return _StringPromptDialog(prompt, title, default).show()


def pick_from_list(options, title="Pick one", prompt="Choose:", display_func=None):
    return _ListPickerDialog(options, title, prompt, display_func).show()


def multi_select_from_list(options, title="Pick options", prompt="Check items:",
                           display_func=None):
    """Multi-select picker with checkboxes. Returns the list of chosen
    objects (in the order they were checked), or ``None`` if cancelled.
    Uses the same checkbox-templated ListBox pattern as the placement
    filters so click-anywhere-on-row toggles selection.
    """
    return _MultiSelectDialog(options, title, prompt, display_func).show()


class _MultiSelectDialog(object):
    def __init__(self, options, title, prompt, display_func):
        self.window = _load_xaml(_MULTI_SELECT_XAML)
        self.window.Title = title or ""
        self.window.FindName("PromptText").Text = prompt or ""
        self._listbox = self.window.FindName("Options")
        self._options = list(options or [])
        for opt in self._options:
            label = display_func(opt) if display_func else str(opt)
            self._listbox.Items.Add(label)
        # Retained handlers (pythonnet GC defence).
        self._h_ok = lambda s, e: self._on_ok(s, e)
        self._h_cancel = lambda s, e: self._on_cancel(s, e)
        self._h_check_all = lambda s, e: self._on_check_all(s, e)
        self._h_uncheck_all = lambda s, e: self._on_uncheck_all(s, e)
        self.window.FindName("OkButton").Click += self._h_ok
        self.window.FindName("CancelButton").Click += self._h_cancel
        self.window.FindName("CheckAllButton").Click += self._h_check_all
        self.window.FindName("UncheckAllButton").Click += self._h_uncheck_all
        self._result = None

    def _on_ok(self, sender, e):
        chosen = []
        # SelectedItems holds the *labels* (since we added strings).
        # Map back to options by label.
        selected_labels = list(self._listbox.SelectedItems)
        if selected_labels:
            label_to_option = {}
            for opt in self._options:
                label = str(opt) if not callable(getattr(opt, "to_label", None)) else opt.to_label()
                label_to_option.setdefault(label, opt)
            # Better: rebuild from the items list directly.
            for i in range(self._listbox.Items.Count):
                container = self._listbox.ItemContainerGenerator.ContainerFromIndex(i)
                if container is not None and getattr(container, "IsSelected", False):
                    chosen.append(self._options[i])
        self._result = chosen
        self.window.Close()

    def _on_cancel(self, sender, e):
        self._result = None
        self.window.Close()

    def _on_check_all(self, sender, e):
        self._listbox.SelectAll()

    def _on_uncheck_all(self, sender, e):
        self._listbox.UnselectAll()

    def show(self):
        self.window.ShowDialog()
        return self._result
