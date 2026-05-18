# -*- coding: utf-8 -*-
"""
Raw WPF helper using pythonnet.

``pyrevit.forms`` is IronPython-only; this module loads XAML directly
through ``System.Windows.Markup.XamlReader`` so windows work in
CPython 3 + pythonnet (and IronPython 2.7 too, if we ever fall back).

Convention for stage 1: every UI is *modal*. Modeless windows in
Revit need ``ExternalEvent`` plumbing for any Revit-API call from an
event handler — we'll add that selectively later.
"""

import io
import os

import clr  # noqa: F401

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.IO import StringReader  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
from System.Windows.Media import VisualTreeHelper  # noqa: E402
from System.Xml import XmlReader  # noqa: E402


def load_xaml(xaml_path_or_text):
    """Load XAML from a file path or a string. Returns the root WPF object.

    The argument is treated as a file path if it points at an existing
    file, otherwise as inline XAML text.
    """
    if os.path.isfile(xaml_path_or_text):
        with io.open(xaml_path_or_text, "r", encoding="utf-8") as f:
            xaml_text = f.read()
    else:
        xaml_text = xaml_path_or_text
    reader = XmlReader.Create(StringReader(xaml_text))
    return XamlReader.Load(reader)


def attach_per_row_handlers(data_grid, on_button_click=None,
                            items_per_combo_name=None,
                            checkboxes_per_name=None,
                            combo_item_resolver=None):
    """Combined per-row wiring done in one Loaded handler:

      * Click handlers on per-row Buttons.
      * ItemsSource on per-row ComboBoxes, plus a binding-resync that
        manually pushes the bound source value through to the combo
        (works around pythonnet 3 not surfacing bound values when
        ItemsSource is set programmatically).
      * IsChecked init + Checked/Unchecked write-back on per-row
        CheckBoxes — same workaround.

    ``on_button_click(button, eventArgs, row_item)`` fires for any
    Button found in the row. ``items_per_combo_name`` maps
    ``combo_x_name`` -> shared ``ItemsSource`` collection.
    ``checkboxes_per_name`` maps ``checkbox_x_name`` -> the row
    property name to read/write (e.g. ``"Selected"`` for a
    ``_PreviewRow.Selected`` ``@property`` + ``@setter`` pair).
    Pass any of these as ``None`` to skip.

    ``combo_item_resolver(combo_name, row_item) -> CLR_list_or_None``
    enables *per-row, kind-aware* ItemsSource. When provided, the
    resolver is called for every combo in the row and its result
    overrides ``items_per_combo_name`` for that combo. Returning
    ``None`` falls back to the static map.

    The resolver fires twice:

      * Once at row-load, against the row's saved state, so the combo
        opens with the right list on first click.
      * Again on each ``DropDownOpened`` event — the user clicked
        the dropdown arrow; we re-resolve at that moment so changes
        the user just made to a controlling combo (e.g. Kind) are
        reflected before the popup materializes.

    ``SelectionChanged`` is intentionally NOT used as the live trigger
    even though it would feel more responsive: that subscription
    accumulates under DataGrid row recycling and re-enters mid-binding,
    which has crashed Revit + pythonnet under WPF. ``DropDownOpened``
    fires only when the user actively opens a dropdown — no feedback
    loop, no typing interference.

    Doing all three in one Loaded handler avoids ordering interference
    we hit when wiring them via separate handlers.
    """
    import clr  # noqa: F401
    from System import EventHandler
    from System.Windows import RoutedEventHandler
    from System.Windows.Controls import (
        Button as _Button,
        CheckBox as _CheckBox,
        ComboBox as _ComboBox,
    )
    from System.Windows.Data import BindingOperations

    items_per_combo_name = items_per_combo_name or {}
    checkboxes_per_name = checkboxes_per_name or {}

    # Single stable handler shared across all rows. Re-resolving on
    # DropDownOpened is safe — the user has clicked the arrow, we're
    # not mid-binding, mid-typing, or mid-selection. We always read
    # row state from ``sender.DataContext`` so the closure doesn't
    # capture per-row data (no stale references under DataGrid row
    # recycling). Detach + attach in the row-load pass keeps
    # subscriptions deduped even if Loaded fires repeatedly on the
    # same combo control as virtualized rows are reused.
    def _on_dropdown_opened(sender, _e):
        if combo_item_resolver is None:
            return
        try:
            name = sender.Name or ""
            row_item = sender.DataContext
        except Exception:
            return
        try:
            source = combo_item_resolver(name, row_item)
        except Exception:
            return
        if source is None:
            return
        try:
            current = sender.ItemsSource
        except Exception:
            current = None
        if current is source:
            return
        try:
            sender.ItemsSource = source
        except Exception:
            pass

    _dropdown_handler = EventHandler(_on_dropdown_opened)

    # Properties whose bindings we re-target after ItemsSource is
    # set. Each cell's SelectedValue / Text / SelectedItem binding
    # was evaluated at row-template load time when ItemsSource was
    # still empty; without this nudge the combo stays blank even
    # when the bound source has a value.
    _resync_props = (
        _ComboBox.SelectedValueProperty,
        _ComboBox.SelectedItemProperty,
        _ComboBox.TextProperty,
    )

    def _resync_combo_bindings(combo):
        # Two passes:
        #   1. Ask each binding to UpdateTarget() — re-read the source
        #      property and try to resolve to an item. Works when WPF
        #      can do the item-vs-value equality match itself.
        #   2. As a backup (because UpdateTarget has been observed to
        #      no-op under pythonnet 3 even when items are present),
        #      read the bound source value directly via the binding's
        #      Path.Path, then push it through to the target property.
        #      For SelectedValue we also fall back to finding a matching
        #      item by stringified equality and setting SelectedItem
        #      directly — the truly bulletproof path.
        for prop in _resync_props:
            try:
                expr = BindingOperations.GetBindingExpression(combo, prop)
            except Exception:
                expr = None
            if expr is None:
                continue
            try:
                expr.UpdateTarget()
            except Exception:
                pass

            # Manual round-trip: read source, push to target.
            try:
                path_obj = getattr(expr.ParentBinding, "Path", None)
                path = getattr(path_obj, "Path", None) if path_obj else None
                ctx = combo.DataContext
            except Exception:
                path = None
                ctx = None
            if not path or ctx is None:
                continue
            try:
                source_val = getattr(ctx, path, None)
            except Exception:
                source_val = None
            if source_val is None:
                continue
            try:
                combo.SetValue(prop, source_val)
            except Exception:
                pass

            # Last-ditch for SelectedValue / SelectedItem: stringify
            # the source value and find a matching item in ItemsSource
            # by string equality. Handles the case where pythonnet
            # wraps Python str differently from the CLR strings stored
            # in our ObservableCollection.
            if prop in (
                _ComboBox.SelectedValueProperty,
                _ComboBox.SelectedItemProperty,
            ):
                target_str = str(source_val).strip()
                if not target_str:
                    continue
                items = combo.ItemsSource
                if items is None:
                    continue
                for item in items:
                    try:
                        if str(item).strip() == target_str:
                            combo.SelectedItem = item
                            break
                    except Exception:
                        continue

    def _walk_row(row):
        """Return ``(buttons, combos, checkboxes)`` found via BFS
        through the row's visual tree. Each control type is a leaf
        — we don't descend into its internal template."""
        buttons = []
        combos = []
        checks = []
        queue = [row]
        while queue:
            node = queue.pop(0)
            # Order matters: CheckBox extends ToggleButton extends
            # ButtonBase (NOT Button), so its isinstance check is safe
            # against the Button check above. ComboBox also doesn't
            # match Button.
            if isinstance(node, _CheckBox):
                checks.append(node)
                continue
            if isinstance(node, _Button):
                buttons.append(node)
                continue
            if isinstance(node, _ComboBox):
                combos.append(node)
                continue
            try:
                count = VisualTreeHelper.GetChildrenCount(node)
            except Exception:
                count = 0
            for i in range(count):
                try:
                    queue.append(VisualTreeHelper.GetChild(node, i))
                except Exception:
                    pass
        return buttons, combos, checks

    def _on_button_click(sender, e):
        row_item = getattr(sender, "DataContext", None)
        if on_button_click is not None:
            on_button_click(sender, e, row_item)

    button_handler = RoutedEventHandler(_on_button_click)

    # Per-checkbox handler factory — closes over the property name
    # from the caller's checkboxes_per_name map. Holds strong refs to
    # the RoutedEventHandlers in the returned dict so pythonnet
    # doesn't GC them.
    _checkbox_handlers = []

    def _wire_checkbox(chk, prop_name):
        def _on_state_change(s, e):
            ctx = s.DataContext
            if ctx is None:
                return
            try:
                setattr(ctx, prop_name, bool(s.IsChecked))
            except Exception:
                pass
        h = RoutedEventHandler(_on_state_change)
        _checkbox_handlers.append(h)
        try:
            chk.Checked -= h
        except Exception:
            pass
        try:
            chk.Unchecked -= h
        except Exception:
            pass
        try:
            chk.Checked += h
            chk.Unchecked += h
        except Exception:
            pass

    def _on_row_loaded(sender, e):
        row = sender
        buttons, combos, checkboxes = _walk_row(row)

        # Wire ComboBox ItemsSource FIRST so any side-effect of
        # ItemsSource assignment (template realisation etc.) happens
        # before we attach Click handlers — that way Click attachment
        # lands on the final, stable Button visuals.
        #
        # Two-pass model:
        #   1. ROW LOAD — resolve once against the row's saved state
        #      so the combo opens with the right list on first click.
        #   2. DROPDOWN OPEN — re-resolve the moment the user clicks
        #      the dropdown arrow, picking up any in-dialog edits to
        #      controlling combos (e.g. Kind) before the popup
        #      materializes. We use ``DropDownOpened`` rather than
        #      SelectionChanged because the latter accumulates under
        #      DataGrid row recycling and re-enters mid-binding —
        #      that combination crashed Revit + pythonnet repeatedly.
        row_item = getattr(row, "DataContext", None)
        for combo in combos:
            try:
                name = combo.Name or ""
            except Exception:
                name = ""
            source = None
            if combo_item_resolver is not None:
                try:
                    source = combo_item_resolver(name, row_item)
                except Exception:
                    source = None
            if source is None and items_per_combo_name:
                source = items_per_combo_name.get(name)
            if source is not None:
                try:
                    combo.ItemsSource = source
                except Exception:
                    pass
                # Re-evaluate any bindings on the combo (Selected*, Text).
                # The XAML binding fired at row-template-load time when
                # ItemsSource was still empty, so the saved value couldn't
                # resolve to an item; now that items are present, push
                # the source value back through.
                _resync_combo_bindings(combo)

            # Subscribe DropDownOpened for live re-resolution on user
            # action. Detach-then-attach so virtualized row reuse
            # doesn't accumulate duplicate subscriptions on the same
            # control. Only meaningful when a resolver is in play.
            if combo_item_resolver is not None:
                try:
                    combo.DropDownOpened -= _dropdown_handler
                except Exception:
                    pass
                try:
                    combo.DropDownOpened += _dropdown_handler
                except Exception:
                    pass

        if on_button_click is not None:
            for btn in buttons:
                try:
                    btn.Click -= button_handler
                except Exception:
                    pass
                try:
                    btn.Click += button_handler
                except Exception:
                    pass

        if checkboxes_per_name:
            for chk in checkboxes:
                try:
                    name = chk.Name or ""
                except Exception:
                    name = ""
                prop_name = checkboxes_per_name.get(name)
                if not prop_name:
                    continue
                ctx = chk.DataContext
                if ctx is None:
                    continue
                # Push the row property's current value into IsChecked
                # so the checkbox visually reflects the saved/default
                # state. Avoids the pythonnet binding-doesn't-display
                # quirk we get with DataGridCheckBoxColumn.
                try:
                    current = getattr(ctx, prop_name)
                except Exception:
                    current = False
                try:
                    chk.IsChecked = bool(current)
                except Exception:
                    pass
                _wire_checkbox(chk, prop_name)

    row_loaded_handler = RoutedEventHandler(_on_row_loaded)

    def _on_loading_row(sender, e):
        row = e.Row
        try:
            row.Loaded -= row_loaded_handler
        except Exception:
            pass
        try:
            row.Loaded += row_loaded_handler
        except Exception:
            pass

    data_grid.LoadingRow += _on_loading_row

    return {
        "loading_row": _on_loading_row,
        "row_loaded": row_loaded_handler,
        "button_click": button_handler,
        "checkbox_handlers": _checkbox_handlers,
    }


def attach_per_row_combo_items_source(data_grid, items_per_name):
    """Wire ``ItemsSource`` onto every per-row ``ComboBox`` in a DataGrid.

    Workaround for an unreliable XAML binding under pythonnet 3:
    ``ItemsSource="{Binding RowProperty}"`` where ``RowProperty`` is a
    Python attribute returning a CLR collection has been observed to
    surface as empty even when direct ``combo.ItemsSource = collection``
    works. This helper sidesteps the binding entirely by walking each
    row's visual tree at row-load time and assigning ItemsSource
    directly.

    ``items_per_name`` is a dict ``{combo_x_name: items_source}`` where
    ``items_source`` is a CLR ``IEnumerable`` (typically an
    ``ObservableCollection``) shared across all rows. The ComboBox's
    ``x:Name`` in the cell template selects which collection to use.

    Like ``attach_per_row_button_handler``, returns a dict of strong
    refs the caller stores on the controller so pythonnet doesn't GC
    them.
    """
    import clr  # noqa: F401
    from System.Windows import RoutedEventHandler
    from System.Windows.Controls import ComboBox as _ComboBox

    def _row_combos(row):
        out = []
        queue = [row]
        while queue:
            node = queue.pop(0)
            if isinstance(node, _ComboBox):
                out.append(node)
                # don't descend into the combo's own visual tree
                continue
            try:
                count = VisualTreeHelper.GetChildrenCount(node)
            except Exception:
                count = 0
            for i in range(count):
                try:
                    queue.append(VisualTreeHelper.GetChild(node, i))
                except Exception:
                    pass
        return out

    def _on_row_loaded(sender, e):
        row = sender
        for combo in _row_combos(row):
            try:
                name = combo.Name or ""
            except Exception:
                name = ""
            source = items_per_name.get(name)
            if source is None:
                continue
            # Re-set every time the row gets re-laid-out; cheap, idempotent.
            try:
                combo.ItemsSource = source
            except Exception:
                pass

    row_loaded_handler = RoutedEventHandler(_on_row_loaded)

    def _on_loading_row(sender, e):
        row = e.Row
        try:
            row.Loaded -= row_loaded_handler
        except Exception:
            pass
        try:
            row.Loaded += row_loaded_handler
        except Exception:
            pass

    data_grid.LoadingRow += _on_loading_row

    return {
        "loading_row": _on_loading_row,
        "row_loaded": row_loaded_handler,
    }


def attach_per_row_button_handler(data_grid, on_click):
    """Wire a Click handler onto every per-row Button in a DataGrid.

    Workaround for unreliable bubbled ``Button.ClickEvent`` handling
    inside DataGrid cells — the grid's selection/edit machinery
    sometimes intercepts the event before it bubbles to the window.

    Strategy: hook ``DataGrid.LoadingRow`` so we get a callback when
    each row materialises; at that point register on the row's
    ``Loaded`` so the visual tree is laid out; then walk the row's
    visual children, find every ``System.Windows.Controls.Button``,
    and attach a Click handler to each. The handler calls ``on_click``
    with ``(sender, eventArgs, row_item)`` where ``row_item`` is the
    row's DataContext (the Python row object).

    The returned dict carries strong references to the CLR delegates
    so pythonnet doesn't GC them. Caller stores it on the controller.
    """
    import clr  # noqa: F401
    from System.Windows import RoutedEventHandler
    from System.Windows.Controls import Button as _Button

    def _row_buttons(row):
        # BFS through the row's visual tree.
        out = []
        queue = [row]
        while queue:
            node = queue.pop(0)
            if isinstance(node, _Button):
                out.append(node)
                continue  # don't descend into a Button's children
            try:
                count = VisualTreeHelper.GetChildrenCount(node)
            except Exception:
                count = 0
            for i in range(count):
                try:
                    queue.append(VisualTreeHelper.GetChild(node, i))
                except Exception:
                    pass
        return out

    def _on_button_click(sender, e):
        # The row item the Button was rendered for is the Button's
        # DataContext — set automatically by DataGrid when the row
        # was laid out.
        row_item = getattr(sender, "DataContext", None)
        on_click(sender, e, row_item)

    button_handler = RoutedEventHandler(_on_button_click)

    def _on_row_loaded(sender, e):
        # ``sender`` is the DataGridRow.
        row = sender
        for btn in _row_buttons(row):
            # Detach first to avoid duplicate registrations when the
            # row gets re-laid-out (e.g. column resize, DataContext
            # swap, virtualization recycle).
            try:
                btn.Click -= button_handler
            except Exception:
                pass
            try:
                btn.Click += button_handler
            except Exception:
                pass

    row_loaded_handler = RoutedEventHandler(_on_row_loaded)

    def _on_loading_row(sender, e):
        # Subscribe to Loaded so we wait until the visual tree is
        # populated. Detach first in case Loaded was already wired
        # by a prior LoadingRow on the same row instance.
        row = e.Row
        try:
            row.Loaded -= row_loaded_handler
        except Exception:
            pass
        try:
            row.Loaded += row_loaded_handler
        except Exception:
            pass

    # ``DataGrid.LoadingRow`` uses ``EventHandler<DataGridRowEventArgs>``,
    # which pythonnet auto-coerces from a plain Python callable.
    data_grid.LoadingRow += _on_loading_row

    # Return the handles so the caller can hold strong refs on them.
    return {
        "loading_row": _on_loading_row,
        "row_loaded": row_loaded_handler,
        "button_click": button_handler,
    }


def resolve_row_from_event(e, row_cls):
    """Find the Python row object behind a bubbled WPF Click event.

    Used by per-row buttons inside a ``DataGridTemplateColumn``: their
    Tag is bound to the row item via ``Tag="{Binding}"``, but the
    bubbled click reaches the window with various visual children as
    the source. This walks the lookup chain robustly:

      1. ``e.Source.Tag`` then ``e.Source.DataContext`` — Button.Click
         normally fires with ``Source = the Button``.
      2. Walk up from ``e.OriginalSource`` via ``VisualTreeHelper``,
         checking ``Tag`` and ``DataContext`` at each level.

    Tolerates pythonnet's habit of round-tripping Python objects
    through CLR ``object`` properties in ways that break ``isinstance``
    — falls back to a class-name match (``__class__.__name__``) so the
    same row object is recognised even when its CLR wrapper has been
    re-boxed.

    Returns the row object or ``None``. The caller is responsible for
    deciding what an unmatched event means (typically: "ignore — this
    click came from Save / Refresh / Close / etc.").
    """
    target_name = getattr(row_cls, "__name__", None)

    def _is_row(obj):
        if obj is None:
            return False
        if isinstance(obj, row_cls):
            return True
        cls = getattr(obj, "__class__", None)
        cls_name = getattr(cls, "__name__", None) if cls else None
        return bool(target_name) and cls_name == target_name

    def _check(elem):
        if elem is None:
            return None
        for attr in ("Tag", "DataContext"):
            try:
                val = getattr(elem, attr, None)
            except Exception:
                val = None
            if _is_row(val):
                return val
        return None

    # Strategy 1: e.Source — for Button.Click this is normally the
    # Button itself (Button.OnClick raises with Source=this).
    found = _check(getattr(e, "Source", None))
    if found is not None:
        return found

    # Strategy 2: walk up from OriginalSource via VisualTreeHelper.
    candidate = getattr(e, "OriginalSource", None)
    for _ in range(20):  # bounded
        found = _check(candidate)
        if found is not None:
            return found
        if candidate is None:
            break
        try:
            parent = VisualTreeHelper.GetParent(candidate)
        except Exception:
            parent = None
        if parent is None:
            try:
                parent = getattr(candidate, "Parent", None)
            except Exception:
                parent = None
        if parent is None or parent is candidate:
            break
        candidate = parent

    return None


class WpfWindow(object):
    """Base class for stage 1 modal windows.

    Subclasses load a XAML file in ``__init__`` and bind events via
    ``self.find(name)``. Call ``show_modal()`` to display.
    """

    def __init__(self, xaml_path):
        self.window = load_xaml(xaml_path)
        self.result = None  # subclasses populate before closing

    def find(self, name):
        return self.window.FindName(name)

    def show_modal(self):
        self.window.ShowDialog()
        return self.result

    def close(self, result=None):
        if result is not None:
            self.result = result
        self.window.Close()
