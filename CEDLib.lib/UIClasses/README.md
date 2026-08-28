# UIClasses Framework

`UIClasses` now includes reusable foundations for new pyRevit UIs:

- `resource_loader.py`
  - Loads centralized dictionaries from `UIClasses/Resources`.
  - Applies the loader-defined theme descriptors and accent (`blue`, `red`, `green`, `neutral`).
- `ui_bases.py`
  - `CEDWindowBase`: base for modeless/modal windows.
  - `CEDPanelBase`: base for dockable panels.
- `filterable_combo_box.py`
  - `FilterableComboBox`: adds live, case-insensitive filtering to an editable
    combo while keeping the shared arrow/dropdown interaction.
- `Resources/Templates/WindowChrome.xaml`
  - Base styles for window/page chrome and section cards.
- `Resources/Templates/ControlPrimitives.xaml`
  - Common icon, separator, empty-state, and overlay primitives.
- `Resources/Styles/ListStyles.xaml` and `Resources/Templates/DataGrids.xaml`
  - Data grids default to the existing flat row treatment for compatibility.
  - Use `CED.DataGrid.Base.Alternating` or `CED.DataGrid.Display.Alternating` to
    enable two-row alternation, and the corresponding `.Flat` key to explicitly
    disable it.
  - Use the `.Alternating` row/cell variants when a grid or column supplies an
    opaque custom background; they preserve selection and semantic state colors.
  - DataGrid row surfaces and selection use dedicated tokens so list-item colors
    do not compete with grid borders, read-only cells, or selected rows.

## Minimal usage

```python
from UIClasses.ui_bases import CEDWindowBase

class MyWindow(CEDWindowBase):
    theme_aware = True
    auto_wire_textboxes = True
    text_select_all_on_click = True
    text_select_all_on_focus = True

    def __init__(self):
        CEDWindowBase.__init__(
            self,
            xaml_source="MyWindow.xaml",  # optional when class-name xaml exists
        )
```

### What is automatic now

- Resolves module-relative XAML path (e.g. `"MyWindow.xaml"`).
- Infers XAML automatically (`<ClassName>.xaml` / `<ModuleName>.xaml`) if `xaml_source` is omitted.
- Resolves workspace/lib/resources paths and appends `CEDLib.lib` to `sys.path`.
- Loads theme/accent from `AE-pyTools-Theme` config when `theme_aware = True`.
- Falls back to Light/Blue when `theme_aware = False`.
- Optional shift+mousewheel horizontal scroll handling.
- Optional textbox select-all behavior (click/focus) via class flags.

For existing tools, migration can stay incremental: move to these base classes first, then remove duplicated local path/theme/input wiring per tool.

## Filterable combo boxes

Use the shared editable combo style with the behavior helper:

```python
from UIClasses import FilterableComboBox

combo.ItemsSource = ["Lighting - Pendant", "Power - Floor Box"]
filter_behavior = FilterableComboBox(combo)
```

Typing opens the dropdown and filters its item view as text changes. Clicking the
arrow opens the current filtered view; selecting an item clears the filter so
the next open starts with all options. Up/Down highlights visible matches while
preserving the query, Enter commits the highlighted match, and double-clicking
the editable text selects the full value.

By default, typed values are allowed and remain the final input when focus
leaves the control:

```python
filter_behavior = FilterableComboBox(combo, allow_custom_values=True)
```

Theme selectors are populated from `UIClasses/resource_loader.py`. Add a
`ThemeDescriptor` there, including its mode, label, description, and XAML
dictionary path, and the Theme picker, Circuit Settings, and Circuit Manager
selectors will include it automatically.

For list-only inputs, set `allow_custom_values=False`. When focus leaves with
text that is not an exact list value, the behavior can resolve to a blank value,
the last valid item, or a configured default item:

```python
from UIClasses import FALLBACK_DEFAULT_ITEM, FilterableComboBox

filter_behavior = FilterableComboBox(
    combo,
    allow_custom_values=False,
    fallback=FALLBACK_DEFAULT_ITEM,
    fallback_item="Lighting - Pendant",
)
```

Use `behavior.value`, `behavior.selected_item`, and
`behavior.has_valid_selection` from the consuming tool. An optional
`on_value_committed(behavior, value, item)` callback runs after an item or final
text value commits.

## Structured search

`structured_search.py` provides the reusable, presentation-independent query
and token state.  Hosts supply `SearchFilterDefinition` objects and own the
item matching rules.  `structured_search_box.py` provides the composited WPF
control for live free-text input, `/` command narrowing, filter chips, and
atomic token selection/deletion.

```python
from UIClasses import SearchFilterDefinition, StructuredSearchBox

filters = [
    SearchFilterDefinition(
        "status",
        "Status",
        matcher=lambda item, value: item.status.lower() == value.lower(),
    ),
]
search = StructuredSearchBox(filters, placeholder="Search items")
search.add_query_changed_handler(on_query_changed)
SearchHost.Children.Add(search)
```

`on_query_changed(sender, args)` receives `args.query`, whose `FreeText` and
meaningful `Filters` are ready for the host's collection predicate.  Empty
tokens and the in-progress `/` command never enter the effective query.  A
definition may set `allow_multiple=True` to append repeated values and
`combine_mode="or"` to OR those values for the same key while other keys stay
ANDed.  `value_hint` supplies accepted-value guidance in the command picker;
token style and text brush resource keys can provide host-specific chip colors.
