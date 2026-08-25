# -*- coding: utf-8 -*-
"""In-Revit smoke test for the native FlowDocument viewer and fixture.

Run through ``pyrevit run`` with ``CED_DOC_TEST_REPO`` set to the repository
root. The script is read-only and does not require an open Revit model.
"""

from __future__ import print_function

import io
import os
import sys
import time

repository_root = os.getenv("CED_DOC_TEST_REPO")
if not repository_root or not os.path.isdir(repository_root):
    raise AssertionError("CED_DOC_TEST_REPO must identify the repository root.")

lib_root = os.path.join(repository_root, "CEDLib.lib")
documentation_root = os.path.join(repository_root, "docs", "user-guide")
if lib_root not in sys.path:
    sys.path.insert(0, lib_root)

from Documentation.viewer import DocumentationViewerWindow
from System import Action
from System.Windows import Style, TextAlignment, Visibility
from System.Windows.Controls import StackPanel, TextBlock
from System.Windows.Documents import BlockUIContainer, Run, Section, Table
from System.Windows.Input import Key, ModifierKeys
from System.Windows.Shapes import Path
from System.Windows.Threading import DispatcherPriority
from UIClasses import resource_loader
from pyrevit import VERSION_STRING as PYREVIT_VERSION_STRING


fixture_path = os.path.join(documentation_root, "_fixtures", "renderer-compatibility.md")
with io.open(fixture_path, "r", encoding="utf-8-sig") as stream:
    fixture_text = stream.read()

startup_started = time.time()
window = DocumentationViewerWindow(
    documentation_root=documentation_root,
    uiapp=globals().get("__revit__"),
    defer_runtime=False,
)
startup_seconds = time.time() - startup_started
try:
    assert window.catalog is not None, "catalog did not load"
    assert window._results_refresh_count == 1, "startup rebuilt the results more than once"
    assert len(window.catalog.documents) >= 51, "catalog document count is unexpectedly low"
    assert window.error_overlay.Visibility == Visibility.Collapsed, "viewer opened with an error"
    assert resource_loader.try_find_resource(window, "CED.__MissingResourceSmoke") is None
    tree_style = resource_loader.try_find_resource(window, "CED.Documentation.TreeItem")
    assert isinstance(tree_style, Style), "tree item resource is not a Style"
    assert window.document_viewer.Document is not None, "home document did not render"
    assert resource_loader.try_find_resource(window, "CED.Brush.DocumentCanvas") is not None
    assert resource_loader.try_find_resource(window, "CED.Brush.Link") is not None
    assert resource_loader.try_find_resource(window, "CED.Brush.LinkHover") is not None
    assert resource_loader.try_find_resource(window, "CED.Brush.SearchHighlightBackground") is not None
    assert resource_loader.try_find_resource(window, "CED.Brush.SearchHighlightForeground") is not None
    assert resource_loader.try_find_resource(window, "CED.Icon.Home") is not None
    assert resource_loader.try_find_resource(window, "CED.Icon.Search") is not None
    assert resource_loader.try_find_resource(window, "CED.Icon.ChevronDown") is not None
    assert window._is_find_shortcut(Key.F, ModifierKeys.Control), "Ctrl+F was not routed to viewer search"
    assert not window._is_find_shortcut(Key.F, ModifierKeys(0)), "unmodified F was treated as find"
    resource_sources = [str(item.Source or "").lower() for item in window.Resources.MergedDictionaries]
    assert any("inputstyles.xaml" in value for value in resource_sources), (
        "shared input and combo scrollbar styles were not loaded"
    )
    assert any("liststyles.xaml" in value for value in resource_sources), (
        "shared list and viewer scrollbar styles were not loaded"
    )
    assert window.back_button.Height == 32
    assert window.home_button.Height == 32
    assert window.search_box.Height == 32
    assert window.extension_filter.Height == 32
    assert window.ribbon_filter.Height == 32
    window.tree_mode_button.IsChecked = True
    window._result_mode_changed(window.tree_mode_button, None)
    assert window.results_tree.Visibility == Visibility.Visible, "tree mode did not become visible"
    assert window.results_tree.Items.Count > 0, "tree mode did not populate"
    window.search_box.Text = "selection"
    assert window.clear_search_button.Visibility == Visibility.Visible, "search clear button is hidden"
    assert window.search_placeholder.Visibility == Visibility.Collapsed, "search placeholder did not hide"
    assert all(item.IsExpanded for item in window.results_tree.Items), "search tree did not expand"
    window._clear_search_clicked(window.clear_search_button, None)
    assert window.clear_search_button.Visibility == Visibility.Collapsed, "search clear did not hide"
    assert window.search_placeholder.Visibility == Visibility.Visible, "search placeholder did not return"
    window._set_tree_expanded(True)
    assert all(item.IsExpanded for item in window.results_tree.Items), "expand all failed"
    window._set_tree_expanded(False)
    assert not any(item.IsExpanded for item in window.results_tree.Items), "collapse all failed"
    electrical_header = window._tree_path_items.get("ced-electools/index.md")
    assert electrical_header is not None and electrical_header.HasItems, "extension index header missing"
    electrical_header.IsSelected = True
    assert window.current_relative_path == "ced-electools/index.md", "header index did not open"
    window.list_mode_button.IsChecked = True
    window._result_mode_changed(window.list_mode_button, None)
    assert window.results_list.Visibility == Visibility.Visible, "list mode did not become visible"
    rendered = window.renderer.render(fixture_text, fixture_path)
    assert rendered is not None, "fixture did not produce a FlowDocument"
    assert rendered.Blocks.Count >= 15, "fixture produced too few WPF blocks"
    assert "target-heading" in window.renderer.heading_blocks, "fixture heading anchor is missing"
    rendered_tables = [block for block in rendered.Blocks if isinstance(block, Table)]
    assert rendered_tables, "fixture table did not render"
    first_cell_paragraph = rendered_tables[0].RowGroups[0].Rows[0].Cells[0].Blocks.FirstBlock
    assert first_cell_paragraph.TextAlignment == TextAlignment.Left, "table text is not left aligned"
    assert any(isinstance(block, Section) for block in rendered.Blocks), (
        "lists are not rendered as selectable FlowDocument text"
    )
    callout_titles = []
    for block in rendered.Blocks:
        if not isinstance(block, BlockUIContainer):
            continue
        border = block.Child
        panel = getattr(border, "Child", None)
        if not isinstance(panel, StackPanel) or panel.Children.Count < 2:
            continue
        header = panel.Children[0]
        if not isinstance(header, StackPanel) or header.Children.Count < 2:
            continue
        if isinstance(header.Children[0], Path) and isinstance(header.Children[1], TextBlock):
            callout_titles.append(str(header.Children[1].Text))
    assert set(callout_titles) == set(("Note", "Tip", "Important", "Warning", "Caution")), (
        "callout type glyphs or labels are incomplete"
    )
    assert window.catalog.search("bounding area"), "full-text catalog search failed"
    title_results = window.catalog.search("Zoom to Selection")
    assert title_results[0]["id"] == "ae-pytools-zoom-to-selection", "exact title did not rank first"
    assert window._open_document(
        "ae-pytools/zoom-to-selection.md", push_history=True
    ), "catalog page did not open"
    window.search_box.Text = "Zoom to Selection"
    window._apply_page_search_highlights()
    highlighted_page_runs = []
    for block in window.document_viewer.Document.Blocks:
        inlines = getattr(block, "Inlines", None)
        if inlines is None:
            continue
        for inline in inlines:
            if isinstance(inline, Run) and inline.Background is not None:
                highlighted_page_runs.append(str(inline.Text))
    assert "Zoom" in highlighted_page_runs and "Selection" in highlighted_page_runs, (
        "page search matches were not highlighted"
    )
    matching_row = next(
        item for item in window.results_list.Items
        if str(item["DocumentId"]) == "ae-pytools-zoom-to-selection"
    )
    title_control = matching_row["TitleDisplay"]
    highlighted_title_runs = [
        str(inline.Text) for inline in title_control.Inlines
        if isinstance(inline, Run) and inline.Background is not None
    ]
    assert highlighted_title_runs == ["Zoom", "Selection"], "list title matches were not highlighted"
    window.tree_mode_button.IsChecked = True
    window._result_mode_changed(window.tree_mode_button, None)
    matching_tree_item = window._tree_path_items.get("ae-pytools/zoom-to-selection.md")
    assert matching_tree_item is not None, "matching tree page is missing"
    highlighted_tree_runs = [
        str(inline.Text) for inline in matching_tree_item.Header.Inlines
        if isinstance(inline, Run) and inline.Background is not None
    ]
    assert highlighted_tree_runs == ["Zoom", "Selection"], "tree title matches were not highlighted"
    window.search_box.Text = ""
    window._apply_page_search_highlights()
    window.navigate("orient-and-rotate-elements.md#what-it-does")
    assert window.current_relative_path == "ae-pytools/orient-and-rotate-elements.md", (
        "cross-page navigation did not reach the target"
    )
    assert window.history.can_back, "back history was not populated"
    window._navigate_history(window.history.back())
    assert window.current_relative_path == "ae-pytools/zoom-to-selection.md", (
        "back navigation did not restore the prior page"
    )
finally:
    try:
        window.Close()
    except Exception:
        pass

deferred_window = DocumentationViewerWindow(
    documentation_root=documentation_root,
    uiapp=globals().get("__revit__"),
    defer_runtime=True,
)
try:
    assert deferred_window.catalog is None, "deferred window loaded before it was shown"
    deferred_window.Show()
    deferred_window.Dispatcher.Invoke(
        DispatcherPriority.ApplicationIdle,
        Action(lambda: None),
    )
    assert deferred_window.catalog is not None, "deferred runtime did not load after show"
    assert deferred_window.error_overlay.Visibility == Visibility.Collapsed
finally:
    try:
        deferred_window.Close()
    except Exception:
        pass

print(
    "DOCUMENTATION_RENDERER_SMOKE_PASS Revit={} pyRevit={} documents={} blocks={} startup_ms={:.0f} stages={}".format(
        getattr(__revit__.Application, "VersionBuild", "unknown"),
        str(PYREVIT_VERSION_STRING).strip(),
        len(window.catalog.documents),
        rendered.Blocks.Count,
        startup_seconds * 1000.0,
        ",".join(
            "{}:{:.0f}ms".format(key, window._startup_timings[key] * 1000.0)
            for key in sorted(window._startup_timings)
        ),
    )
)
