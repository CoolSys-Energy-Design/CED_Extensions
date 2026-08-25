# -*- coding: utf-8 -*-
"""Modeless searchable documentation browser window."""

from __future__ import print_function

import hashlib
import io
import os
import time

import clr

for _assembly in ("WindowsBase", "PresentationCore", "PresentationFramework", "System.Data"):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass

from System import Action, Object, TimeSpan
from System.Data import DataTable
from System.Windows import FontWeights, HorizontalAlignment, Style, TextWrapping, Visibility
from System.Windows.Controls import TextBlock, TreeViewItem
from System.Windows.Documents import Run
from System.Windows.Input import Key, Keyboard, ModifierKeys
from System.Windows.Media import Brushes
from System.Windows.Threading import DispatcherPriority, DispatcherTimer
from pyrevit import script as pyrevit_script

from Documentation.catalog import Catalog, build_ribbon_tree
from Documentation.highlighting import highlight_segments
from Documentation.history import NavigationHistory
from Documentation.markdown_parser import runtime_frontmatter
from Documentation.pathing import (
    DocumentationPathError,
    has_uri_scheme,
    is_external_http,
    relative_to_root,
    resolve_documentation_root,
    resolve_local_path,
    split_target,
)
from UIClasses import resource_loader
from UIClasses.revit_theme_bridge import RevitThemeBridge
from UIClasses.ui_bases import CEDWindowBase


TITLE = "CED Documentation"
HOME_DOCUMENT_ID = "ae-pytools-index"
DOCUMENTATION_STYLE = os.path.join("Styles", "DocumentationStyles.xaml")


class DocumentationViewerWindow(CEDWindowBase):
    theme_aware = True
    use_config_theme = True
    base_resource_relative_paths = (
        os.path.join("Themes", "CED.Sizes.xaml"),
        os.path.join("Themes", "CED.Colors.xaml"),
        os.path.join("Themes", "CED.Brushes.xaml"),
        os.path.join("Styles", "ButtonStyles.xaml"),
        os.path.join("Styles", "TextStyles.xaml"),
        os.path.join("Styles", "InputStyles.xaml"),
        os.path.join("Styles", "ListStyles.xaml"),
        os.path.join("Icons", "Icons.xaml"),
        os.path.join("Controls", "SearchBox.xaml"),
        DOCUMENTATION_STYLE,
    )

    def __init__(self, documentation_root=None, uiapp=None, logger=None, defer_runtime=True):
        self._startup_started_at = time.time()
        self._startup_timings = {}
        self._logger = logger
        self._requested_root = documentation_root
        self.documentation_root = None
        self.catalog = None
        self.renderer = None
        self.history = NavigationHistory()
        self.current_relative_path = None
        self._suppress_selection = False
        self._retry_action = None
        self._theme_bridge = None
        self._result_table = None
        self._results_refresh_count = 0
        self._tree_path_items = {}
        self._suppress_tree_selection = False
        self._suppress_refresh = False
        self._tree_dirty = True
        self._current_results = []
        self._current_markdown_text = None
        self._current_absolute_path = None
        self._rendered_highlight_query = ""
        self._runtime_started = False
        self._strong_references = [
            Catalog,
            NavigationHistory,
            resolve_documentation_root,
        ]

        xaml = os.path.join(os.path.abspath(os.path.dirname(__file__)), "DocumentationViewer.xaml")
        shell_started = time.time()
        CEDWindowBase.__init__(
            self,
            xaml_source=xaml,
            theme_aware=True,
            handle_esc=False,
        )
        self._startup_timings["shell_and_resources"] = time.time() - shell_started

        self.Tag = "_ced_documentation_viewer_modeless_v1"
        self.back_button = self.FindName("BackButton")
        self.forward_button = self.FindName("ForwardButton")
        self.home_button = self.FindName("HomeButton")
        self.search_box = self.FindName("SearchBox")
        self.search_placeholder = self.FindName("SearchPlaceholderText")
        self.clear_search_button = self.FindName("ClearSearchButton")
        self.extension_filter = self.FindName("ExtensionFilter")
        self.ribbon_filter = self.FindName("RibbonFilter")
        self.results_summary = self.FindName("ResultsSummary")
        self.results_list = self.FindName("ResultsList")
        self.results_tree = self.FindName("ResultsTree")
        self.list_mode_button = self.FindName("ListModeButton")
        self.tree_mode_button = self.FindName("TreeModeButton")
        self.tree_actions_panel = self.FindName("TreeActionsPanel")
        self.collapse_tree_button = self.FindName("CollapseTreeButton")
        self.expand_tree_button = self.FindName("ExpandTreeButton")
        self.document_viewer = self.FindName("DocumentViewer")
        self.status_text = self.FindName("StatusText")
        self.location_text = self.FindName("LocationText")
        self.error_overlay = self.FindName("ErrorOverlay")
        self.error_title = self.FindName("ErrorTitle")
        self.error_message = self.FindName("ErrorMessage")
        self.retry_button = self.FindName("RetryButton")
        self.close_error_button = self.FindName("CloseErrorButton")

        self._highlight_timer = DispatcherTimer()
        self._highlight_timer.Interval = TimeSpan.FromMilliseconds(175)
        self._highlight_timer.Tick += self._highlight_timer_tick

        self.back_button.Click += self._back_clicked
        self.forward_button.Click += self._forward_clicked
        self.home_button.Click += self._home_clicked
        self.search_box.TextChanged += self._search_changed
        self.clear_search_button.Click += self._clear_search_clicked
        self.extension_filter.SelectionChanged += self._filter_changed
        self.ribbon_filter.SelectionChanged += self._filter_changed
        self.results_list.SelectionChanged += self._result_selected
        self.results_tree.SelectedItemChanged += self._tree_result_selected
        self.list_mode_button.Checked += self._result_mode_changed
        self.tree_mode_button.Checked += self._result_mode_changed
        self.collapse_tree_button.Click += self._collapse_tree_clicked
        self.expand_tree_button.Click += self._expand_tree_clicked
        self.retry_button.Click += self._retry_clicked
        self.close_error_button.Click += self._close_error_clicked
        self.PreviewKeyDown += self._preview_key_down
        self.Closed += self._window_closed

        if uiapp is not None:
            self._theme_bridge = RevitThemeBridge(uiapp, self._revit_theme_changed, logger=logger)
            self._theme_bridge.attach()

        self.status_text.Text = "Loading documentation…"
        self.back_button.IsEnabled = False
        self.forward_button.IsEnabled = False
        if defer_runtime:
            self.Loaded += self._window_loaded
        else:
            self._runtime_started = True
            self._load_runtime()

    def _log(self, message):
        if self._logger is None:
            return
        try:
            self._logger.warning(message)
        except Exception:
            pass

    def _debug(self, message):
        if self._logger is None:
            return
        try:
            self._logger.debug(message)
        except Exception:
            pass

    def _window_loaded(self, sender, args):
        if self._runtime_started:
            return
        self._runtime_started = True
        try:
            self.Loaded -= self._window_loaded
        except Exception:
            pass
        try:
            self.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                Action(self._load_runtime),
            )
        except Exception:
            self._load_runtime()

    def _load_runtime(self):
        started = time.time()
        try:
            stage = time.time()
            self.documentation_root = resolve_documentation_root(
                start_path=os.path.dirname(__file__),
                configured_root=self._requested_root,
            )
            self._startup_timings["resolve_root"] = time.time() - stage
            stage = time.time()
            self.catalog = Catalog.load(self.documentation_root)
            self._startup_timings["catalog"] = time.time() - stage
            stage = time.time()
            from Documentation.flow_renderer import FlowDocumentRenderer

            self._strong_references.append(FlowDocumentRenderer)
            self.renderer = FlowDocumentRenderer(
                self,
                self.documentation_root,
                navigate_callback=self.navigate,
                warning_callback=self._show_warning,
            )
            self._startup_timings["renderer_import"] = time.time() - stage
            stage = time.time()
            self._populate_filters()
            self._refresh_results()
            self._startup_timings["results"] = time.time() - stage
            self._hide_error()
            home = self.catalog.by_id.get(HOME_DOCUMENT_ID)
            if home is None and self.catalog.documents:
                home = self.catalog.documents[0]
            if home is not None:
                stage = time.time()
                self._open_document(home["path"], push_history=True)
                self._startup_timings["home_render"] = time.time() - stage
            else:
                self._show_empty_document("No documentation pages are available.")
            self._debug(
                "Documentation viewer ready in {:.3f}s (runtime {:.3f}s; {}).".format(
                    time.time() - self._startup_started_at,
                    time.time() - started,
                    ", ".join(
                        "{}={:.3f}s".format(key, self._startup_timings[key])
                        for key in sorted(self._startup_timings)
                    ),
                )
            )
        except Exception as error:
            self._show_error(
                "Documentation unavailable",
                "The offline documentation could not be loaded.\n\n{}".format(error),
                retry_action=self._load_runtime,
            )

    def _populate_filters(self):
        previous_suppression = self._suppress_refresh
        self._suppress_refresh = True
        try:
            self.extension_filter.Items.Clear()
            self.ribbon_filter.Items.Clear()
            self.extension_filter.Items.Add("All extensions")
            self.ribbon_filter.Items.Add("All ribbon locations")
            for value in self.catalog.extensions:
                self.extension_filter.Items.Add(value)
            for value in self.catalog.ribbon_paths:
                self.ribbon_filter.Items.Add(value)
            self.extension_filter.SelectedIndex = 0
            self.ribbon_filter.SelectedIndex = 0
        finally:
            self._suppress_refresh = previous_suppression

    def _selected_filter(self, control, all_label):
        selected = control.SelectedItem
        if selected is None:
            return ""
        value = str(selected)
        return "" if value == all_label else value

    def _refresh_results(self):
        if self.catalog is None:
            return
        self._results_refresh_count += 1
        query = str(self.search_box.Text or "")
        extension = self._selected_filter(self.extension_filter, "All extensions")
        ribbon = self._selected_filter(self.ribbon_filter, "All ribbon locations")
        results = self.catalog.search(query, extension, ribbon)
        table = DataTable()
        for name in ("DocumentId", "CatalogPath", "Title", "Extension", "RibbonPath"):
            table.Columns.Add(name)
        for name in ("TitleDisplay", "ExtensionDisplay", "RibbonPathDisplay"):
            table.Columns.Add(name, Object)
        for item in results:
            row = table.NewRow()
            row["DocumentId"] = item["id"]
            row["CatalogPath"] = item["path"]
            row["Title"] = item["title"]
            row["Extension"] = item["extension"]
            row["RibbonPath"] = item["ribbon_location"]
            row["TitleDisplay"] = self._highlight_textblock(
                item["title"],
                query,
                foreground_key="CED.Brush.PrimaryText",
                font_weight=FontWeights.SemiBold,
            )
            row["ExtensionDisplay"] = self._highlight_textblock(
                item["extension"],
                query,
                foreground_key="CED.Brush.SecondaryText",
                font_size=11.0,
            )
            row["RibbonPathDisplay"] = self._highlight_textblock(
                item["ribbon_location"],
                query,
                foreground_key="CED.Brush.SecondaryText",
                font_size=11.0,
            )
            table.Rows.Add(row)
        self._result_table = table
        self._suppress_selection = True
        try:
            self.results_list.ItemsSource = table.DefaultView
        finally:
            self._suppress_selection = False
        self._current_results = list(results)
        self._tree_dirty = True
        if bool(self.tree_mode_button.IsChecked):
            self._ensure_tree()
            if query.strip():
                self._set_tree_expanded(True)
        suffix = "result" if len(results) == 1 else "results"
        self.results_summary.Text = "{} {}".format(len(results), suffix)
        if not results and query:
            self.status_text.Text = "No documentation matched the current search and filters."

    def _highlight_textblock(
        self,
        value,
        query,
        foreground_key="CED.Brush.PrimaryText",
        font_weight=None,
        font_size=None,
    ):
        control = TextBlock()
        control.TextWrapping = TextWrapping.Wrap
        control.Foreground = resource_loader.try_find_resource(self, foreground_key) or Brushes.Black
        if font_weight is not None:
            control.FontWeight = font_weight
        if font_size is not None:
            control.FontSize = font_size
        highlight_background = resource_loader.try_find_resource(
            self,
            "CED.Brush.SearchHighlightBackground",
        ) or Brushes.Yellow
        highlight_foreground = resource_loader.try_find_resource(
            self,
            "CED.Brush.SearchHighlightForeground",
        ) or Brushes.Black
        for segment, matched in highlight_segments(value, query):
            run = Run(segment)
            if matched:
                run.Background = highlight_background
                run.Foreground = highlight_foreground
            control.Inlines.Add(run)
        return control

    def _tree_item(self, header, tag=None, level=0, is_group=False):
        item = TreeViewItem()
        item.Header = self._highlight_textblock(
            header,
            str(self.search_box.Text or ""),
            font_weight=FontWeights.SemiBold if is_group else FontWeights.Normal,
        )
        item.Tag = tag
        item.HorizontalContentAlignment = HorizontalAlignment.Stretch
        item.FontWeight = FontWeights.SemiBold if is_group else FontWeights.Normal
        style = resource_loader.try_find_resource(self, "CED.Documentation.TreeItem")
        if isinstance(style, Style):
            item.Style = style
        return item

    def _populate_tree(self, results):
        self._suppress_tree_selection = True
        self._tree_path_items = {}
        try:
            self.results_tree.Items.Clear()
            tree = build_ribbon_tree(results, index_documents=self.catalog.documents)

            def add_document(parent_items, document, level):
                item = self._tree_item(
                    document["title"],
                    tag=document["path"],
                    level=level,
                    is_group=False,
                )
                item.ToolTip = document["ribbon_location"]
                parent_items.Add(item)
                self._tree_path_items[document["path"].lower()] = item

            def add_group(parent_items, node, level):
                index_document = node.get("index")
                group = self._tree_item(
                    node["label"],
                    tag=index_document.get("path") if index_document else None,
                    level=level,
                    is_group=True,
                )
                if index_document:
                    group.ToolTip = index_document.get("summary") or index_document.get("title")
                    self._tree_path_items[index_document["path"].lower()] = group
                parent_items.Add(group)
                for document in node.get("documents", []):
                    add_document(group.Items, document, level + 1)
                for child in node.get("groups", []):
                    add_group(group.Items, child, level + 1)

            for document in tree.get("documents", []):
                add_document(self.results_tree.Items, document, 0)
            for node in tree.get("groups", []):
                add_group(self.results_tree.Items, node, 0)
        finally:
            self._suppress_tree_selection = False
        self._tree_dirty = False
        self._select_tree_result(self.current_relative_path)

    def _ensure_tree(self):
        if self._tree_dirty:
            self._populate_tree(self._current_results)

    def _select_tree_result(self, relative_path):
        if not relative_path or self._tree_dirty:
            return
        tree_item = self._tree_path_items.get(str(relative_path).lower())
        if tree_item is None:
            return
        previous_suppression = self._suppress_tree_selection
        self._suppress_tree_selection = True
        try:
            parent = tree_item.Parent
            while isinstance(parent, TreeViewItem):
                parent.IsExpanded = True
                parent = parent.Parent
            if tree_item.HasItems:
                tree_item.IsExpanded = True
            tree_item.IsSelected = True
            tree_item.BringIntoView()
        finally:
            self._suppress_tree_selection = previous_suppression

    def _open_document(self, relative_path, anchor="", push_history=True):
        if not self.documentation_root:
            return False
        try:
            absolute = resolve_local_path(
                self.documentation_root,
                relative_path,
                current_document=None,
                must_exist=True,
            )
            if os.path.splitext(absolute)[1].lower() != ".md":
                raise DocumentationPathError("Only Markdown pages open inside the viewer.")
            with io.open(absolute, "r", encoding="utf-8-sig") as stream:
                markdown_text = stream.read()
            page_metadata = runtime_frontmatter(markdown_text)
            self.current_relative_path = relative_to_root(self.documentation_root, absolute)
            catalog_item = self.catalog.get_by_path(self.current_relative_path) if self.catalog else None
            if catalog_item and str(page_metadata.get("id")) != catalog_item["id"]:
                raise ValueError(
                    "Page id '{}' does not match catalog id '{}'.".format(
                        page_metadata.get("id"), catalog_item["id"]
                    )
                )
            stale_message = ""
            if catalog_item:
                source_hash = hashlib.sha256(markdown_text.encode("utf-8")).hexdigest()
                if source_hash.lower() != catalog_item["source_sha256"]:
                    stale_message = (
                        "This page is newer than the installed search catalog; search results may be stale."
                    )
            self.status_text.Text = ""
            highlight_query = str(self.search_box.Text or "")
            document = self.renderer.render(
                markdown_text,
                absolute,
                highlight_query=highlight_query,
            )
            self.document_viewer.Document = document
            self._current_markdown_text = markdown_text
            self._current_absolute_path = absolute
            self._rendered_highlight_query = highlight_query
            self.location_text.Text = (
                catalog_item["ribbon_location"] if catalog_item else self.current_relative_path
            )
            if push_history:
                self.history.push(self.current_relative_path, anchor)
            self._update_navigation_buttons()
            self._select_result(self.current_relative_path)
            self._hide_error()
            if anchor and not self.renderer.navigate_to_heading(anchor):
                self._show_warning("The heading '#{}' was not found on this page.".format(anchor))
                return False
            try:
                self.document_viewer.ScrollToHome()
            except Exception:
                pass
            if anchor:
                self.renderer.navigate_to_heading(anchor)
            if stale_message:
                self._show_warning(stale_message)
            return True
        except Exception as error:
            self._show_error(
                "Page unavailable",
                "The requested documentation page could not be opened.\n\n{}".format(error),
                retry_action=lambda: self._open_document(relative_path, anchor, push_history=False),
            )
            return False

    def _select_result(self, relative_path):
        self._suppress_selection = True
        self._suppress_tree_selection = True
        try:
            list_match = None
            for item in self.results_list.Items:
                try:
                    if str(item["CatalogPath"]).lower() == str(relative_path).lower():
                        list_match = item
                        break
                except Exception:
                    continue
            self.results_list.SelectedItem = list_match
            if list_match is not None:
                self.results_list.ScrollIntoView(list_match)
            self._select_tree_result(relative_path)
        finally:
            self._suppress_selection = False
            self._suppress_tree_selection = False

    def navigate(self, target):
        target = str(target or "").strip()
        if not target:
            return
        if is_external_http(target):
            self._open_external(target)
            return
        if has_uri_scheme(target):
            self._show_error(
                "Unsupported link",
                "Only HTTP and HTTPS external links may be opened.",
                retry_action=None,
            )
            return
        path_part, anchor = split_target(target)
        if not path_part:
            if not self.current_relative_path:
                return
            self.history.push(self.current_relative_path, anchor)
            self._update_navigation_buttons()
            if not self.renderer.navigate_to_heading(anchor):
                self._show_warning("The heading '#{}' was not found on this page.".format(anchor))
            return
        extension = os.path.splitext(path_part)[1].lower()
        if extension != ".md":
            self._show_error(
                "Unsupported local link",
                "Only Markdown pages open inside the documentation viewer.",
                retry_action=None,
            )
            return
        try:
            current = resolve_local_path(
                self.documentation_root,
                self.current_relative_path,
                must_exist=True,
            )
            absolute = resolve_local_path(
                self.documentation_root,
                path_part,
                current_document=current,
                must_exist=True,
            )
            relative = relative_to_root(self.documentation_root, absolute)
            self._open_document(relative, anchor=anchor, push_history=True)
        except Exception as error:
            self._show_error("Broken documentation link", str(error), retry_action=None)

    def _open_external(self, target):
        try:
            pyrevit_script.open_url(target)
            self.status_text.Text = "Opened the external link in the system browser."
        except Exception as error:
            self._show_error("External link failed", str(error), retry_action=None)

    def _show_empty_document(self, message):
        from System.Windows.Documents import FlowDocument, Paragraph, Run

        document = FlowDocument()
        document.Blocks.Add(Paragraph(Run(message)))
        self.document_viewer.Document = document

    def _show_warning(self, message):
        self.status_text.Text = str(message)
        self._log(message)

    def _show_error(self, title, message, retry_action=None):
        self.error_title.Text = str(title)
        self.error_message.Text = str(message)
        self._retry_action = retry_action
        self.retry_button.Visibility = Visibility.Visible if retry_action else Visibility.Collapsed
        self.error_overlay.Visibility = Visibility.Visible

    def _hide_error(self):
        self.error_overlay.Visibility = Visibility.Collapsed
        self._retry_action = None

    def _update_navigation_buttons(self):
        self.back_button.IsEnabled = self.history.can_back
        self.forward_button.IsEnabled = self.history.can_forward

    def _navigate_history(self, item):
        if item:
            self._open_document(item[0], anchor=item[1], push_history=False)
        self._update_navigation_buttons()

    def _back_clicked(self, sender, args):
        self._navigate_history(self.history.back())

    def _forward_clicked(self, sender, args):
        self._navigate_history(self.history.forward())

    def _home_clicked(self, sender, args):
        previous_suppression = self._suppress_refresh
        self._suppress_refresh = True
        try:
            self.search_box.Text = ""
            self.extension_filter.SelectedIndex = 0
            self.ribbon_filter.SelectedIndex = 0
        finally:
            self._suppress_refresh = previous_suppression
        self._update_search_chrome()
        self._refresh_results()
        home = self.catalog.by_id.get(HOME_DOCUMENT_ID) if self.catalog else None
        if home:
            self._open_document(home["path"], push_history=True)

    def _search_changed(self, sender, args):
        self._update_search_chrome()
        if self._suppress_refresh:
            return
        self._refresh_results()
        self._schedule_page_highlight()

    @staticmethod
    def _is_find_shortcut(key, modifiers):
        try:
            return key == Key.F and modifiers == ModifierKeys.Control
        except Exception:
            return False

    def _preview_key_down(self, sender, args):
        if not self._is_find_shortcut(getattr(args, "Key", None), Keyboard.Modifiers):
            return
        try:
            self.search_box.Focus()
            self.search_box.SelectAll()
        finally:
            args.Handled = True

    def _schedule_page_highlight(self):
        if self._highlight_timer is None:
            return
        self._highlight_timer.Stop()
        self._highlight_timer.Start()

    def _highlight_timer_tick(self, sender, args):
        self._highlight_timer.Stop()
        self._apply_page_search_highlights()

    def _apply_page_search_highlights(self):
        if (
            self.renderer is None
            or self._current_markdown_text is None
            or not self._current_absolute_path
        ):
            return
        query = str(self.search_box.Text or "")
        if query == self._rendered_highlight_query:
            return
        try:
            vertical_offset = self.document_viewer.VerticalOffset
        except Exception:
            vertical_offset = 0.0
        document = self.renderer.render(
            self._current_markdown_text,
            self._current_absolute_path,
            highlight_query=query,
        )
        self.document_viewer.Document = document
        self._rendered_highlight_query = query
        try:
            self.document_viewer.UpdateLayout()
            self.document_viewer.ScrollToVerticalOffset(vertical_offset)
        except Exception:
            pass

    def _update_search_chrome(self):
        has_text = bool(str(self.search_box.Text or "").strip())
        self.clear_search_button.Visibility = Visibility.Visible if has_text else Visibility.Collapsed
        self.search_placeholder.Visibility = Visibility.Collapsed if has_text else Visibility.Visible

    def _clear_search_clicked(self, sender, args):
        self.search_box.Text = ""
        self.search_box.Focus()

    def _filter_changed(self, sender, args):
        if self._suppress_refresh:
            return
        self._refresh_results()

    def _result_selected(self, sender, args):
        if self._suppress_selection:
            return
        selected = self.results_list.SelectedItem
        if selected is None:
            return
        try:
            self._open_document(str(selected["CatalogPath"]), push_history=True)
        except Exception as error:
            self._show_error("Page unavailable", str(error), retry_action=None)

    def _tree_result_selected(self, sender, args):
        if self._suppress_tree_selection:
            return
        selected = self.results_tree.SelectedItem
        if selected is None:
            return
        path = str(getattr(selected, "Tag", "") or "")
        if not path:
            return
        self._open_document(path, push_history=True)

    def _result_mode_changed(self, sender, args):
        tree_mode = bool(self.tree_mode_button.IsChecked)
        if tree_mode:
            self._ensure_tree()
        self.results_list.Visibility = Visibility.Collapsed if tree_mode else Visibility.Visible
        self.results_tree.Visibility = Visibility.Visible if tree_mode else Visibility.Collapsed
        self.tree_actions_panel.Visibility = Visibility.Visible if tree_mode else Visibility.Collapsed

    def _set_tree_expanded(self, expanded):
        self._ensure_tree()

        def visit(items):
            for item in list(items):
                if not isinstance(item, TreeViewItem):
                    continue
                item.IsExpanded = bool(expanded)
                visit(item.Items)

        visit(self.results_tree.Items)

    def _collapse_tree_clicked(self, sender, args):
        self._set_tree_expanded(False)

    def _expand_tree_clicked(self, sender, args):
        self._set_tree_expanded(True)

    def _retry_clicked(self, sender, args):
        action = self._retry_action
        self._hide_error()
        if action:
            action()

    def _close_error_clicked(self, sender, args):
        self._hide_error()

    def _revit_theme_changed(self, is_dark):
        try:
            self.refresh_ced_theme_from_config()
        except Exception:
            pass

    def _window_closed(self, sender, args):
        if self._highlight_timer is not None:
            self._highlight_timer.Stop()
            self._highlight_timer.Tick -= self._highlight_timer_tick
            self._highlight_timer = None
        if self._theme_bridge is not None:
            self._theme_bridge.detach()
        try:
            self.PreviewKeyDown -= self._preview_key_down
        except Exception:
            pass
        self._strong_references = None
