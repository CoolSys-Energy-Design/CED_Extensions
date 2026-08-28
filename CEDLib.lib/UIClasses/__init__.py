# -*- coding: utf-8 -*-
"""UIClasses package exports."""

from UIClasses import Resources, pathing, resource_loader
from UIClasses.structured_search import (
    SearchFilter,
    SearchFilterDefinition,
    SearchFilterToken,
    SearchQuery,
    SearchQueryChangedEventArgs,
    SearchCommandChangedEventArgs,
    SearchInteractionChangedEventArgs,
    StructuredSearchState,
    make_contains_matcher,
)
from UIClasses.structured_search_box import StructuredSearchBox
from UIClasses.filterable_combo_box import (
    FALLBACK_BLANK,
    FALLBACK_DEFAULT_ITEM,
    FALLBACK_LAST_VALID,
    FilterableComboBox,
    find_exact_item,
    filter_items,
    highlight_segments,
    item_text,
    navigation_index,
    resolve_fallback_item,
)
from UIClasses.revit_theme_bridge import DOCK_PANE_FRAME_DARK, DOCK_PANE_FRAME_LIGHT, RevitThemeBridge
from UIClasses.ui_bases import (
    CEDPanelBase,
    CEDWindowBase,
    THEME_CONFIG_ACCENT_KEY,
    THEME_CONFIG_SECTION,
    THEME_CONFIG_THEME_KEY,
    load_theme_state_from_config,
    wire_textbox_select_all,
)

__all__ = [
    "CEDPanelBase",
    "CEDWindowBase",
    "DOCK_PANE_FRAME_DARK",
    "DOCK_PANE_FRAME_LIGHT",
    "Resources",
    "RevitThemeBridge",
    "pathing",
    "resource_loader",
    "SearchFilter",
    "SearchFilterDefinition",
    "SearchFilterToken",
    "SearchQuery",
    "SearchQueryChangedEventArgs",
    "SearchCommandChangedEventArgs",
    "SearchInteractionChangedEventArgs",
    "StructuredSearchState",
    "make_contains_matcher",
    "StructuredSearchBox",
    "FilterableComboBox",
    "FALLBACK_BLANK",
    "FALLBACK_DEFAULT_ITEM",
    "FALLBACK_LAST_VALID",
    "find_exact_item",
    "filter_items",
    "highlight_segments",
    "navigation_index",
    "item_text",
    "resolve_fallback_item",
    "THEME_CONFIG_SECTION",
    "THEME_CONFIG_THEME_KEY",
    "THEME_CONFIG_ACCENT_KEY",
    "load_theme_state_from_config",
    "wire_textbox_select_all",
]
