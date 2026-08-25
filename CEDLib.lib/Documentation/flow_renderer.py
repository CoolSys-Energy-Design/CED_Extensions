# -*- coding: utf-8 -*-
"""Native WPF FlowDocument renderer for the supported Markdown subset."""

from __future__ import print_function

import os

import clr

for _assembly in ("WindowsBase", "PresentationCore", "PresentationFramework"):
    try:
        clr.AddReference(_assembly)
    except Exception:
        pass

from System import Uri
from System.Windows import (
    CornerRadius,
    FontStyles,
    FontWeights,
    Style,
    TextAlignment,
    TextWrapping,
    Thickness,
)
from System.Windows.Controls import Border, Image, Orientation, StackPanel, TextBlock
from System.Windows.Documents import (
    BlockUIContainer,
    Bold,
    FlowDocument,
    Hyperlink,
    InlineUIContainer,
    Italic,
    LineBreak,
    Paragraph,
    Run,
    Section,
    Table,
    TableCell,
    TableColumn,
    TableRow,
    TableRowGroup,
)
from System.Windows.Media import Brushes, FontFamily, Stretch
from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage
from System.Windows.Shapes import Path

from Documentation.highlighting import highlight_segments
from Documentation import markdown_parser
from Documentation.pathing import resolve_local_path, split_target
from UIClasses import resource_loader


class FlowDocumentRenderer(object):
    HEADING_SIZES = {1: 28.0, 2: 22.0, 3: 18.0, 4: 16.0, 5: 14.0, 6: 13.0}

    def __init__(self, owner, documentation_root, navigate_callback=None, warning_callback=None):
        self.owner = owner
        self.documentation_root = os.path.abspath(documentation_root)
        self.navigate_callback = navigate_callback
        self.warning_callback = warning_callback
        self.heading_blocks = {}
        self._handlers = []
        self._current_document = None
        self._highlight_query = ""

    def _resource(self, key, fallback):
        value = resource_loader.try_find_resource(self.owner, key)
        return value if value is not None else fallback

    def render(self, markdown_text, source_path, highlight_query=""):
        self.heading_blocks = {}
        self._handlers = []
        self._current_document = os.path.abspath(source_path)
        self._highlight_query = str(highlight_query or "")
        document = FlowDocument()
        document.PagePadding = Thickness(30, 22, 30, 40)
        document.FontFamily = FontFamily("Segoe UI")
        document.FontSize = 13.0
        document.Foreground = self._resource("CED.Brush.PrimaryText", Brushes.Black)
        document.Background = self._resource("CED.Brush.DocumentCanvas", Brushes.White)
        try:
            style = resource_loader.try_find_resource(self.owner, "CED.Documentation.FlowDocument")
            if style is not None:
                document.Style = style
        except Exception:
            pass

        for block in markdown_parser.parse(markdown_text):
            rendered = self._render_block(block)
            if rendered is not None:
                document.Blocks.Add(rendered)
        return document

    def _render_block(self, block):
        block_type = block.get("type")
        if block_type == "heading":
            paragraph = Paragraph()
            paragraph.Margin = Thickness(0, 18 if block["level"] > 1 else 0, 0, 8)
            paragraph.FontSize = self.HEADING_SIZES.get(block["level"], 13.0)
            paragraph.FontWeight = FontWeights.SemiBold
            paragraph.Foreground = self._resource("CED.Brush.HeadingText", Brushes.Black)
            self._append_inlines(paragraph.Inlines, block["text"])
            self.heading_blocks[block["anchor"]] = paragraph
            return paragraph
        if block_type == "paragraph":
            paragraph = self._paragraph(block.get("text", ""))
            if block.get("unsupported_html"):
                paragraph.Background = self._resource("CED.Brush.AlertWarningBackground", Brushes.LemonChiffon)
                if self.warning_callback:
                    self.warning_callback("Raw HTML is unsupported and was rendered as plain text.")
            return paragraph
        if block_type == "code":
            return self._code_block(block)
        if block_type == "list":
            return self._list_block(block)
        if block_type == "table":
            return self._table_block(block)
        if block_type == "blockquote":
            return self._quote_block(block)
        if block_type == "alert":
            return self._alert_block(block)
        if block_type == "rule":
            border = Border()
            border.Height = 1
            border.Margin = Thickness(0, 14, 0, 14)
            border.Background = self._resource("CED.Brush.SectionDivider", Brushes.Gray)
            return BlockUIContainer(border)
        return None

    def _paragraph(self, text):
        paragraph = Paragraph()
        paragraph.Margin = Thickness(0, 0, 0, 10)
        paragraph.LineHeight = 20
        self._append_inlines(paragraph.Inlines, text)
        return paragraph

    def _append_inlines(self, collection, text):
        for token in markdown_parser.parse_inlines(text):
            token_type = token["type"]
            if token_type == "text":
                self._add_highlighted_runs(collection, token["text"])
            elif token_type == "bold":
                bold = Bold()
                self._add_highlighted_runs(bold.Inlines, token["text"])
                collection.Add(bold)
            elif token_type == "italic":
                italic = Italic()
                self._add_highlighted_runs(italic.Inlines, token["text"])
                collection.Add(italic)
            elif token_type == "bold_italic":
                bold = Bold()
                italic = Italic()
                self._add_highlighted_runs(italic.Inlines, token["text"])
                bold.Inlines.Add(italic)
                collection.Add(bold)
            elif token_type == "code":
                self._add_highlighted_runs(
                    collection,
                    token["text"],
                    font_family=FontFamily("Consolas"),
                    normal_background=self._resource("CED.Brush.SurfaceAlt", Brushes.Gainsboro),
                )
            elif token_type == "link":
                link = Hyperlink()
                link.Foreground = self._resource("CED.Brush.Link", Brushes.RoyalBlue)
                self._add_highlighted_runs(link.Inlines, token["text"])
                link_style = resource_loader.try_find_resource(self.owner, "CED.Text.Hyperlink")
                if isinstance(link_style, Style):
                    link.Style = link_style
                link.ToolTip = token["target"]

                def handler(sender, args, target=token["target"]):
                    if self.navigate_callback:
                        self.navigate_callback(target)

                link.Click += handler
                self._handlers.append(handler)
                collection.Add(link)
            elif token_type == "image":
                control = self._image(token["target"], token.get("alt", ""))
                collection.Add(InlineUIContainer(control))

    def _add_highlighted_runs(self, collection, text, font_family=None, normal_background=None):
        highlight_background = self._resource(
            "CED.Brush.SearchHighlightBackground",
            Brushes.Yellow,
        )
        highlight_foreground = self._resource(
            "CED.Brush.SearchHighlightForeground",
            Brushes.Black,
        )
        for value, matched in highlight_segments(text, self._highlight_query):
            run = Run(value)
            if font_family is not None:
                run.FontFamily = font_family
            if matched:
                run.Background = highlight_background
                run.Foreground = highlight_foreground
            elif normal_background is not None:
                run.Background = normal_background
            collection.Add(run)

    def _image(self, target, alt_text):
        try:
            path_text, _anchor = split_target(target)
            path = resolve_local_path(
                self.documentation_root,
                path_text,
                current_document=self._current_document,
                must_exist=True,
            )
            bitmap = BitmapImage()
            bitmap.BeginInit()
            bitmap.CacheOption = BitmapCacheOption.OnLoad
            bitmap.UriSource = Uri(path)
            bitmap.EndInit()
            image = Image()
            image.Source = bitmap
            image.MaxWidth = 720
            image.MaxHeight = 520
            image.Margin = Thickness(0, 8, 0, 8)
            image.Stretch = Stretch.Uniform
            image.ToolTip = alt_text
            return image
        except Exception as error:
            fallback = TextBlock()
            fallback.Text = "Image unavailable: {}".format(alt_text or target)
            fallback.Foreground = self._resource("CED.Brush.AlertErrorText", Brushes.DarkRed)
            if self.warning_callback:
                self.warning_callback("Image could not be loaded: {}".format(error))
            return fallback

    def _code_block(self, block):
        text = TextBlock()
        self._add_highlighted_runs(text.Inlines, block.get("text", ""))
        text.FontFamily = FontFamily("Consolas")
        text.FontSize = 12.0
        text.TextWrapping = TextWrapping.Wrap
        text.Foreground = self._resource("CED.Brush.PrimaryText", Brushes.Black)
        if block.get("language"):
            text.ToolTip = block["language"]
        border = Border()
        border.Background = self._resource("CED.Brush.SurfaceAlt", Brushes.Gainsboro)
        border.BorderBrush = self._resource("CED.Brush.Border", Brushes.Gray)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(4)
        border.Padding = Thickness(12)
        border.Margin = Thickness(0, 4, 0, 12)
        border.Child = text
        return BlockUIContainer(border)

    def _list_block(self, block):
        section = Section()
        section.Margin = Thickness(0, 0, 0, 10)
        ordered_counts = {}
        for item in block.get("items", []):
            level = int(item.get("level", 0))
            if item.get("ordered"):
                ordered_counts[level] = ordered_counts.get(level, 0) + 1
                marker = "{}.".format(ordered_counts[level])
            else:
                marker = "•" if level % 2 == 0 else "◦"
            row = Paragraph()
            row.Margin = Thickness(18 + (level * 22), 2, 0, 2)
            prefix = Run(marker + "  ")
            prefix.FontWeight = FontWeights.SemiBold
            row.Inlines.Add(prefix)
            self._append_inlines(row.Inlines, item.get("text", ""))
            section.Blocks.Add(row)
        return section

    def _table_block(self, block):
        headers = block.get("headers", [])
        rows = block.get("rows", [])
        column_count = max([len(headers)] + [len(row) for row in rows] + [1])
        table = Table()
        table.CellSpacing = 0
        table.Margin = Thickness(0, 4, 0, 14)
        for _index in range(column_count):
            table.Columns.Add(TableColumn())
        group = TableRowGroup()
        table.RowGroups.Add(group)
        if headers:
            row = TableRow()
            row.Background = self._resource("CED.Brush.SurfaceAlt", Brushes.Gainsboro)
            for value in headers:
                paragraph = self._paragraph(value)
                paragraph.Margin = Thickness(0)
                paragraph.TextAlignment = TextAlignment.Left
                paragraph.FontWeight = FontWeights.SemiBold
                cell = TableCell(paragraph)
                self._style_cell(cell)
                row.Cells.Add(cell)
            group.Rows.Add(row)
        for values in rows:
            row = TableRow()
            for index in range(column_count):
                paragraph = self._paragraph(values[index] if index < len(values) else "")
                paragraph.Margin = Thickness(0)
                paragraph.TextAlignment = TextAlignment.Left
                cell = TableCell(paragraph)
                self._style_cell(cell)
                row.Cells.Add(cell)
            group.Rows.Add(row)
        return table

    def _style_cell(self, cell):
        cell.Padding = Thickness(8, 6, 8, 6)
        cell.BorderBrush = self._resource("CED.Brush.Border", Brushes.Gray)
        cell.BorderThickness = Thickness(0.5)

    def _quote_block(self, block):
        text = TextBlock()
        text.TextWrapping = TextWrapping.Wrap
        text.FontStyle = FontStyles.Italic
        self._append_inlines(text.Inlines, block.get("text", "").replace("\n", " "))
        border = Border()
        border.BorderBrush = self._resource("CED.Brush.Accent", Brushes.RoyalBlue)
        border.BorderThickness = Thickness(4, 0, 0, 0)
        border.Background = self._resource("CED.Brush.SurfaceAlt", Brushes.Gainsboro)
        border.Padding = Thickness(12, 9, 12, 9)
        border.Margin = Thickness(0, 4, 0, 12)
        border.Child = text
        return BlockUIContainer(border)

    def _alert_block(self, block):
        alert_type = block.get("alert_type", "NOTE")
        key_map = {
            "NOTE": ("CED.Brush.AlertInfoBackground", "CED.Brush.AlertInfoBorder", "CED.Brush.AlertInfoText"),
            "TIP": ("CED.Brush.AlertSuccessBackground", "CED.Brush.AlertSuccessBorder", "CED.Brush.AlertSuccessText"),
            "IMPORTANT": ("CED.Brush.AlertImportantBackground", "CED.Brush.AlertImportantBorder", "CED.Brush.AlertImportantText"),
            "WARNING": ("CED.Brush.AlertWarningBackground", "CED.Brush.AlertWarningBorder", "CED.Brush.AlertWarningText"),
            "CAUTION": ("CED.Brush.AlertErrorBackground", "CED.Brush.AlertErrorBorder", "CED.Brush.AlertErrorText"),
        }
        icon_map = {
            "NOTE": "CED.Icon.Alert.Info",
            "TIP": "CED.Icon.Alert.Success",
            "IMPORTANT": "CED.Icon.AlertBoxOutline",
            "WARNING": "CED.Icon.Alert.Warning",
            "CAUTION": "CED.Icon.Alert.Error",
        }
        background_key, border_key, text_key = key_map.get(alert_type, key_map["NOTE"])
        panel = StackPanel()
        header = StackPanel()
        header.Orientation = Orientation.Horizontal
        header.Margin = Thickness(0, 0, 0, 4)
        text_brush = self._resource(text_key, Brushes.Black)
        icon_geometry = self._resource(icon_map.get(alert_type, icon_map["NOTE"]), None)
        if icon_geometry is not None:
            glyph = Path()
            glyph.Data = icon_geometry
            glyph.Fill = text_brush
            glyph.Width = 15
            glyph.Height = 15
            glyph.Margin = Thickness(0, 0, 6, 0)
            glyph.Stretch = Stretch.Uniform
            header.Children.Add(glyph)
        title = TextBlock()
        title.Text = alert_type.title()
        title.FontWeight = FontWeights.Bold
        title.Foreground = text_brush
        header.Children.Add(title)
        content = TextBlock()
        content.TextWrapping = TextWrapping.Wrap
        content.Foreground = text_brush
        self._append_inlines(content.Inlines, block.get("text", "").replace("\n", " "))
        panel.Children.Add(header)
        panel.Children.Add(content)
        border = Border()
        border.Background = self._resource(background_key, Brushes.LemonChiffon)
        border.BorderBrush = self._resource(border_key, Brushes.Goldenrod)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(4)
        border.Padding = Thickness(12)
        border.Margin = Thickness(0, 4, 0, 12)
        border.Child = panel
        return BlockUIContainer(border)

    def navigate_to_heading(self, anchor):
        block = self.heading_blocks.get(str(anchor or "").lower())
        if block is None:
            return False
        try:
            block.BringIntoView()
            return True
        except Exception:
            return False
