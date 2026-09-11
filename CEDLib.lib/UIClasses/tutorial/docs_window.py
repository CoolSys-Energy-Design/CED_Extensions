# -*- coding: utf-8 -*-
"""A real window for the user guides, instead of the pyRevit output pane.

The guides in ``docs/user-guide`` are proper markdown - headings, tables,
numbered steps, cross-links - so they get a proper reader: a themed window with
a contents list on the left, the rendered page on the right, working links
between guides, and a Back button.

Rendering goes to a WPF FlowDocument rather than HTML in a browser control, so
the page inherits the CED theme, selects and copies like text, and needs no
external dependency. The markdown subset covered is what the guides actually
use, measured across all 90 of them: h1-h3, paragraphs, bullet and numbered
lists, pipe tables, blockquotes, fenced and inline code, bold, italic,
horizontal rules and links.
"""

import io
import os
import re

import clr

for _wpf_asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_wpf_asm)
    except Exception:
        pass

from System.Windows import (
    FontStyles,
    FontWeights,
    GridLength,
    GridUnitType,
    TextMarkerStyle,
    Thickness,
    Visibility,
)
from System.Windows.Controls import ListBoxItem
from System.Windows.Documents import (
    Bold,
    FlowDocument,
    Hyperlink,
    Italic,
    ListItem,
    Paragraph,
    Run,
    Section,
    Span,
    Table,
    TableCell,
    TableColumn,
    TableRow,
    TableRowGroup,
)
from System.Windows.Documents import List as FlowList
from System.Windows.Media import FontFamily

from pyrevit import script

from UIClasses.ui_bases import CEDWindowBase, load_theme_state_from_config

THIS_DIR = os.path.abspath(os.path.dirname(__file__))
XAML_PATH = os.path.join(THIS_DIR, "DocsWindow.xaml")

_LOGGER = script.get_logger()

HEADING_SIZES = {1: 21.0, 2: 16.0, 3: 13.5, 4: 12.5, 5: 12.0, 6: 12.0}
BODY_SIZE = 12.5
MONO = FontFamily("Consolas, Courier New, monospace")

_INLINE = re.compile(
    r"(?P<code>`[^`]+`)"
    r"|(?P<bold>\*\*[^*]+\*\*)"
    r"|(?P<italic>(?<!\*)\*(?!\*)[^*]+\*(?!\*))"
    r"|(?P<link>\[[^\]]+\]\([^)]+\))"
)


# ---------------------------------------------------------------------------
# markdown -> FlowDocument
# ---------------------------------------------------------------------------

class MarkdownRenderer(object):
    """Turns one guide into a FlowDocument, collecting its headings."""

    def __init__(self, brushes=None, on_link=None, skip_first_h1=True):
        self.brushes = brushes or {}
        self.on_link = on_link
        self.skip_first_h1 = skip_first_h1
        self.headings = []          # (level, text, Paragraph)
        self.front_matter = {}      # the guide's YAML header, already stripped
        self.doc_title = None       # from front matter, or the first h1

    def brush(self, key, fallback=None):
        value = self.brushes.get(key)
        return value if value is not None else fallback

    # -- inline ----------------------------------------------------------

    def _inlines(self, text):
        out = []
        pos = 0
        for match in _INLINE.finditer(text):
            if match.start() > pos:
                out.append(Run(self._unescape(text[pos:match.start()])))
            kind = match.lastgroup
            raw = match.group()
            if kind == "code":
                run = Run(raw[1:-1])
                run.FontFamily = MONO
                run.FontSize = BODY_SIZE - 0.5
                accent = self.brush("CED.Brush.Accent")
                if accent is not None:
                    run.Foreground = accent
                out.append(run)
            elif kind == "bold":
                out.append(Bold(Run(self._unescape(raw[2:-2]))))
            elif kind == "italic":
                out.append(Italic(Run(self._unescape(raw[1:-1]))))
            elif kind == "link":
                label, target = re.match(r"\[([^\]]+)\]\(([^)]+)\)", raw).groups()
                out.append(self._link(self._unescape(label), target))
            pos = match.end()
        if pos < len(text):
            out.append(Run(self._unescape(text[pos:])))
        return out

    @staticmethod
    def _unescape(text):
        return text.replace("\\*", "*").replace("\\_", "_").replace("\\`", "`")

    def _link(self, label, target):
        if self.on_link is None or not target:
            run = Run(label)
            accent = self.brush("CED.Brush.Accent")
            if accent is not None:
                run.Foreground = accent
            return run
        link = Hyperlink(Run(label))
        link.Tag = target
        try:
            link.ToolTip = target
        except Exception:
            pass
        accent = self.brush("CED.Brush.Accent")
        if accent is not None:
            link.Foreground = accent
        link.Click += self._link_clicked
        return link

    def _link_clicked(self, sender, args):
        try:
            self.on_link(str(getattr(sender, "Tag", "") or ""))
        except Exception as exc:
            _LOGGER.debug("Docs link failed: %s", exc)

    def _para(self, text, size=None, bold=False, top=0.0, bottom=6.0):
        para = Paragraph()
        for inline in self._inlines(text):
            para.Inlines.Add(inline)
        para.FontSize = size or BODY_SIZE
        if bold:
            para.FontWeight = FontWeights.SemiBold
        para.Margin = Thickness(0, top, 0, bottom)
        return para

    # -- blocks ----------------------------------------------------------

    def render(self, text):
        doc = FlowDocument()
        doc.FontSize = BODY_SIZE
        doc.PagePadding = Thickness(14, 12, 14, 16)
        doc.ColumnWidth = 100000.0          # one column, never newspaper mode
        doc.LineHeight = BODY_SIZE * 1.55
        fg = self.brush("CED.Brush.PrimaryText")
        if fg is not None:
            doc.Foreground = fg

        body = self._strip_front_matter(str(text or "").replace("\r\n", "\n"))
        lines = body.split("\n")
        i = 0
        n = len(lines)
        seen_h1 = False
        while i < n:
            line = lines[i]
            stripped = line.strip()

            if not stripped:
                i += 1
                continue

            # fenced code
            if stripped.startswith("```"):
                body, i = self._collect_fence(lines, i)
                doc.Blocks.Add(self._code_block(body))
                continue

            # heading
            m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
            if m:
                level = len(m.group(1))
                title = m.group(2).strip().rstrip("#").strip()
                if level == 1 and not seen_h1:
                    seen_h1 = True
                    if not self.doc_title:
                        self.doc_title = title
                    # the window header already shows the page title
                    if self.skip_first_h1:
                        i += 1
                        continue
                para = self._para(title, size=HEADING_SIZES.get(level, BODY_SIZE),
                                  bold=True, top=(14.0 if doc.Blocks.Count else 0.0),
                                  bottom=4.0)
                heading_brush = self.brush("CED.Brush.HeadingText")
                if heading_brush is not None:
                    para.Foreground = heading_brush
                doc.Blocks.Add(para)
                self.headings.append((level, title, para))
                i += 1
                continue

            # horizontal rule
            if re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped):
                doc.Blocks.Add(self._rule())
                i += 1
                continue

            # table
            if stripped.startswith("|") and i + 1 < n and re.match(
                    r"^\|?[\s:|-]+\|[\s:|-]*$", lines[i + 1].strip()):
                table, i = self._collect_table(lines, i)
                doc.Blocks.Add(table)
                continue

            # blockquote
            if stripped.startswith(">"):
                body, i = self._collect_quote(lines, i)
                doc.Blocks.Add(self._quote_block(body))
                continue

            # lists
            if re.match(r"^[-*+]\s+\S", stripped):
                block, i = self._collect_list(lines, i, ordered=False)
                doc.Blocks.Add(block)
                continue
            if re.match(r"^\d+[.)]\s+\S", stripped):
                block, i = self._collect_list(lines, i, ordered=True)
                doc.Blocks.Add(block)
                continue

            # paragraph: join until a blank line or a new block starts
            chunk = []
            while i < n:
                nxt = lines[i].strip()
                if not nxt or self._starts_block(nxt):
                    break
                chunk.append(nxt)
                i += 1
            doc.Blocks.Add(self._para(" ".join(chunk)))

        return doc

    def _strip_front_matter(self, text):
        """Remove the YAML header every guide carries, keeping its fields.

        catalog.json is generated from these, so the block is always present -
        and rendering it as a paragraph is exactly the kind of thing that makes
        a docs page look broken.
        """
        if not text.startswith("---"):
            return text
        lines = text.split("\n")
        if lines[0].strip() != "---":
            return text
        for index in range(1, min(len(lines), 60)):
            if lines[index].strip() in ("---", "..."):
                for raw in lines[1:index]:
                    pair = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(.*)$", raw.strip())
                    if pair:
                        value = pair.group(2).strip().strip('"').strip("'")
                        self.front_matter[pair.group(1)] = value
                self.doc_title = self.front_matter.get("title") or None
                return "\n".join(lines[index + 1:])
        return text

    @staticmethod
    def _starts_block(stripped):
        return bool(
            stripped.startswith("```")
            or stripped.startswith(">")
            or stripped.startswith("|")
            or re.match(r"^#{1,6}\s", stripped)
            or re.match(r"^[-*+]\s+\S", stripped)
            or re.match(r"^\d+[.)]\s+\S", stripped)
            or re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped)
        )

    def _collect_fence(self, lines, i):
        i += 1
        body = []
        while i < len(lines) and not lines[i].strip().startswith("```"):
            body.append(lines[i])
            i += 1
        return "\n".join(body), i + 1

    def _collect_quote(self, lines, i):
        body = []
        while i < len(lines) and lines[i].strip().startswith(">"):
            body.append(re.sub(r"^\s*>\s?", "", lines[i]))
            i += 1
        return " ".join(part.strip() for part in body if part.strip()), i

    def _collect_list(self, lines, i, ordered):
        pattern = r"^\d+[.)]\s+(.*)$" if ordered else r"^[-*+]\s+(.*)$"
        items = []
        while i < len(lines):
            stripped = lines[i].strip()
            if not stripped:
                # a blank line inside a list is tolerated if the list continues
                if i + 1 < len(lines) and re.match(pattern, lines[i + 1].strip()):
                    i += 1
                    continue
                break
            m = re.match(pattern, stripped)
            if m:
                items.append(m.group(1).strip())
                i += 1
                continue
            if items and not self._starts_block(stripped):
                items[-1] = items[-1] + " " + stripped      # wrapped line
                i += 1
                continue
            break

        flow = FlowList()
        flow.MarkerStyle = TextMarkerStyle.Decimal if ordered else TextMarkerStyle.Disc
        flow.Margin = Thickness(16, 2, 0, 8)
        flow.Padding = Thickness(0)
        for text in items:
            para = self._para(text, bottom=2.0)
            flow.ListItems.Add(ListItem(para))
        return flow, i

    def _collect_table(self, lines, i):
        def cells(row):
            row = row.strip()
            if row.startswith("|"):
                row = row[1:]
            if row.endswith("|"):
                row = row[:-1]
            return [c.strip() for c in row.split("|")]

        header = cells(lines[i])
        i += 2                                  # header + separator
        body = []
        while i < len(lines) and lines[i].strip().startswith("|"):
            body.append(cells(lines[i]))
            i += 1

        table = Table()
        table.CellSpacing = 0
        table.Margin = Thickness(0, 4, 0, 10)
        width = max(len(header), max([len(r) for r in body]) if body else 0)
        for _ in range(width):
            col = TableColumn()
            col.Width = GridLength(1, GridUnitType.Star)
            table.Columns.Add(col)

        border = self.brush("CED.Brush.DataGridCellBorder") or self.brush("CED.Brush.Border")
        head_bg = self.brush("CED.Brush.DataGridHeaderBackground") or self.brush("CED.Brush.SurfaceAlt")

        group = TableRowGroup()

        def add_row(values, is_header):
            row = TableRow()
            if is_header and head_bg is not None:
                row.Background = head_bg
            for index in range(width):
                text = values[index] if index < len(values) else ""
                cell = TableCell(self._para(text, bold=is_header, bottom=0.0))
                cell.Padding = Thickness(7, 4, 7, 4)
                cell.BorderThickness = Thickness(0, 0, 0, 1)
                if border is not None:
                    cell.BorderBrush = border
                row.Cells.Add(cell)
            group.Rows.Add(row)

        add_row(header, True)
        for values in body:
            add_row(values, False)
        table.RowGroups.Add(group)
        return table, i

    def _code_block(self, body):
        para = Paragraph(Run(body))
        para.FontFamily = MONO
        para.FontSize = BODY_SIZE - 0.5
        para.Margin = Thickness(0, 4, 0, 10)
        para.Padding = Thickness(10, 8, 10, 8)
        bg = self.brush("CED.Brush.SurfaceAlt")
        if bg is not None:
            para.Background = bg
        edge = self.brush("CED.Brush.Border")
        if edge is not None:
            para.BorderBrush = edge
            para.BorderThickness = Thickness(1)
        return para

    def _quote_block(self, body):
        section = Section()
        para = self._para(body, bottom=0.0)
        para.FontStyle = FontStyles.Italic
        section.Blocks.Add(para)
        section.Margin = Thickness(2, 4, 0, 10)
        section.Padding = Thickness(10, 6, 8, 6)
        edge = self.brush("CED.Brush.Accent") or self.brush("CED.Brush.Border")
        if edge is not None:
            section.BorderBrush = edge
            section.BorderThickness = Thickness(3, 0, 0, 0)
        bg = self.brush("CED.Brush.InfoPanelBackground")
        if bg is not None:
            section.Background = bg
        return section

    def _rule(self):
        para = Paragraph()
        para.Margin = Thickness(0, 8, 0, 10)
        para.Padding = Thickness(0)
        edge = self.brush("CED.Brush.SectionDivider") or self.brush("CED.Brush.Border")
        if edge is not None:
            para.BorderBrush = edge
            para.BorderThickness = Thickness(0, 0, 0, 1)
        return para


# ---------------------------------------------------------------------------
# the window
# ---------------------------------------------------------------------------

BRUSH_KEYS = (
    "CED.Brush.PrimaryText", "CED.Brush.HeadingText", "CED.Brush.Accent",
    "CED.Brush.Border", "CED.Brush.SectionDivider", "CED.Brush.Surface",
    "CED.Brush.SurfaceAlt", "CED.Brush.InfoPanelBackground",
    "CED.Brush.DataGridCellBorder", "CED.Brush.DataGridHeaderBackground",
)


class DocsWindow(CEDWindowBase):
    """Reader for the markdown user guides."""

    theme_aware = True

    def __init__(self, doc_path, entry=None, docs_root=None, caller_file=None):
        self._docs_root = docs_root
        self._caller_file = caller_file
        self._history = []
        self._current = None
        self._heading_blocks = []
        self._suppress_contents = False

        theme_mode, accent_mode = load_theme_state_from_config(
            default_theme="light", default_accent="neutral")
        CEDWindowBase.__init__(self, xaml_source=XAML_PATH,
                               theme_mode=theme_mode, accent_mode=accent_mode,
                               theme_aware=True, handle_esc=True)

        self._title_text = self.FindName("DocTitleText")
        self._meta_text = self.FindName("DocMetaText")
        self._source_text = self.FindName("SourceText")
        self._viewer = self.FindName("DocViewer")
        self._contents = self.FindName("ContentsList")
        self._back = self.FindName("BackButton")
        self._tour = self.FindName("TourButton")
        self._open_file = self.FindName("OpenFileButton")
        self._close = self.FindName("CloseButton")

        if self._back is not None:
            self._back.Click += self._back_clicked
        if self._close is not None:
            self._close.Click += self._close_clicked
        if self._open_file is not None:
            self._open_file.Click += self._open_source_clicked
        if self._tour is not None:
            self._tour.Click += self._tour_clicked
            if not self._tour_available():
                self._tour.Visibility = Visibility.Collapsed
        if self._contents is not None:
            self._contents.SelectionChanged += self._contents_selected

        self._brushes = {}
        for key in BRUSH_KEYS:
            try:
                self._brushes[key] = self.TryFindResource(key)
            except Exception:
                self._brushes[key] = None

        self.load(doc_path, entry=entry, push_history=False)

    # -- loading ---------------------------------------------------------

    def _tour_available(self):
        if not self._caller_file:
            return False
        return os.path.exists(os.path.join(
            os.path.dirname(os.path.abspath(self._caller_file)), "tutorial_content.py"))

    def load(self, doc_path, entry=None, push_history=True):
        path = os.path.abspath(doc_path)
        if not os.path.exists(path):
            _LOGGER.debug("Docs page not found: %s", path)
            return False
        try:
            text = io.open(path, encoding="utf-8-sig").read()
        except Exception as exc:
            _LOGGER.warning("Could not read %s: %s", path, exc)
            return False

        if push_history and self._current:
            self._history.append(self._current)
        self._current = path

        renderer = MarkdownRenderer(brushes=self._brushes, on_link=self._navigate)
        try:
            document = renderer.render(text)
        except Exception as exc:
            _LOGGER.warning("Could not render %s: %s", path, exc)
            return False
        if self._viewer is not None:
            self._viewer.Document = document

        title = (renderer.doc_title or self._first_heading(renderer)
                 or (entry or {}).get("title")
                 or os.path.splitext(os.path.basename(path))[0])
        if self._title_text is not None:
            self._title_text.Text = str(title)
        if self._meta_text is not None:
            self._meta_text.Text = self._meta_line(entry or renderer.front_matter)
        if self._source_text is not None:
            rel = path
            if self._docs_root:
                try:
                    rel = os.path.relpath(path, self._docs_root)
                except Exception:
                    rel = path
            self._source_text.Text = "docs/user-guide/{0}".format(rel.replace("\\", "/"))

        self._fill_contents(renderer)
        if self._back is not None:
            self._back.IsEnabled = bool(self._history)
        return True

    @staticmethod
    def _first_heading(renderer):
        for level, text, _block in renderer.headings:
            if level == 1:
                return text
        return renderer.headings[0][1] if renderer.headings else None

    @staticmethod
    def _meta_line(entry):
        if not entry:
            return ""
        bits = []
        if entry.get("ribbon_path"):
            bits.append(str(entry["ribbon_path"]))
        if entry.get("status"):
            bits.append("status: {0}".format(entry["status"]))
        if entry.get("last_verified"):
            bits.append("last verified {0}".format(entry["last_verified"]))
        return "   |   ".join(bits)

    def _fill_contents(self, renderer):
        if self._contents is None:
            return
        self._suppress_contents = True
        try:
            self._contents.Items.Clear()
            self._heading_blocks = []
            for level, text, block in renderer.headings:
                if level > 3:
                    continue
                item = ListBoxItem()
                item.Content = text
                item.Padding = Thickness(6 + (level - 1) * 12, 3, 6, 3)
                item.FontSize = 11.5
                # Set the colour explicitly - the shared ListBoxItem style is
                # written for the circuit rows and leaves these unreadable.
                fg = self._brushes.get("CED.Brush.PrimaryText")
                if fg is not None:
                    item.Foreground = fg
                if level == 1:
                    item.FontWeight = FontWeights.SemiBold
                self._contents.Items.Add(item)
                self._heading_blocks.append(block)
            host = self.FindName("ContentsHost")
            if host is not None:
                host.Visibility = (Visibility.Visible if self._heading_blocks
                                   else Visibility.Collapsed)
        finally:
            self._suppress_contents = False

    # -- navigation ------------------------------------------------------

    def _navigate(self, target):
        """Follow a link. Relative .md links move within the guides."""
        if not target:
            return
        if target.startswith(("http://", "https://", "mailto:")):
            try:
                os.startfile(target)
            except Exception as exc:
                _LOGGER.debug("Could not open %s: %s", target, exc)
            return
        clean = target.split("#", 1)[0]
        if not clean:
            return
        base = os.path.dirname(self._current or "")
        candidates = [os.path.join(base, clean)]
        if self._docs_root:
            candidates.append(os.path.join(self._docs_root, clean))
        if not clean.lower().endswith(".md"):
            candidates = [c + ".md" for c in candidates] + candidates
        for candidate in candidates:
            if os.path.exists(candidate):
                self.load(candidate)
                return
        _LOGGER.debug("Docs link went nowhere: %s", target)

    def _contents_selected(self, sender, args):
        if self._suppress_contents:
            return
        try:
            index = int(self._contents.SelectedIndex)
        except Exception:
            return
        if 0 <= index < len(self._heading_blocks):
            try:
                self._heading_blocks[index].BringIntoView()
            except Exception as exc:
                _LOGGER.debug("Could not scroll to heading: %s", exc)

    # -- buttons ---------------------------------------------------------

    def _back_clicked(self, sender, args):
        if not self._history:
            return
        previous = self._history.pop()
        self._current = None            # do not re-push what we came from
        self.load(previous, push_history=False)
        if self._back is not None:
            self._back.IsEnabled = bool(self._history)

    def _open_source_clicked(self, sender, args):
        if not self._current:
            return
        try:
            os.startfile(self._current)
        except Exception as exc:
            _LOGGER.warning("Could not open %s: %s", self._current, exc)

    def _tour_clicked(self, sender, args):
        try:
            from UIClasses.tutorial import show_tour

            self.Close()
            show_tour(self._caller_file)
        except Exception as exc:
            _LOGGER.warning("Could not start the tour: %s", exc)

    def _close_clicked(self, sender, args):
        self.Close()


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------

def show(caller_file, modal=False):
    """Open the guide for the tool that owns ``caller_file``."""
    from UIClasses.tutorial import docs

    entry, root, bundle = docs.resolve_document(caller_file)
    path = None
    if entry is not None and root:
        candidate = os.path.join(root, str(entry.get("path", "")))
        if os.path.exists(candidate):
            path = candidate
    if path is None:
        return None

    window = DocsWindow(path, entry=entry, docs_root=root, caller_file=caller_file)
    if modal:
        window.ShowDialog()
    else:
        window.Show()
    return window
