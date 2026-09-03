# -*- coding: utf-8 -*-
"""File and clipboard exporters for Revision Cloud Manager."""

import base64
import os
import shutil
import subprocess
import tempfile
import time

from revision_manager_core import plain_text_report, report_column_value, text


def _html_escape(value):
    return (text(value).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _logo_data_uri(logo_path):
    if not logo_path or not os.path.isfile(logo_path):
        return ""
    with open(logo_path, "rb") as logo_file:
        encoded = base64.b64encode(logo_file.read())
    if not isinstance(encoded, str):
        encoded = encoded.decode("ascii")
    return "data:image/png;base64,{}".format(encoded)


def _column_weight(column):
    try:
        return max(0.5, min(4.0, float(column.width_weight)))
    except Exception:
        pass
    if column.key == "comments":
        return 2.4
    if column.key == "placement":
        return 1.35
    if column.key == "issues":
        return 1.35
    if column.key == "sheet_number":
        return 1.0
    return 1.35


def _supported_font(value):
    requested = text(value, "Arial")
    supported = ("Arial", "Calibri", "Segoe UI", "Times New Roman")
    return requested if requested in supported else "Arial"


def build_report_html(metadata, groups, columns, logo_path=None, pdf_layout=False,
                      orientation="landscape", font_name="Arial", table_font_size=10.5):
    page_orientation = "portrait" if text(orientation).lower() == "portrait" else "landscape"
    font_name = _supported_font(font_name)
    try:
        table_font_size = max(8.0, min(14.0, float(table_font_size)))
    except Exception:
        table_font_size = 10.5
    page_rule = "@page{{size:Letter {};margin:0.45in;}}".format(page_orientation) if pdf_layout else ""
    parts = [
        "<html><head><meta charset='utf-8'><style>",
        page_rule,
        "*{box-sizing:border-box;}body{font-family:'" + font_name + "',sans-serif;color:#1F2E3D;font-size:9.5pt;margin:0;}",
        ".report-title{font-size:18pt;font-weight:bold;color:#03437B;margin:0 0 4px;}",
        ".revision-title{font-size:11.5pt;font-weight:bold;color:#03437B;margin:16px 0 6px;}",
        "p.meta{margin:2px 0;}table{border-collapse:collapse;width:100%;margin:0 0 12px;table-layout:fixed;}",
        "thead{display:table-header-group;}tr{break-inside:avoid;page-break-inside:avoid;}",
        "th{background:#E7EDF5;text-align:left;font-weight:bold;}",
        "th,td{border:1px solid #AAB4BF;padding:4px 5px;vertical-align:top;text-align:left;white-space:pre-wrap;word-wrap:break-word;font-size:" + str(table_font_size) + "pt;}",
        "td.col-comments{text-align:left;}td.col-issues{color:#9D1825;}td.col-bold{font-weight:bold;}",
        ".portrait{font-size:9pt;}.portrait th,.portrait td{padding:3px 4px;}",
        ".muted{color:#4A5866;}.report-header{display:flex;align-items:flex-start;justify-content:space-between;margin:0 0 8px;}",
        ".logo{display:block;max-width:155px;max-height:44px;margin:0 0 0 18px;object-fit:contain;}",
        ".revision-section{break-inside:auto;page-break-inside:auto;}",
        # A revision heading stranded at the foot of a page with its table overleaf
        # reads as an empty revision, so keep the heading with the rows it introduces.
        # (Row splitting itself is already prevented by the tr rule above.)
        ".revision-title{break-after:avoid;page-break-after:avoid;}",
        "</style></head><body class='{}'><!--StartFragment-->".format(page_orientation),
    ]
    logo_uri = _logo_data_uri(logo_path)
    parts.extend([
        "<div class='report-header'><div>",
        "<div class='report-title'>Project Revision Summary</div>",
        "<p class='muted'><strong>CoolSys Energy Design</strong></p>",
        "</div>",
    ])
    if logo_uri:
        parts.append("<img class='logo' alt='CoolSys Energy Design' src='{}'/>".format(logo_uri))
    parts.append("</div>")
    for label, key in (
        ("Project Number", "project_number"),
        ("Client", "client"),
        ("Project Name", "project_name"),
        ("Report Date", "report_date"),
    ):
        parts.append("<p class='meta'><strong>{}:</strong> {}</p>".format(label, _html_escape(metadata.get(key))))
    for group in groups:
        parts.append("<section class='revision-section'>")
        parts.append("<div class='revision-title'>Revision {} | Date: {} | Description: {}</div>".format(
            _html_escape(group["number"]), _html_escape(group["date"]), _html_escape(group["description"])))
        parts.append("<table><colgroup>")
        total_weight = sum([_column_weight(column) for column in columns]) or 1.0
        for column in columns:
            percent = 100.0 * _column_weight(column) / total_weight
            parts.append("<col style='width:{:.3f}%'/>".format(percent))
        parts.append("</colgroup><thead><tr>")
        for column in columns:
            parts.append("<th>{}</th>".format(_html_escape(column.label)))
        parts.append("</tr></thead><tbody>")
        if not group["rows"]:
            parts.append("<tr><td colspan='{}'>No revision clouds found for this revision.</td></tr>".format(
                max(1, len(columns))))
        for row in group["rows"]:
            parts.append("<tr>")
            for column in columns:
                css_class = "col-{}".format(column.key.replace("::", "-").replace("_", "-"))
                if bool(getattr(column, "is_bold", False)):
                    css_class += " col-bold"
                parts.append("<td class='{}'>{}</td>".format(
                    css_class, _html_escape(report_column_value(row, column))))
            parts.append("</tr>")
        parts.append("</tbody></table></section>")
    parts.append("<!--EndFragment--></body></html>")
    return "".join(parts)


def _cf_html(html):
    """Wrap HTML with Windows CF_HTML byte offsets."""
    marker_start = "<!--StartFragment-->"
    marker_end = "<!--EndFragment-->"
    header_template = (
        "Version:0.9\r\n"
        "StartHTML:{:010d}\r\n"
        "EndHTML:{:010d}\r\n"
        "StartFragment:{:010d}\r\n"
        "EndFragment:{:010d}\r\n"
    )
    empty_header = header_template.format(0, 0, 0, 0)
    start_html = len(empty_header.encode("utf-8"))
    html_bytes = html.encode("utf-8")
    start_fragment = start_html + html_bytes.find(marker_start.encode("ascii")) + len(marker_start)
    end_fragment = start_html + html_bytes.find(marker_end.encode("ascii"))
    end_html = start_html + len(html_bytes)
    return header_template.format(start_html, end_html, start_fragment, end_fragment) + html


def copy_report_to_clipboard(metadata, groups, columns, logo_path=None, font_name="Arial", table_font_size=10.5):
    from System.Windows import Clipboard, DataFormats, DataObject

    # Word should receive report content only. The logo is intentionally excluded
    # even when it is enabled for the PDF and Excel exports.
    html = build_report_html(
        metadata, groups, columns, logo_path=None,
        font_name=font_name, table_font_size=table_font_size)
    plain = plain_text_report(metadata, groups, columns)
    data = DataObject()
    data.SetData(DataFormats.Html, _cf_html(html))
    data.SetData(DataFormats.UnicodeText, plain)
    data.SetData(DataFormats.Text, plain)
    Clipboard.SetDataObject(data, True)
    return len(plain)


def export_xlsx(path, metadata, groups, columns, logo_path=None, font_name="Arial"):
    import xlsxwriter

    workbook = xlsxwriter.Workbook(path)
    font_name = _supported_font(font_name)
    title_format = workbook.add_format({"bold": True, "font_size": 18, "font_color": "#03437B", "font_name": font_name})
    label_format = workbook.add_format({"bold": True, "font_color": "#1F2E3D"})
    section_format = workbook.add_format({"bold": True, "font_size": 12, "font_color": "#03437B", "font_name": font_name})
    bold_value_format = workbook.add_format({"bold": True, "font_name": font_name})

    summary = workbook.add_worksheet("Revision Summary")
    summary.hide_gridlines(2)
    summary.set_column("A:A", 19)
    summary.set_column("B:B", 42)
    summary.write("A1", "Project Revision Summary", title_format)
    if logo_path and os.path.isfile(logo_path):
        try:
            summary.insert_image("D1", logo_path, {"x_scale": 0.35, "y_scale": 0.35})
        except Exception:
            pass
    row_index = 2
    for label, key in (
        ("Project Number", "project_number"),
        ("Client", "client"),
        ("Project Name", "project_name"),
        ("Report Date", "report_date"),
    ):
        summary.write(row_index, 0, label, label_format)
        summary.write(row_index, 1, text(metadata.get(key)))
        row_index += 1
    row_index += 1
    summary.write(row_index, 0, "Revision", label_format)
    summary.write(row_index, 1, "Description", label_format)
    summary.write(row_index, 2, "Date", label_format)
    summary.write(row_index, 3, "Cloud Rows", label_format)
    row_index += 1
    for group in groups:
        summary.write(row_index, 0, group["number"])
        summary.write(row_index, 1, group["description"])
        summary.write(row_index, 2, group["date"])
        summary.write_number(row_index, 3, len(group["rows"]))
        row_index += 1

    details = workbook.add_worksheet("Revision Clouds")
    details.freeze_panes(1, 0)
    headers = [column.label for column in columns]
    detail_rows = []
    for group in groups:
        for row in group["rows"]:
            detail_rows.append([report_column_value(row, column) for column in columns])
    for column_index, header in enumerate(headers):
        details.write(0, column_index, header)
    for output_row, values in enumerate(detail_rows, 1):
        for column_index, value in enumerate(values):
            value_format = bold_value_format if bool(getattr(columns[column_index], "is_bold", False)) else None
            details.write(output_row, column_index, value, value_format)
    for column_index, column in enumerate(columns):
        width = max(10, min(60, int(round(18.0 * _column_weight(column)))))
        details.set_column(column_index, column_index, width)
    if detail_rows and headers:
        details.add_table(0, 0, len(detail_rows), len(headers) - 1, {
            "name": "RevisionCloudTable",
            "style": "Table Style Medium 2",
            "columns": [{"header": item} for item in headers],
        })
    else:
        details.write(2, 0, "No revision cloud rows matched the report selection.", section_format)
    workbook.close()
    return path


def _find_chromium_browser():
    candidates = []
    for env_name in ("PROGRAMFILES(X86)", "PROGRAMFILES", "LOCALAPPDATA"):
        base = os.environ.get(env_name)
        if not base:
            continue
        candidates.extend([
            os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"),
        ])
    candidates.extend([
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ])
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return ""


def _file_url(path):
    return "file:///{}".format(os.path.abspath(path).replace("\\", "/").replace(" ", "%20"))


def export_pdf(path, metadata, groups, columns, logo_path=None, timeout_seconds=45,
               orientation="landscape", font_name="Arial", table_font_size=10.5):
    """Render the report HTML to PDF through locally installed Edge or Chrome."""
    browser = _find_chromium_browser()
    if not browser:
        raise Exception("Microsoft Edge or Google Chrome is required for PDF export.")
    output_path = os.path.abspath(path)
    work_dir = tempfile.mkdtemp(prefix="ced_revision_pdf_")
    html_path = os.path.join(work_dir, "revision-report.html")
    profile_path = os.path.join(work_dir, "browser-profile")
    try:
        html = build_report_html(
            metadata, groups, columns, logo_path=logo_path, pdf_layout=True,
            orientation=orientation, font_name=font_name,
            table_font_size=table_font_size)
        with open(html_path, "wb") as html_file:
            html_file.write(html.encode("utf-8"))
        arguments = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-pdf-header-footer",
            "--user-data-dir={}".format(profile_path),
            "--print-to-pdf={}".format(output_path),
            _file_url(html_path),
        ]
        devnull = open(os.devnull, "wb")
        try:
            creation_flags = 0x08000000 if os.name == "nt" else 0
            process = subprocess.Popen(
                arguments, stdout=devnull, stderr=devnull, creationflags=creation_flags)
            started = time.time()
            while process.poll() is None and time.time() - started < float(timeout_seconds):
                time.sleep(0.1)
            if process.poll() is None:
                process.kill()
                raise Exception("PDF export timed out while waiting for the browser renderer.")
            if process.returncode != 0:
                raise Exception("The browser PDF renderer exited with code {}.".format(process.returncode))
        finally:
            devnull.close()
        if not os.path.isfile(output_path) or os.path.getsize(output_path) < 1000:
            raise Exception("The browser renderer did not create a valid PDF file.")
        with open(output_path, "rb") as pdf_file:
            if pdf_file.read(5) != b"%PDF-":
                raise Exception("The exported file is not a valid PDF.")
        return output_path
    finally:
        try:
            shutil.rmtree(work_dir)
        except Exception:
            pass
