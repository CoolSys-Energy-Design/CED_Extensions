# -*- coding: utf-8 -*-

import importlib.util
import os
import shutil
import sys
import tempfile
import types
import unittest
import zipfile

TEST_DIR = os.path.abspath(os.path.dirname(__file__))
BUTTON_DIR = os.path.dirname(TEST_DIR)
if BUTTON_DIR not in sys.path:
    sys.path.insert(0, BUTTON_DIR)

PYREVIT_SITE_PACKAGES = r"C:\Users\Aevelina\AppData\Roaming\pyRevit-Master\site-packages"
if os.path.isdir(PYREVIT_SITE_PACKAGES) and PYREVIT_SITE_PACKAGES not in sys.path:
    sys.path.append(PYREVIT_SITE_PACKAGES)

from revision_manager_core import (
    CloudRow,
    build_report_column_options,
    build_report_rows,
    filter_clouds,
    group_report_rows,
    report_column_value,
    selected_report_columns,
)
from revision_manager_exporters import build_report_html, export_pdf, export_xlsx
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None


def _row(cloud_id, comments, placement="On Sheet", revision_id=1, sheet="E2.01"):
    return CloudRow(
        cloud_id=cloud_id,
        revision_id=revision_id,
        owner_view_id=100 + cloud_id,
        revision_number="6",
        revision_display="R6",
        revision_sort=6,
        revision_date="09/01/2026",
        revision_description="Permit Revision",
        sheet_ids=[200],
        sheet_numbers_list=[sheet],
        sheet_names_list=["LIGHTING PLAN"],
        view_name="LEVEL 1 LIGHTING",
        comments=comments,
        placement=placement,
        missing_comment=not bool(comments),
        cloud_in_view=placement != "On Sheet",
        not_on_sheet=placement == "Not on Sheet",
        revision_parameters={"Approved By": "A. Reviewer"},
        cloud_parameters={"Cloud Note": "Coordination"},
        worksharing_owner="B. Owner",
        created_by="A. Creator",
        edited_by="C. Editor",
    )


class RevisionManagerCoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="revision_manager_test_")

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_filters_search_issue_and_placement(self):
        rows = [
            _row(1, "Relocated lighting controls"),
            _row(2, "", placement="Not on Sheet", sheet="M1.02"),
        ]
        self.assertEqual([1], [row.cloud_id for row in filter_clouds(rows, search="controls")])
        self.assertEqual([2], [row.cloud_id for row in filter_clouds(rows, issue="Missing Comment")])
        self.assertEqual([2], [row.cloud_id for row in filter_clouds(rows, placement="Not on Sheet")])
        self.assertEqual("E2.01", rows[0].sheet_number_links[0].text)
        self.assertEqual(200, rows[0].sheet_number_links[0].target_id)

    def test_report_deduplicates_same_revision_sheet_comment(self):
        rows = [_row(1, "Updated panel schedule"), _row(2, "Updated panel schedule")]
        self.assertEqual(1, len(build_report_rows(rows, [1], deduplicate=True)))
        self.assertEqual(2, len(build_report_rows(rows, [1], deduplicate=False)))

    def test_comment_update_refreshes_issue_and_search_state(self):
        row = _row(2, "", placement="Not on Sheet", sheet="M1.02")
        self.assertTrue(row.missing_comment)
        self.assertIn("Missing Comment", row.issues_text)
        row.update_comment("Clarified duct routing")
        self.assertFalse(row.missing_comment)
        self.assertNotIn("Missing Comment", row.issues_text)
        self.assertIn("clarified duct routing", row.search_text)
        self.assertEqual(["Not on Sheet"], [badge.text for badge in row.issue_badges])

    def test_ui_contract_has_two_tabs_and_guarded_comment_editor(self):
        with open(os.path.join(BUTTON_DIR, "RevisionManagerWindow.xaml"), "r", encoding="utf-8") as xaml_file:
            xaml = xaml_file.read()
        with open(os.path.join(BUTTON_DIR, "script.py"), "r", encoding="utf-8") as script_file:
            script_text = script_file.read()
        self.assertIn('Header="Cloud Review"', xaml)
        self.assertIn('Header="Report Builder"', xaml)
        self.assertIn('Name="CommentBox" Style="{DynamicResource CED.Input.TextBox.ReadOnly}"', xaml)
        self.assertIn('Name="ApplyCommentButton" Content="Apply Comment" Style="{DynamicResource CED.Button.Apply}"', xaml)
        self.assertIn('Name="RevisionChecklist"', xaml)
        self.assertIn('Name="ReportColumnGrid"', xaml)
        self.assertIn('ColumnHeaderStyle="{DynamicResource CED.DataGrid.Header.Wrap}"', xaml)
        self.assertIn('AcceptsReturn="False" TextWrapping="Wrap"', xaml)
        self.assertIn('Name="IncludeLogoCheck"', xaml)
        self.assertIn('Name="RevitOwnerBox"', xaml)
        self.assertIn('Name="CreatedByBox"', xaml)
        self.assertIn('Name="ShowCheckedColumnsOnlyCheck"', xaml)
        self.assertIn('Name="PageOrientationCombo"', xaml)
        self.assertIn('Name="PreviewScroll"', xaml)
        self.assertIn('Name="PreviewPageView"', xaml)
        self.assertIn('Name="PreviewZoomSlider"', xaml)
        self.assertIn('x:Name="PreviewScaleTransform"', xaml)
        self.assertIn('<Border.LayoutTransform>', xaml)
        self.assertNotIn('Name="PreviewModeCombo"', xaml)
        self.assertNotIn('ResizeDirection="Rows" ResizeBehavior="PreviousAndNext"', xaml)
        # Selected Cloud panel is a fixed-height pane; 270 -> 286 when the Apply
        # Comment row was given its own spacing and the review strip more room.
        self.assertIn('<RowDefinition Height="286" MinHeight="286"/>', xaml)
        self.assertIn('Text="{Binding comments}" TextWrapping="Wrap"', xaml)
        self.assertIn('MouseLeftButtonDown="location_link_double_clicked"', xaml)
        self.assertIn('Style="{StaticResource ControlSectionHeader}"', xaml)
        self.assertLess(xaml.index('Content="PDF"'), xaml.index('Content="Excel"'))
        self.assertIn('Content="PDF" Style="{DynamicResource CED.Button.Base}"', xaml)
        self.assertIn('Foreground="{DynamicResource CED.Brush.AccentRed}" BorderBrush="{DynamicResource CED.Brush.AccentRed}"', xaml)
        self.assertIn('Content="Excel" Style="{DynamicResource CED.Button.Apply}"', xaml)
        self.assertNotIn('Content="Choose revisions"', xaml)
        self.assertIn('DB.Transaction(doc, "Update Revision Cloud Comment")', script_text)
        self.assertIn('DB.WorksharingUtils.GetCheckoutStatus(doc, element.Id)', script_text)
        self.assertIn('if owned_by_other:', script_text)
        self.assertIn('len(self._selected_rows()) != 1', script_text)
        self.assertIn('def _execute_navigate(self, application, payload):', script_text)
        self.assertIn('def review_column_toggled(self, sender, args):', script_text)
        self.assertIn('review_columns_json', script_text)
        self.assertIn('VerticalContentAlignment="Top" TextAlignment="Left"', xaml)
        self.assertIn('Content="Fit Height"', xaml)
        self.assertIn('Content="Fit Width"', xaml)
        self.assertIn('Name="ReportFontCombo"', xaml)
        self.assertIn('Name="SelectedColumnWidthText"', xaml)
        self.assertIn('Name="SelectedColumnWidthSlider"', xaml)
        self.assertIn('Name="SelectedColumnBoldCheck"', xaml)
        self.assertNotIn('Value="{Binding width_weight, Mode=TwoWay}"', xaml)
        self.assertGreaterEqual(xaml.count('CED.DataGrid.Display.Alternating'), 3)
        self.assertIn('Name="StaleToast"', xaml)
        self.assertIn('Name="StaleOverlay"', xaml)
        self.assertIn('EventHandler[DocumentChangedEventArgs]', script_text)
        self.assertIn('def _load_clr_system_types():', script_text)
        self.assertIn('not getattr(loaded_system, "__file__", None)', script_text)
        self.assertIn('DB.BuiltInParameter.REVISION_CLOUD_REVISION', script_text)
        self.assertNotIn('_id_value(cloud.OwnerViewId) != active_view_id', script_text)

    def test_report_columns_merge_saved_order_and_parameter_values(self):
        saved = [
            {"key": "comments", "selected": True, "bold": True},
            {"key": "cloud::Cloud Note", "selected": True, "width": 3.25},
            {"key": "placement", "selected": False},
        ]
        options = build_report_column_options(
            ["Approved By"], ["Cloud Note", "Design Option", "Type Name", "Workset"], saved_schema=saved)
        self.assertEqual("comments", options[0].key)
        self.assertEqual("cloud::Cloud Note", options[1].key)
        self.assertFalse(options[2].is_selected)
        self.assertFalse(any(option.source == "Revision" for option in options))
        self.assertFalse(any(option.parameter_name in ("Design Option", "Type Name", "Workset") for option in options))
        row = build_report_rows([_row(1, "Updated panel schedule")], [1])[0]
        selected = selected_report_columns(options)
        self.assertEqual("Updated panel schedule", report_column_value(row, selected[0]))
        self.assertEqual("Coordination", report_column_value(row, selected[1]))
        self.assertEqual(3.25, selected[1].width_weight)
        self.assertTrue(selected[0].is_bold)
        worksharing = dict((option.key, option) for option in options if option.source == "Worksharing")
        self.assertEqual(set(["created_by", "edited_by", "owned_by"]), set(worksharing.keys()))
        self.assertTrue(all(not option.is_selected for option in worksharing.values()))
        self.assertEqual("A. Creator", report_column_value(row, worksharing["created_by"]))
        self.assertEqual("C. Editor", report_column_value(row, worksharing["edited_by"]))
        self.assertEqual("B. Owner", report_column_value(row, worksharing["owned_by"]))

    def test_report_html_uses_word_safe_titles_supported_font_and_saved_widths(self):
        rows = [_row(1, "Updated panel schedule")]
        groups = group_report_rows(build_report_rows(rows, [1]))
        columns = selected_report_columns(build_report_column_options([], []))
        columns[0].width_weight = 2.75
        columns[0].is_bold = True
        html = build_report_html(
            {"project_number": "1", "client": "A", "project_name": "B", "report_date": "C"},
            groups, columns, font_name="Segoe UI", table_font_size=10.5)
        self.assertNotIn("<h1", html)
        self.assertNotIn("<h2", html)
        self.assertIn("class='report-title'", html)
        self.assertIn("class='revision-title'", html)
        self.assertIn("color:#03437B", html)
        self.assertIn("font-family:'Segoe UI'", html)
        self.assertIn("font-size:10.5pt", html)
        self.assertIn("col-bold", html)

    def test_shared_elementid_helper_accepts_native_and_numeric_boundaries(self):
        class RepresentativeElementId(object):
            InvalidElementId = None

            def __init__(self, value):
                self.Value = int(value)
                self.IntegerValue = int(value)

        RepresentativeElementId.InvalidElementId = RepresentativeElementId(-1)
        db_module = types.ModuleType("Autodesk.Revit.DB")
        db_module.ElementId = RepresentativeElementId
        autodesk_module = types.ModuleType("Autodesk")
        revit_module = types.ModuleType("Autodesk.Revit")
        system_module = types.ModuleType("System")
        system_module.Int64 = int

        saved = {}
        for name, module in (
            ("Autodesk", autodesk_module),
            ("Autodesk.Revit", revit_module),
            ("Autodesk.Revit.DB", db_module),
            ("System", system_module),
        ):
            saved[name] = sys.modules.get(name)
            sys.modules[name] = module
        try:
            helper_path = os.path.abspath(os.path.join(
                BUTTON_DIR, "..", "..", "..", "..", "..",
                "CEDLib.lib", "Snippets", "revit_helpers.py"))
            spec = importlib.util.spec_from_file_location("revision_manager_test_revit_helpers", helper_path)
            helpers = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(helpers)
            native_id = RepresentativeElementId(8421)
            self.assertEqual(8421, helpers.coerce_elementid_value(native_id))
            self.assertEqual(8421, helpers.coerce_elementid_value(8421))
            rehydrated = helpers.elementid_from_value(8421)
            self.assertIsInstance(rehydrated, RepresentativeElementId)
            self.assertEqual(8421, helpers.get_elementid_value(rehydrated))
        finally:
            for name, previous in saved.items():
                if previous is None:
                    sys.modules.pop(name, None)
                else:
                    sys.modules[name] = previous

    @unittest.skipUnless(PdfReader is not None, "pypdf is required for PDF container inspection")
    def test_pdf_and_excel_exports_are_valid_containers(self):
        rows = [_row(1, "Updated panel schedule")]
        groups = group_report_rows(build_report_rows(rows, [1], deduplicate=True))
        metadata = {
            "project_number": "CED-24018",
            "client": "H-E-B",
            "project_name": "Beaumont Store",
            "report_date": "09/01/2026",
        }
        columns = selected_report_columns(build_report_column_options(["Approved By"], ["Cloud Note"]))

        pdf_path = os.path.join(self.temp_dir, "report.pdf")
        export_pdf(pdf_path, metadata, groups, columns, orientation="portrait")
        with open(pdf_path, "rb") as pdf_file:
            contents = pdf_file.read()
        self.assertTrue(contents.startswith(b"%PDF-"))
        self.assertTrue(contents.rstrip().endswith(b"%%EOF"))
        pdf = PdfReader(pdf_path)
        self.assertEqual(612, round(float(pdf.pages[0].mediabox.width)))
        self.assertEqual(792, round(float(pdf.pages[0].mediabox.height)))

        xlsx_path = os.path.join(self.temp_dir, "report.xlsx")
        export_xlsx(xlsx_path, metadata, groups, columns)
        self.assertTrue(zipfile.is_zipfile(xlsx_path))
        with zipfile.ZipFile(xlsx_path, "r") as workbook:
            names = set(workbook.namelist())
        self.assertIn("xl/tables/table1.xml", names)
        self.assertIn("xl/worksheets/sheet2.xml", names)


if __name__ == "__main__":
    unittest.main()
