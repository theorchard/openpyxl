# Copyright (c) 2010-2014 openpyxl

# test imports
import pytest
from .helper import compare_xml

# package imports
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet

from openpyxl.exceptions import (
    CellCoordinatesException,
    SheetTitleException,
    InsufficientCoordinatesException,
    NamedRangeException
    )
from openpyxl.writer.worksheet import write_worksheet

class TestWorkSheetWriter(object):

    @classmethod
    def setup_class(cls):
        cls.wb = Workbook()

    def test_write_empty(self):
        ws = Worksheet(self.wb)
        xml = write_worksheet(ws, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData/>
          <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_page_margins(self):
        ws = Worksheet(self.wb)
        ws.page_margins.left = 2.0
        ws.page_margins.right = 2.0
        ws.page_margins.top = 2.0
        ws.page_margins.bottom = 2.0
        ws.page_margins.header = 1.5
        ws.page_margins.footer = 1.5
        xml = write_worksheet(ws, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData/>
          <pageMargins left="2" right="2" top="2" bottom="2" header="1.5" footer="1.5"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_merge(self):
        ws = Worksheet(self.wb)
        string_table = ['Cell A1', 'Cell B1']
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:B1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData>
            <row r="1" spans="1:2">
              <c r="A1" t="s">
                <v>0</v>
              </c>
              <c r="B1" t="s">
                <v>1</v>
              </c>
            </row>
          </sheetData>
          <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
        </worksheet>
        """

        ws.cell('A1').value = 'Cell A1'
        ws.cell('B1').value = 'Cell B1'
        xml = write_worksheet(ws, string_table)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

        ws.merge_cells('A1:B1')
        xml = write_worksheet(ws, string_table)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:B1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData>
            <row r="1" spans="1:2">
              <c r="A1" t="s">
                <v>0</v>
              </c>
              <c r="B1" t="s"/>
            </row>
          </sheetData>
          <mergeCells count="1">
            <mergeCell ref="A1:B1"/>
          </mergeCells>
          <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

        ws.unmerge_cells('A1:B1')
        xml = write_worksheet(ws, string_table)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:B1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData>
            <row r="1" spans="1:2">
              <c r="A1" t="s">
                <v>0</v>
              </c>
              <c r="B1" t="s"/>
            </row>
          </sheetData>
          <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_printer_settings(self):
        ws = Worksheet(self.wb)
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToHeight = 0
        ws.page_setup.fitToWidth = 1
        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = True
        xml = write_worksheet(ws, None)
        expected = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
            <pageSetUpPr fitToPage="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData/>
          <printOptions horizontalCentered="1" verticalCentered="1"/>
          <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
          <pageSetup orientation="landscape" paperSize="3" fitToHeight="0" fitToWidth="1"/>
        </worksheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_header_footer(self):
        ws = Worksheet(self.wb)
        ws.header_footer.left_header.text = "Left Header Text"
        ws.header_footer.center_header.text = "Center Header Text"
        ws.header_footer.center_header.font_name = "Arial,Regular"
        ws.header_footer.center_header.font_size = 6
        ws.header_footer.center_header.font_color = "445566"
        ws.header_footer.right_header.text = "Right Header Text"
        ws.header_footer.right_header.font_name = "Arial,Bold"
        ws.header_footer.right_header.font_size = 8
        ws.header_footer.right_header.font_color = "112233"
        ws.header_footer.left_footer.text = "Left Footer Text\nAnd &[Date] and &[Time]"
        ws.header_footer.left_footer.font_name = "Times New Roman,Regular"
        ws.header_footer.left_footer.font_size = 10
        ws.header_footer.left_footer.font_color = "445566"
        ws.header_footer.center_footer.text = "Center Footer Text &[Path]&[File] on &[Tab]"
        ws.header_footer.center_footer.font_name = "Times New Roman,Bold"
        ws.header_footer.center_footer.font_size = 12
        ws.header_footer.center_footer.font_color = "778899"
        ws.header_footer.right_footer.text = "Right Footer Text &[Page] of &[Pages]"
        ws.header_footer.right_footer.font_name = "Times New Roman,Italic"
        ws.header_footer.right_footer.font_size = 14
        ws.header_footer.right_footer.font_color = "AABBCC"
        xml_string = write_worksheet(ws, None)
        diff = compare_xml(xml_string, """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData/>
          <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
          <headerFooter>
            <oddHeader>&amp;L&amp;"Calibri,Regular"&amp;K000000Left Header Text&amp;C&amp;"Arial,Regular"&amp;6&amp;K445566Center Header Text&amp;R&amp;"Arial,Bold"&amp;8&amp;K112233Right Header Text</oddHeader>
            <oddFooter>&amp;L&amp;"Times New Roman,Regular"&amp;10&amp;K445566Left Footer Text_x000D_And &amp;D and &amp;T&amp;C&amp;"Times New Roman,Bold"&amp;12&amp;K778899Center Footer Text &amp;Z&amp;F on &amp;A&amp;R&amp;"Times New Roman,Italic"&amp;14&amp;KAABBCCRight Footer Text &amp;P of &amp;N</oddFooter>
          </headerFooter>
        </worksheet>
        """)
        assert diff is None, diff

        ws = Worksheet(self.wb)
        xml_string = write_worksheet(ws, None)
        diff = compare_xml(xml_string, """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheetPr>
            <outlinePr summaryRight="1" summaryBelow="1"/>
          </sheetPr>
          <dimension ref="A1:A1"/>
          <sheetViews>
            <sheetView workbookViewId="0">
              <selection sqref="A1" activeCell="A1"/>
            </sheetView>
          </sheetViews>
          <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
          <sheetData/>
          <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
        </worksheet>
        """)
        assert diff is None, diff
