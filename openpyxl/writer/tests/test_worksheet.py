from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import datetime
import decimal
from io import BytesIO

import pytest

from openpyxl.xml.functions import XMLGenerator, tostring
from openpyxl import Workbook

from .. styles import StyleWriter
from .. worksheet import write_worksheet

from openpyxl.tests.helper import compare_xml


class DummyWorksheet:

    def __init__(self):
        self._styles = {}
        self.column_dimensions = {}


@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.fixture
def out():
    return BytesIO()


@pytest.fixture
def doc(out):
    doc = XMLGenerator(out)
    return doc


@pytest.fixture
def write_cols():
    from .. worksheet import write_worksheet_cols
    return write_worksheet_cols


@pytest.fixture
def write_sheet_format():
    from .. worksheet import write_worksheet_format
    return write_worksheet_format


@pytest.fixture
def ColumnDimension():
    from openpyxl.worksheet.dimensions import ColumnDimension
    return ColumnDimension


def test_write_no_cols(out, doc, write_cols):
    write_cols(doc, DummyWorksheet())
    doc.endDocument()
    assert out.getvalue() == b""


def test_write_col_widths(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    worksheet.column_dimensions['A'] = ColumnDimension(width=4)
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols><col width="4" min="1" max="1" customWidth="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_cols_style(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    worksheet.column_dimensions['A'] = ColumnDimension()
    worksheet._styles['A'] = 1
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols><col max="1" min="1" style="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_lots_cols(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    from openpyxl.cell import get_column_letter
    for i in range(1, 15):
        label = get_column_letter(i)
        worksheet._styles[label] = i
        worksheet.column_dimensions[label] = ColumnDimension()
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols>
   <col max="1" min="1" style="1"></col>
   <col max="2" min="2" style="2"></col>
   <col max="3" min="3" style="3"></col>
   <col max="4" min="4" style="4"></col>
   <col max="5" min="5" style="5"></col>
   <col max="6" min="6" style="6"></col>
   <col max="7" min="7" style="7"></col>
   <col max="8" min="8" style="8"></col>
   <col max="9" min="9" style="9"></col>
   <col max="10" min="10" style="10"></col>
   <col max="11" min="11" style="11"></col>
   <col max="12" min="12" style="12"></col>
   <col max="13" min="13" style="13"></col>
   <col max="14" min="14" style="14"></col>
 </cols>
"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_formula(out, doc, datadir):
    from .. worksheet import write_worksheet_data
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F1').value = 10
    ws.cell('F2').value = 32
    ws.cell('F3').value = '=F1+F2'
    ws.cell('A4').value = '=A1+A2+A3'
    ws.formula_attributes['A4'] = {'t': 'shared', 'ref': 'A4:C4', 'si': '0'}
    ws.cell('B4').value = '='
    ws.formula_attributes['B4'] = {'t': 'shared', 'si': '0'}
    ws.cell('C4').value = '='
    ws.formula_attributes['C4'] = {'t': 'shared', 'si': '0'}

    write_worksheet_data(doc, ws, [])
    doc.endDocument()
    xml = out.getvalue()
    expected = """
    <sheetData>
      <row r="1" spans="1:6">
        <c r="F1" t="n">
          <v>10</v>
        </c>
      </row>
      <row r="2" spans="1:6">
        <c r="F2" t="n">
          <v>32</v>
        </c>
      </row>
      <row r="3" spans="1:6">
        <c r="F3">
          <f>F1+F2</f>
          <v></v>
        </c>
      </row>
      <row r="4" spans="1:6">
        <c r="A4">
          <f ref="A4:C4" si="0" t="shared">A1+A2+A3</f>
          <v></v>
        </c>
        <c r="B4">
          <f si="0" t="shared"></f>
          <v></v>
        </c>
        <c r="C4">
          <f si="0" t="shared"></f>
          <v></v>
        </c>
      </row>
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_height(out, doc, worksheet):
    from .. worksheet import write_worksheet_data

    ws = worksheet
    ws.cell('F1').value = 10
    ws.row_dimensions[ws.cell('F1').row].height = 30
    ws.row_dimensions[ws.cell('F2').row].height = 30
    ws._garbage_collect()

    write_worksheet_data(doc, ws, {})
    xml = out.getvalue()
    expected = """
     <sheetData>
       <row customHeight="1" ht="30" r="1" spans="1:6">
         <c r="F1" t="n">
           <v>10</v>
         </c>
       </row>
       <row customHeight="1" ht="30" r="2" spans="1:6"></row>
     </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_get_rows_to_write(worksheet):
    from .. worksheet import get_rows_to_write

    ws = worksheet
    ws.cell('A10').value = "test"
    ws.row_dimensions[ws.cell('A10').row].height = 30
    ws.row_dimensions[ws.cell('C2').row].height = 30
    ws._garbage_collect()

    cells_by_row = get_rows_to_write(ws)

    assert len(cells_by_row) == 2
    assert len(cells_by_row[10]) == 1
    assert len(cells_by_row[2]) == 0


def test_write_hyperlink(out, doc, worksheet):
    from .. worksheet import write_worksheet_hyperlinks

    ws = worksheet
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com"

    write_worksheet_hyperlinks(doc, ws)
    xml = out.getvalue()
    expected = """
    <hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <hyperlink display="http://test.com" r:id="rId1" ref="A1"/>
    </hyperlinks>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.xfail
@pytest.mark.pil_required
def test_write_hyperlink_image_rels(Workbook, Image, datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com/"
    i = Image( "plain.png")
    ws.add_image(i)
    raise ValueError("Resulting file is invalid")
    # TODO write integration test with duplicate relation ids then fix


@pytest.fixture
def write_worksheet_autofilter():
    from .. worksheet import write_worksheet_autofilter
    return write_worksheet_autofilter


def test_write_auto_filter(out, doc, worksheet, write_worksheet_autofilter):
    ws = worksheet
    ws.auto_filter.ref = 'A1:F1'

    write_worksheet_autofilter(doc, ws)
    xml = out.getvalue()
    expected = """<autoFilter ref="A1:F1"></autoFilter>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_auto_filter_filter_column(out, doc, worksheet, write_worksheet_autofilter):
    ws = worksheet
    ws.auto_filter.ref = 'A1:F1'
    ws.auto_filter.add_filter_column(0, ["0"], blank=True)

    write_worksheet_autofilter(doc, ws)
    xml = out.getvalue()
    expected = """
    <autoFilter ref="A1:F1">
      <filterColumn colId="0">
        <filters blank="1">
          <filter val="0"></filter>
        </filters>
      </filterColumn>
    </autoFilter>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_auto_filter_sort_condition(out, doc, worksheet, write_worksheet_autofilter):
    ws = worksheet
    ws.cell('A1').value = 'header'
    ws.cell('A2').value = 1
    ws.cell('A3').value = 0
    ws.auto_filter.ref = 'A2:A3'
    ws.auto_filter.add_sort_condition('A2:A3', descending=True)

    write_worksheet_autofilter(doc, ws)
    xml = out.getvalue()
    expected = """
    <autoFilter ref="A2:A3">
    <sortState ref="A2:A3">
      <sortCondtion descending="1" ref="A2:A3"></sortCondtion>
    </sortState>
    </autoFilter>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def write_worksheet_sheetviews():
    from .. worksheet import write_worksheet_sheetviews
    return write_worksheet_sheetviews


def test_freeze_panes_horiz(out, doc, worksheet, write_worksheet_sheetviews):
    ws = worksheet
    ws.freeze_panes = 'A4'

    write_worksheet_sheetviews(doc, ws)
    xml = out.getvalue()
    expected = """
    <sheetViews>
    <sheetView workbookViewId="0">
      <pane topLeftCell="A4" ySplit="3" state="frozen" activePane="bottomLeft"/>
      <selection activeCell="A1" pane="bottomLeft" sqref="A1"/>
    </sheetView>
    </sheetViews>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_freeze_panes_vert(out, doc, worksheet, write_worksheet_sheetviews):
    ws = worksheet
    ws.freeze_panes = 'D1'

    write_worksheet_sheetviews(doc, ws)
    xml = out.getvalue()
    expected = """
    <sheetViews>
      <sheetView workbookViewId="0">
        <pane xSplit="3" topLeftCell="D1" activePane="topRight" state="frozen"/>
        <selection pane="topRight" activeCell="A1" sqref="A1"/>
      </sheetView>
    </sheetViews>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_freeze_panes_both(out, doc, worksheet, write_worksheet_sheetviews):
    ws = worksheet
    ws.freeze_panes = 'D4'

    write_worksheet_sheetviews(doc, ws)
    xml = out.getvalue()
    expected = """
    <sheetViews>
      <sheetView workbookViewId="0">
        <pane xSplit="3" ySplit="3" topLeftCell="D4" activePane="bottomRight" state="frozen"/>
        <selection pane="topRight"/>
        <selection pane="bottomLeft"/>
        <selection pane="bottomRight" activeCell="A1" sqref="A1"/>
      </sheetView>
    </sheetViews>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff



@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="s" r="A1"><v>0</v></c>"""),
                             ("", """<c r="A1" t="s"></c>"""),
                             (None, """<c r="A1" t="s"></c>"""),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                         ])
def test_write_cell(out, doc, worksheet, value, expected):
    from .. worksheet import write_cell

    ws = worksheet
    ws['A1'] = value
    write_cell(doc, ws, ws['A1'], ['Hello'])
    doc.endDocument()
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff

def test_write_cell_string(out, doc, worksheet):
    from .. worksheet import write_cell

    ws = worksheet
    ws['A1'] = "Hello"
    write_cell(doc, ws, ws['A1'], None)
    assert ws.parent.shared_strings == ["Hello"]

def test_write_sheetdata(out, doc, worksheet):
    from .. worksheet import write_worksheet_data

    ws = worksheet
    ws['A1'] = 10
    write_worksheet_data(doc, ws, [])
    doc.endDocument()
    xml = out.getvalue()
    expected = """<sheetData><row r="1" spans="1:1"><c t="n" r="A1"><v>10</v></c></row></sheetData>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_sheet_format(out, doc, write_sheet_format, ColumnDimension):
    worksheet = DummyWorksheet()
    write_sheet_format(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<sheetFormatPr defaultRowHeight="15" baseColWidth="10"/>"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


def test_outline_format(out, doc, write_sheet_format, ColumnDimension):
    worksheet = DummyWorksheet()
    worksheet.column_dimensions['A'] = ColumnDimension(outline_level=1)
    write_sheet_format(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<sheetFormatPr defaultRowHeight="15" baseColWidth="10" outlineLevelCol="1" />"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


def test_outline_cols(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    worksheet.column_dimensions['A'] = ColumnDimension(outline_level=1)
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols><col max="1" min="1" outlineLevel="1"/></cols>"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


def test_write_empty(worksheet):
    ws = worksheet
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


def test_page_margins(worksheet):
    ws = worksheet
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


def test_merge(out, doc, worksheet):
    from .. worksheet import write_worksheet_mergecells

    ws = worksheet
    ws.cell('A1').value = 'Cell A1'
    ws.cell('B1').value = 'Cell B1'

    ws.merge_cells('A1:B1')
    write_worksheet_mergecells(doc, ws)
    xml = out.getvalue()
    expected = """
      <mergeCells count="1">
        <mergeCell ref="A1:B1"/>
      </mergeCells>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    out.truncate(0) # reset
    ws.unmerge_cells('A1:B1')
    write_worksheet_mergecells(doc, ws)
    xml = out.getvalue()
    assert xml == b""


def test_printer_settings(worksheet):
    ws = worksheet
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


def test_header_footer(out, doc, worksheet):
    ws = worksheet
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

    from .. worksheet import write_header_footer
    write_header_footer(doc, ws)
    xml = out.getvalue()
    expected = """
    <headerFooter>
      <oddHeader>&amp;L&amp;"Calibri,Regular"&amp;K000000Left Header Text&amp;C&amp;"Arial,Regular"&amp;6&amp;K445566Center Header Text&amp;R&amp;"Arial,Bold"&amp;8&amp;K112233Right Header Text</oddHeader>
      <oddFooter>&amp;L&amp;"Times New Roman,Regular"&amp;10&amp;K445566Left Footer Text_x000D_And &amp;D and &amp;T&amp;C&amp;"Times New Roman,Bold"&amp;12&amp;K778899Center Footer Text &amp;Z&amp;F on &amp;A&amp;R&amp;"Times New Roman,Italic"&amp;14&amp;KAABBCCRight Footer Text &amp;P of &amp;N</oddFooter>
    </headerFooter>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_no_header(out, doc, worksheet):
    from .. worksheet import write_header_footer
    ws = worksheet

    write_header_footer(doc, ws)
    assert out.getvalue() == b""


def test_write_pagebreaks(out, doc, worksheet):
    from .. worksheet import write_pagebreaks

    write_pagebreaks(doc, worksheet)
    assert out.getvalue() == b""


def test_data_validation(out, doc, worksheet):
    from .. worksheet import write_worksheet_datavalidations
    from openpyxl.datavalidation import DataValidation, ValidationType

    ws = worksheet
    dv = DataValidation(ValidationType.LIST, formula1='"Dog,Cat,Fish"')
    dv.add_cell(ws['A1'])
    ws.add_data_validation(dv)

    write_worksheet_datavalidations(doc, worksheet)
    xml = out.getvalue()
    expected = """
    <dataValidations count="1">
    <dataValidation allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" type="list">
      <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
      <formula2>None</formula2>
    </dataValidation>
    </dataValidations>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_vba(worksheet):
    ws = worksheet
    ws.xml_source = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheetPr codeName="Sheet1"/>
        <legacyDrawing r:id="rId2"/>
    </worksheet>
    """

    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr codeName="Sheet1">
        <outlinePr summaryBelow="1" summaryRight="1"/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection activeCell="A1" sqref="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
      <legacyDrawing r:id="rId2"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_protection(out, worksheet):
    ws = worksheet
    ws.protection.enable()
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetPr>
        <outlinePr summaryBelow="1" summaryRight="1"/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection activeCell="A1" sqref="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <sheetProtection sheet="1" objects="0" selectLockedCells="0" selectUnlockedCells="0" scenarios="0" formatCells="1" formatColumns="1" formatRows="1" insertColumns="1" insertRows="1" insertHyperlinks="1" deleteColumns="1" deleteRows="1" sort="1" autoFilter="1" pivotTables="1"/>
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_comments(out, worksheet):
    ws = worksheet
    worksheet._comment_count = 1
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr>
        <outlinePr summaryBelow="1" summaryRight="1"/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection activeCell="A1" sqref="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
      <legacyDrawing r:id="commentsvml"></legacyDrawing>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
