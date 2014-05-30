from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# stdlib
import datetime
import decimal
from io import BytesIO

# package
from openpyxl import Workbook
from lxml.etree import xmlfile

# test imports
import pytest
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active

@pytest.fixture
def out():
    return BytesIO()


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
def test_write_cell(out, value, expected):
    from .. lxml_worksheet import write_cell

    wb = Workbook()
    ws = wb.active
    ws['A1'] = value
    with xmlfile(out) as xf:
        write_cell(xf, ws, ws['A1'], ["Hello"])
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def write_rows():
    from .. lxml_worksheet import write_worksheet_data
    return write_worksheet_data


def test_write_sheetdata(out, worksheet, write_rows):
    ws = worksheet
    ws['A1'] = 10
    with xmlfile(out) as xf:
        write_rows(xf, ws, [])
    xml = out.getvalue()
    expected = """<sheetData><row r="1" spans="1:1"><c t="n" r="A1"><v>10</v></c></row></sheetData>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_formula(out, worksheet, write_rows):
    ws = worksheet

    ws.cell('F1').value = 10
    ws.cell('F2').value = 32
    ws.cell('F3').value = '=F1+F2'
    ws.cell('A4').value = '=A1+A2+A3'
    ws.formula_attributes['A4'] = {'t': 'shared', 'ref': 'A4:C4', 'si': '0'}
    ws.cell('B4').value = '='
    ws.formula_attributes['B4'] = {'t': 'shared', 'si': '0'}
    ws.cell('C4').value = '='
    ws.formula_attributes['C4'] = {'t': 'shared', 'si': '0'}

    with xmlfile(out) as xf:
        write_rows(xf, ws, [])

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


def test_row_height(out, worksheet, write_rows):
    ws = worksheet
    ws.cell('F1').value = 10
    ws.row_dimensions[ws.cell('F1').row].height = 30

    with xmlfile(out) as xf:
        write_rows(xf, ws, {})
    xml = out.getvalue()
    expected = """
     <sheetData>
     <row customHeight="1" ht="30" r="1" spans="1:6">
     <c r="F1" t="n">
       <v>10</v>
     </c>
   </row>
   </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def DummyWorksheet():
    class DummyWorksheet:

        def __init__(self):
            self._styles = {}
            self.column_dimensions = {}
    return DummyWorksheet()


@pytest.fixture
def write_cols():
    from .. lxml_worksheet import write_cols
    return write_cols


@pytest.fixture
def ColumnDimension():
    from openpyxl.worksheet.dimensions import ColumnDimension
    return ColumnDimension


@pytest.mark.xfail
def test_write_no_cols(out, write_cols, DummyWorksheet):
    with xmlfile(out) as xf:
        write_cols(xf, DummyWorksheet)
    assert out.getvalue() == b""


def test_write_col_widths(out, write_cols, ColumnDimension, DummyWorksheet):
    ws = DummyWorksheet
    ws.column_dimensions['A'] = ColumnDimension(width=4)
    with xmlfile(out) as xf:
        write_cols(xf, ws)
    xml = out.getvalue()
    expected = """<cols><col width="4" min="1" max="1" customWidth="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_cols_style(out, write_cols, ColumnDimension, DummyWorksheet):
    ws = DummyWorksheet
    ws.column_dimensions['A'] = ColumnDimension()
    ws._styles['A'] = 1
    with xmlfile(out) as xf:
        write_cols(xf, ws)
    xml = out.getvalue()
    expected = """<cols><col max="1" min="1" style="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_lots_cols(out, write_cols, ColumnDimension, DummyWorksheet):
    ws = DummyWorksheet
    from openpyxl.cell import get_column_letter
    for i in range(1, 15):
        label = get_column_letter(i)
        ws._styles[label] = i
        ws.column_dimensions[label] = ColumnDimension()
    with xmlfile(out) as xf:
        write_cols(xf, ws)
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


@pytest.fixture
def write_format():
    from .. lxml_worksheet import write_format
    return write_format


def test_write_sheet_format(out, write_format, ColumnDimension, DummyWorksheet):
    with xmlfile(out) as xf:
        write_format(xf, DummyWorksheet)
    xml = out.getvalue()
    expected = """<sheetFormatPr defaultRowHeight="15" baseColWidth="10"/>"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


def test_outline_format(out, write_format, ColumnDimension, DummyWorksheet):
    worksheet = DummyWorksheet
    worksheet.column_dimensions['A'] = ColumnDimension(outline_level=1)
    with xmlfile(out) as xf:
        write_format(xf, worksheet)
    xml = out.getvalue()
    expected = """<sheetFormatPr defaultRowHeight="15" baseColWidth="10" outlineLevelCol="1" />"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


def test_outline_cols(out, write_cols, ColumnDimension, DummyWorksheet):
    worksheet = DummyWorksheet
    worksheet.column_dimensions['A'] = ColumnDimension(outline_level=1)
    with xmlfile(out) as xf:
        write_cols(xf, worksheet)
    xml = out.getvalue()
    expected = """<cols><col max="1" min="1" outlineLevel="1"/></cols>"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


@pytest.fixture
def write_autofilter():
    from .. lxml_worksheet import write_autofilter
    return write_autofilter


def test_write_auto_filter(out, worksheet, write_autofilter):
    ws = worksheet
    ws.auto_filter.ref = 'A1:F1'
    with xmlfile(out) as xf:
        write_autofilter(xf, ws)
    xml = out.getvalue()
    expected = """<autoFilter ref="A1:F1"></autoFilter>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_auto_filter_filter_column(out, worksheet, write_autofilter):
    ws = worksheet
    ws.auto_filter.ref = 'A1:F1'
    ws.auto_filter.add_filter_column(0, ["0"], blank=True)

    with xmlfile(out) as xf:
        write_autofilter(xf, ws)
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


def test_write_auto_filter_sort_condition(out, worksheet, write_autofilter):
    ws = worksheet
    ws.cell('A1').value = 'header'
    ws.cell('A2').value = 1
    ws.cell('A3').value = 0
    ws.auto_filter.ref = 'A2:A3'
    ws.auto_filter.add_sort_condition('A2:A3', descending=True)

    with xmlfile(out) as xf:
        write_autofilter(xf, ws)
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
def write_sheetviews():
    from .. lxml_worksheet import write_sheetviews
    return write_sheetviews


def test_freeze_panes_horiz(out, worksheet, write_sheetviews):
    ws = worksheet
    ws.freeze_panes = 'A4'

    with xmlfile(out) as xf:
        write_sheetviews(xf, ws)
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


def test_freeze_panes_vert(out, worksheet, write_sheetviews):
    ws = worksheet
    ws.freeze_panes = 'D1'

    with xmlfile(out) as xf:
        write_sheetviews(xf, ws)
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


def test_freeze_panes_both(out, worksheet, write_sheetviews):
    ws = worksheet
    ws.freeze_panes = 'D4'

    with xmlfile(out) as xf:
        write_sheetviews(xf, ws)
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


@pytest.fixture
def write_worksheet():
    from .. lxml_worksheet import write_worksheet
    return write_worksheet


@pytest.mark.xfail
def test_page_margins(worksheet, out):
    ws = worksheet
    ws.page_margins.left = 2.0
    ws.page_margins.right = 2.0
    ws.page_margins.top = 2.0
    ws.page_margins.bottom = 2.0
    ws.page_margins.header = 1.5
    ws.page_margins.footer = 1.5
    with xmlfile(out) as xf:
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


def test_merge(out, worksheet):
    from .. lxml_worksheet import write_mergecells

    ws = worksheet
    ws.cell('A1').value = 'Cell A1'
    ws.cell('B1').value = 'Cell B1'

    ws.merge_cells('A1:B1')
    with xmlfile(out) as xf:
        write_mergecells(xf, ws)
    xml = out.getvalue()
    expected = """
      <mergeCells count="1">
        <mergeCell ref="A1:B1"/>
      </mergeCells>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    out = BytesIO()
    ws.unmerge_cells('A1:B1')
    with xmlfile(out) as xf:
        write_mergecells(xf, ws)
    xml = out.getvalue()
    expected = """<mergeCells/>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff
