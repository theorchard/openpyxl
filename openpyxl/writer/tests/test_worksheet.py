from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import decimal
from io import BytesIO

import pytest

from openpyxl.xml.functions import XMLGenerator
from openpyxl import Workbook

from .. strings import create_string_table
from .. styles import StyleWriter
from .. workbook import write_workbook
from .. worksheet import write_worksheet, write_worksheet_rels

from openpyxl.tests.helper import compare_xml


class DummyWorksheet:

    def __init__(self):
        self._styles = {}
        self.column_dimensions = {}


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



def test_write_hidden_worksheet(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.sheet_state = ws.SHEETSTATE_HIDDEN
    ws.cell('F42').value = 'hello'
    strings = create_string_table(wb)
    content = write_worksheet(ws, strings, {})
    with open('sheet1.xml') as expected:
        diff = compare_xml(content, expected.read())
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

    write_worksheet_data(doc, ws, [], None)
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


# check style tests
def test_write_style(datadir):
    datadir.chdir()
    wb = Workbook(guess_types=True)
    ws = wb.create_sheet()
    ws.cell('F1').value = '13%'
    ws._styles['F'] = ws._styles['F1']
    styles = StyleWriter(wb).styles
    content = write_worksheet(ws, {}, styles)
    with open('sheet1_style.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_height(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F1').value = 10
    ws.row_dimensions[ws.cell('F1').row].height = 30
    content = write_worksheet(ws, {}, {})
    with open('sheet1_height.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_hyperlink(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com"
    strings = create_string_table(wb)
    content = write_worksheet(ws, strings, {})
    with open('sheet1_hyperlink.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_hyperlink_rels(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    assert 0 == len(ws.relationships)
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com/"
    assert 1 == len(ws.relationships)
    ws.cell('A2').value = "test"
    ws.cell('A2').hyperlink = "http://test2.com/"
    assert 2 == len(ws.relationships)
    content = write_worksheet_rels(ws, 1, 1)
    with open('sheet1_hyperlink.xml.rels') as expected:
        diff = compare_xml(content, expected.read())
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


def test_hyperlink_value():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').hyperlink = "http://test.com"
    assert "http://test.com" == ws.cell('A1').value
    ws.cell('A1').value = "test"
    assert "test" == ws.cell('A1').value


from .. worksheet import write_worksheet_autofilter


def test_write_auto_filter(datadir, out, doc):
    datadir.chdir()
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.cell('F42').value = 'hello'
    ws.auto_filter.ref = 'A1:F1'

    write_worksheet_autofilter(doc, ws)
    xml = out.getvalue()
    expected = """<autoFilter ref="A1:F1"></autoFilter>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    content = write_workbook(wb)
    with open('workbook_auto_filter.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_auto_filter_filter_column(datadir, out, doc):
    datadir.chdir()
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.cell('F42').value = 'hello'
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


def test_write_auto_filter_sort_condition(datadir, out, doc):
    datadir.chdir()
    wb = Workbook()
    ws = wb.worksheets[0]
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


def test_freeze_panes_horiz(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = 'hello'
    ws.freeze_panes = 'A4'
    strings = create_string_table(wb)
    content = write_worksheet(ws, strings, {})
    with open('sheet1_freeze_panes_horiz.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_freeze_panes_vert(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = 'hello'
    ws.freeze_panes = 'D1'
    strings = create_string_table(wb)
    content = write_worksheet(ws, strings, {})
    with open('sheet1_freeze_panes_vert.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

def test_freeze_panes_both(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('F42').value = 'hello'
    ws.freeze_panes = 'D4'
    strings = create_string_table(wb)
    content = write_worksheet(ws, strings, {})
    with open('sheet1_freeze_panes_both.xml') as expected:
        diff = compare_xml(content, expected.read())
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
                         ])
def test_write_cell(out, doc, value, expected):
    from .. worksheet import write_cell

    wb = Workbook()
    ws = wb.active
    ws['A1'] = value
    write_cell(doc, ws, ws['A1'], ['Hello'])
    doc.endDocument()
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_sheetdata(out, doc):
    from .. worksheet import write_worksheet_data

    wb = Workbook()
    ws = wb.active
    ws['A1'] = 10
    write_worksheet_data(doc, ws, [], None)
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
