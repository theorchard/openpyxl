from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# stdlib
import datetime
import decimal
from io import BytesIO

# package
from openpyxl import Workbook
from lxml.etree import xmlfile, tostring

# test imports
import pytest
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="s" r="A1"><v>0</v></c>"""),
                             ("", """<c r="A1" t="s"></c>"""),
                             (None, """<c r="A1" t="n"></c>"""),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                         ])
def test_write_cell(value, expected):
    from .. lxml_worksheet import write_cell

    wb = Workbook()
    ws = wb.active
    ws['A1'] = value

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, ws['A1'])
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff

def test_write_cell_string(worksheet):
    from .. lxml_worksheet import write_cell

    ws = worksheet
    ws['A1'] = "Hello"

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, ws['A1'])
    assert ws.parent.shared_strings == ["Hello"]


@pytest.fixture
def write_rows():
    from .. lxml_worksheet import write_rows
    return write_rows


def test_write_sheetdata(worksheet, write_rows):
    ws = worksheet
    ws['A1'] = 10

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)
    xml = out.getvalue()
    expected = """<sheetData><row r="1" spans="1:1"><c t="n" r="A1"><v>10</v></c></row></sheetData>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_formula(worksheet, write_rows):
    ws = worksheet

    ws.cell('F1').value = 10
    ws.cell('F2').value = 32
    ws.cell('F3').value = '=F1+F2'
    ws.cell('A4').value = '=A1+A2+A3'
    ws.formula_attributes['A4'] = {'t': 'shared', 'ref': 'A4:C4', 'si': '0'}
    ws.cell('B4').value = '=1'
    ws.formula_attributes['B4'] = {'t': 'shared', 'si': '0'}
    ws.cell('C4').value = '=1'
    ws.formula_attributes['C4'] = {'t': 'shared', 'si': '0'}

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)

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


def test_row_height(worksheet, write_rows):
    ws = worksheet
    ws.cell('F1').value = 10
    ws.row_dimensions[ws.cell('F1').row].height = 30
    ws.row_dimensions[ws.cell('F2').row].height = 30
    ws._garbage_collect()

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)
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



@pytest.fixture
def write_worksheet():
    from .. lxml_worksheet import write_worksheet
    return write_worksheet


@pytest.mark.lxml_required
def test_empty_worksheet(worksheet, write_worksheet):
    xml = write_worksheet(worksheet, None)
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
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.lxml_required
def test_printer_settings(worksheet, write_worksheet):
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
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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


@pytest.mark.lxml_required
def test_page_margins(worksheet, write_worksheet):
    ws = worksheet
    ws.page_margins.left = 2.0
    ws.page_margins.right = 2.0
    ws.page_margins.top = 2.0
    ws.page_margins.bottom = 2.0
    ws.page_margins.header = 1.5
    ws.page_margins.footer = 1.5
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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


@pytest.mark.lxml_required
def test_vba(worksheet, write_worksheet):
    ws = worksheet
    ws.vba_code = {"codeName":"Sheet1"}
    ws.vba_controls = "rId2"

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


@pytest.mark.lxml_required
def test_protection(worksheet, write_worksheet):
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


@pytest.mark.lxml_required
def test_write_comments(worksheet, write_worksheet):
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
