from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


import datetime
import decimal
from io import BytesIO
from lxml.etree import tostring, tounicode, xmlfile

from openpyxl.tests.helper import compare_xml

import pytest
from . test_dump import DummyWorkbook

@pytest.fixture
def LXMLWorksheet():
    from .. dump_lxml import LXMLWorksheet
    return LXMLWorksheet(DummyWorkbook(), title="TestWorksheet")


def test_write_header(LXMLWorksheet):
    ws = LXMLWorksheet
    doc = ws._write_header()
    next(doc)
    doc.close()
    header = open(ws.filename)
    xml = header.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
    <sheetData/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_close_content(LXMLWorksheet):
    pass


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
def test_write_cell(LXMLWorksheet, value, expected):
    from openpyxl.cell import Cell
    from .. dump_lxml import write_cell
    ws = LXMLWorksheet
    c = Cell(ws, 'A', 1, value)
    el = write_cell(ws, c)
    xml = tounicode(el)
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_append(LXMLWorksheet):
    ws = LXMLWorksheet

    def _writer(doc):
        with xmlfile(doc) as xf:
            with xf.element('sheetData'):
                try:
                    while True:
                        body = (yield)
                        xf.write(body)
                except GeneratorExit:
                    pass

    doc = BytesIO()
    ws.writer = _writer(doc)
    next(ws.writer)

    ws.append([1, "s"])
    ws.append(['2', 3])
    ws.writer.close()
    xml = doc.getvalue()
    expected = """
    <sheetData>
      <row r="1" spans="1:2">
        <c r="A1" t="n">
          <v>1</v>
        </c>
        <c r="B1" t="s">
          <v>0</v>
        </c>
      </row>
      <row r="2" spans="1:2">
        <c r="A2" t="s">
          <v>1</v>
        </c>
        <c r="B2" t="n">
          <v>3</v>
        </c>
      </row>
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_cell_comment(LXMLWorksheet):
    ws = LXMLWorksheet
    from openpyxl.comments import Comment
    from .. dump_worksheet import WriteOnlyCell
    cell = WriteOnlyCell(ws, 1)
    comment = Comment('hello', 'me')
    cell.comment = comment
    ws.append([cell])
    assert ws._comments == [comment]


def test_cannot_save_twice(LXMLWorksheet):
    from .. dump_worksheet import WorkbookAlreadySaved

    ws = LXMLWorksheet
    ws.close()
    with pytest.raises(WorkbookAlreadySaved):
        ws.close()
    with pytest.raises(WorkbookAlreadySaved):
        ws.append([1])


def test_close(LXMLWorksheet):
    ws = LXMLWorksheet
    ws.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
    <sheetData/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
