from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest
import datetime
import decimal
from lxml.etree import tostring

from openpyxl.tests.helper import compare_xml
from openpyxl.compat import unicode

from . test_dump import DummyWorkbook

@pytest.fixture
def LXMLWorksheet():
    from .. dump_lxml import LXMLWorksheet
    return LXMLWorksheet(DummyWorkbook(), title="TestWorksheet")


def test_write_header(LXMLWorksheet):
    ws = LXMLWorksheet
    doc = ws.write_header()
    header = open(ws.filename)
    xml = header.read()
    expected = """<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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
    xml = unicode(tostring(el), "utf-8")
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_append(LXMLWorksheet):
    ws = LXMLWorksheet
    ws.append([1, "s"])
    ws.writer.close()
    with open(ws._fileobj_content_name) as rows:
        xml = rows.read()
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
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
