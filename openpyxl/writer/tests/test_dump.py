from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import datetime
import decimal
from io import BytesIO

from openpyxl.xml.functions import tostring, xmlfile

from openpyxl.utils.indexed_list import IndexedList
from openpyxl.utils.datetime  import CALENDAR_WINDOWS_1900

from openpyxl.styles import Style
from openpyxl.styles.styleable import StyleId

from openpyxl.tests.helper import compare_xml

import pytest


class DummyLocalData:

    pass


class DummyWorkbook:

    def __init__(self):
        self.shared_strings = IndexedList()
        self.shared_styles = [Style()]
        self._cell_styles = IndexedList([StyleId(0, 0, 0, 0, 0, 0)])
        self._number_formats = IndexedList()
        self._local_data = DummyLocalData()
        self.encoding = "UTF-8"
        self.excel_base_date = CALENDAR_WINDOWS_1900

    def get_sheet_names(self):
        return []



@pytest.fixture
def DumpWorksheet():
    from .. dump_worksheet import DumpWorksheet
    return DumpWorksheet(DummyWorkbook(), title="TestWorksheet")


@pytest.mark.lxml_required
def test_write_header(DumpWorksheet):
    ws = DumpWorksheet
    doc = ws._write_header()
    next(doc)
    doc.close()
    header = open(ws.filename)
    xml = header.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
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




def test_append(DumpWorksheet):
    ws = DumpWorksheet

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
    ws.append(i for i in [1, 2])
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
      <row r="3" spans="1:2">
        <c r="A3" t="n">
          <v>1</v>
        </c>
        <c r="B3" t="n">
          <v>2</v>
        </c>
      </row>
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_dirty_cell(DumpWorksheet):
    ws = DumpWorksheet

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

    ws.append((datetime.date(2001, 1, 1), 1))
    ws.writer.close()
    xml = doc.getvalue()
    expected = """
    <sheetData>
    <row r="1" spans="1:2">
      <c r="A1" t="n" s="1"><v>36892</v></c>
      <c r="B1" t="n"><v>1</v></c>
      </row>
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("row", ("string", dict()))
def test_invalid_append(DumpWorksheet, row):
    ws = DumpWorksheet
    with pytest.raises(TypeError):
        ws.append(row)


@pytest.mark.lxml_required
def test_cell_comment(DumpWorksheet):
    ws = DumpWorksheet
    from openpyxl.comments import Comment
    from .. dump_worksheet import WriteOnlyCell
    cell = WriteOnlyCell(ws, 1)
    comment = Comment('hello', 'me')
    cell.comment = comment
    ws.append([cell])
    assert ws._comments == [comment]
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
    <sheetData>
    <row r="1" spans="1:1"><c r="A1" t="n"><v>1</v></c></row>
    </sheetData>
    <legacyDrawing r:id="commentsvml"></legacyDrawing>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.lxml_required
def test_cannot_save_twice(DumpWorksheet):
    from .. dump_worksheet import WorkbookAlreadySaved

    ws = DumpWorksheet
    ws.close()
    with pytest.raises(WorkbookAlreadySaved):
        ws.close()
    with pytest.raises(WorkbookAlreadySaved):
        ws.append([1])


@pytest.mark.lxml_required
def test_close(DumpWorksheet):
    ws = DumpWorksheet
    ws.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
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

@pytest.mark.lxml_required
def test_auto_filter(DumpWorksheet):
    ws = DumpWorksheet
    ws.auto_filter.ref = 'A1:F1'
    ws.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
    <sheetData/>
    <autoFilter ref="A1:F1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

@pytest.mark.lxml_required
def test_frozen_panes(DumpWorksheet):
    ws = DumpWorksheet
    ws.freeze_panes = 'D4'
    ws.close()
    with open(ws.filename) as src:
        xml = src.read()
    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <pane xSplit="3" ySplit="3" topLeftCell="D4" activePane="bottomRight" state="frozen"/>
        <selection pane="topRight"/>
        <selection pane="bottomLeft"/>
        <selection pane="bottomRight" activeCell="A1" sqref="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
    <sheetData/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.lxml_required
def test_write_empty_row(DumpWorksheet):
    ws = DumpWorksheet
    ws.append(['1', '2', '3'])
    ws.append([])
    ws.close()

    with open(ws.filename) as src:
        xml = src.read()

    expected = """
    <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
      <pageSetUpPr/>
    </sheetPr>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
    <sheetData>
    <row r="1" spans="1:3">
      <c r="A1" t="s">
        <v>0</v>
      </c>
      <c r="B1" t="s">
        <v>1</v>
      </c>
      <c r="C1" t="s">
        <v>2</v>
      </c>
    </row>
    <row r="2"/>
    </sheetData>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
