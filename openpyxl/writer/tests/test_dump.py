from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest

from openpyxl.compat import file
from openpyxl.tests.helper import compare_xml

from openpyxl.collections import IndexedList
from openpyxl.workbook import CALENDAR_WINDOWS_1900
from openpyxl.styles import Style


class DummyLocalData:

    pass


class DummyWorkbook:

    def __init__(self):
        self.shared_strings = IndexedList()
        self.shared_styles = IndexedList()
        self.shared_styles.add(Style())
        self._local_data = DummyLocalData()
        self.encoding = "UTF-8"
        self.excel_base_date = CALENDAR_WINDOWS_1900

    def get_sheet_names(self):
        return []


@pytest.fixture
def DumpWorksheet():
    from .. dump_worksheet import DumpWorksheet
    return DumpWorksheet(DummyWorkbook(), "TestWorkSheet")


def test_ctor(DumpWorksheet):
    ws = DumpWorksheet
    assert isinstance(ws._parent, DummyWorkbook)
    assert ws.title == "TestWorkSheet"
    assert ws._max_col == 0
    assert ws._max_row == 0
    assert hasattr(ws, '_fileobj_header_name')
    assert hasattr(ws, '_fileobj_content_name')
    assert hasattr(ws, '_fileobj_name')
    assert ws._comments == []


def test_cache(DumpWorksheet):
    ws = DumpWorksheet
    assert ws._descriptors_cache == {}


def test_filename(DumpWorksheet):
    ws = DumpWorksheet
    assert ws.filename == ws._fileobj_name


def test_get_temporary_file(DumpWorksheet):
    ws = DumpWorksheet
    fobj = ws.get_temporary_file(ws.filename)
    assert fobj.mode == "r+"


def test_get_content_generator(DumpWorksheet):
    from xml.sax.saxutils import XMLGenerator
    ws = DumpWorksheet
    doc = ws._get_content_generator()
    assert isinstance(doc, XMLGenerator)


def test_dimensions(DumpWorksheet):
    ws = DumpWorksheet
    assert ws.get_dimensions() == 'A1'


def test_write_header(DumpWorksheet):
    ws = DumpWorksheet
    doc = ws.write_header()
    header = ws.get_temporary_file(ws._fileobj_header_name)
    header.seek(0)
    xml = header.read()
    xml += "</worksheet>"
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


def test_append_data(DumpWorksheet):
    ws = DumpWorksheet
    doc = ws.append([1])
    assert ws._max_col == 1
    assert ws._max_row == 1
    content = ws.get_temporary_file(ws._fileobj_content_name)
    content.seek(0)
    xml = content.read()
    expected = """<row r="1" spans="1:1"><c r="A1" t="n"><v>1</v></c></row>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("row", ("string", dict()))
def test_invalid_append(DumpWorksheet, row):
    ws = DumpWorksheet
    with pytest.raises(TypeError):
        ws.append(row)


def test_write_only_cell():
    from .. dump_worksheet import WriteOnlyCell
    c = WriteOnlyCell()
    assert c.parent is None
    assert c.value is None
    assert c.column == 'A'
    assert c.row == 1


def test_append_cell(DumpWorksheet):
    from .. dump_worksheet import WriteOnlyCell
    ws = DumpWorksheet
    cell = WriteOnlyCell(ws, "Hello there")
    assert ws.parent.shared_strings == []
    cell._style = 5
    doc = ws.append([cell])
    assert ws.parent.shared_strings == ['Hello there']
    content = ws.get_temporary_file(ws._fileobj_content_name)
    content.seek(0)
    xml = content.read()
    expected = """<row r="1" spans="1:1"><c r="A1" t="s" s="5"><v>0</v></c></row>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_close_content(DumpWorksheet):
    ws = DumpWorksheet
    ws._close_content()
    content = open(ws._fileobj_content_name).read()
    expected = "</sheetData></worksheet>"
    content == expected


def test_close(DumpWorksheet):
    ws = DumpWorksheet
    ws.write_header()
    header = ws.get_temporary_file(ws._fileobj_header_name)
    header.write("<sheetData>") # well-formed XML needed
    ws.close()
    with open(ws.filename) as content:
        xml = content.read()
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
    </sheetPr>
    <dimension ref="A1:A1"/>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection activeCell="A1" sqref="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15" baseColWidth="10"/>
    <sheetData/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_cleanup(DumpWorksheet):
    ws = DumpWorksheet
    ws.close()
    ws._cleanup()
    assert ws._fileobj_header_name is None
    assert ws._fileobj_content_name is None
    assert ws._fileobj_name is None


def test_cannot_save_twice(DumpWorksheet):
    from .. dump_worksheet import WorkbookAlreadySaved
    ws = DumpWorksheet
    fname = ws.filename
    ws.close()
    ws._cleanup()
    with pytest.raises(WorkbookAlreadySaved):
        ws.get_temporary_file(fname)
