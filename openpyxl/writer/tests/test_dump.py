from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest

from openpyxl.compat import file
from openpyxl.tests.helper import compare_xml

class DummyLocalData:

    pass


class DummyWorkbook:

    def __init__(self):
        self.shared_strings = []
        self.shared_styles = []
        self._local_data = DummyLocalData()

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


def test_dimensions(DumpWorksheet):
    ws = DumpWorksheet
    assert ws.get_dimensions() == 'A1'


def test_write_header(DumpWorksheet):
    ws = DumpWorksheet
    doc = ws.write_header()
    doc._flush()
    xml = open(ws._fileobj_header_name).read()
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
    doc._flush()
    assert ws._max_col == 1
    assert ws._max_row == 1
    xml = open(ws._fileobj_content_name).read()
    expected = """<row r="1" spans="1:1"><c r="A1" t="n"><v>1</v></c></row>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff
