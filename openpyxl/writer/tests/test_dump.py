from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from datetime import datetime, date
from tempfile import NamedTemporaryFile
import os

import pytest

from openpyxl.tests.helper import compare_xml

from openpyxl.utils.indexed_list import IndexedList
from openpyxl.utils.datetime  import CALENDAR_WINDOWS_1900
from openpyxl.styles import Style
from openpyxl.styles.fonts import Font
from openpyxl.styles.proxy import StyleId
from openpyxl.comments.comments import Comment

class DummyLocalData:

    pass


class DummyWorkbook:

    def __init__(self):
        self.shared_strings = IndexedList()
        self.shared_styles = IndexedList()
        self.shared_styles.add(Style())
        self._cell_styles = IndexedList([StyleId(0, 0, 0, 0, 0, 0)])
        self._number_formats = IndexedList()
        self._local_data = DummyLocalData()
        self.encoding = "UTF-8"
        self.excel_base_date = CALENDAR_WINDOWS_1900

    def get_sheet_names(self):
        return []


@pytest.fixture
def DumpWorksheet(request):
    from .. dump_worksheet import DumpWorksheet
    ws = DumpWorksheet(DummyWorkbook(), "TestWorkSheet")
    return ws


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
    assert fobj.mode == "rb+"


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
    ws.write_header()
    header = ws.get_temporary_file(ws._fileobj_header_name)
    header.seek(0)
    xml = header.read()
    xml += b"</worksheet>"
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


def test_append_cell(DumpWorksheet):
    ws = DumpWorksheet
    ws.append([1])
    assert ws._max_col == 1
    assert ws._max_row == 1
    content = ws.get_temporary_file(ws._fileobj_content_name)
    content.seek(0)
    xml = content.read()
    expected = """<row r="1" spans="1:1"><c r="A1" t="n"><v>1</v></c></row>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff

def test_append_cell_with_string(DumpWorksheet):
    from .. dump_worksheet import WriteOnlyCell
    ws = DumpWorksheet
    cell = WriteOnlyCell(ws, "Hello there")
    ws.append([cell])
    assert ws.parent.shared_strings == ['Hello there']
    content = ws.get_temporary_file(ws._fileobj_content_name)
    content.seek(0)
    xml = content.read()
    expected = """<row r="1" spans="1:1"><c r="A1" t="s"><v>0</v></c></row>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_append_cell_with_style(DumpWorksheet):
    from .. dump_worksheet import WriteOnlyCell
    ws = DumpWorksheet
    cell = WriteOnlyCell(ws, "Hello there")
    cell.number_format = "AB"
    ws.append([cell])
    assert ws.parent.shared_strings == ['Hello there']
    content = ws.get_temporary_file(ws._fileobj_content_name)
    content.seek(0)
    xml = content.read()
    expected = """<row r="1" spans="1:1"><c r="A1" t="s" s="1"><v>0</v></c></row>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_dirty_cell(DumpWorksheet):
    ws = DumpWorksheet
    ws.append((date(2001, 1, 1), 1))
    content = ws.get_temporary_file(ws._fileobj_content_name)
    content.seek(0)
    xml = content.read()
    expected = """
    <row r="1" spans="1:2">
      <c r="A1" t="n" s="1"><v>36892</v></c>
      <c r="B1" t="n"><v>1</v></c>
      </row>"""
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

def test_close_content(DumpWorksheet):
    ws = DumpWorksheet
    ws.close()
    content = open(ws._fileobj_content_name).read()
    expected = "</sheetData></worksheet>"
    assert content == expected


def test_dump_with_comment(DumpWorksheet):
    ws = DumpWorksheet
    from openpyxl.writer.dump_worksheet import WriteOnlyCell
    from openpyxl.comments import Comment

    user_comment = Comment(text='hello world', author='me')
    cell = WriteOnlyCell(ws, value="hello")
    cell.comment = user_comment

    ws.append([cell])
    assert user_comment in ws._comments
    ws.write_header()
    header = ws.get_temporary_file(ws._fileobj_header_name)
    header.write(b"<sheetData>") # well-formed XML needed
    ws.close()
    content = open(ws._fileobj_name).read()
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
    <sheetData>
    <row r="1" spans="1:1"><c r="A1" t="s"><v>0</v></c></row>
    </sheetData>
    <legacyDrawing r:id="commentsvml"></legacyDrawing>
    </worksheet>
    """
    diff = compare_xml(content, expected)
    assert diff is None, diff

def test_close(DumpWorksheet):
    ws = DumpWorksheet
    ws.write_header()
    header = ws.get_temporary_file(ws._fileobj_header_name)
    header.write(b"<sheetData>") # well-formed XML needed
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


#Integration tests all default to using LXML backend!

from openpyxl import Workbook, load_workbook
from openpyxl.cell import get_column_letter


@pytest.fixture
def temp_file(tmpdir, request):

    test_file = NamedTemporaryFile(mode='w', prefix='openpyxl.',
                                   suffix='.xlsx', delete=False, dir=tmpdir.dirname)
    test_file.close()
    return test_file.name


def test_dump_sheet_title(temp_file):
    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet(title='Test1')
    wb.save(temp_file)
    wb2 = load_workbook(temp_file)
    ws = wb2.get_sheet_by_name('Test1')
    assert 'Test1' == ws.title


def test_dump_string_table():
    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    letters = [get_column_letter(x + 1) for x in range(10)]

    for row in range(5):
        ws.append(['%s%d' % (letter, row + 1) for letter in letters])
    table = list(wb.shared_strings)
    assert table == ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1',
                     'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2',
                     'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3',
                     'A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4',
                     'A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5', 'J5',
                     ]


def test_dump_sheet_with_styles(temp_file):
    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    letters = [get_column_letter(x + 1) for x in range(20)]
    expected_rows = []

    for row in range(20):
        expected_rows.append(['%s%d' % (letter, row + 1) for letter in letters])

    for row in range(20):
        expected_rows.append([(row + 1) for letter in letters])

    for row in range(10):
        expected_rows.append([datetime(2010, ((x % 12) + 1), row + 1) for x in range(len(letters))])

    for row in range(20):
        expected_rows.append(['=%s%d' % (letter, row + 1) for letter in letters])

    for row in expected_rows:
        ws.append(row)

    wb.save(temp_file)
    wb2 = load_workbook(temp_file)
    ws = wb2.worksheets[0]

    for ex_row, ws_row in zip(expected_rows[:-20], ws.rows):
        for ex_cell, ws_cell in zip(ex_row, ws_row):
            assert ex_cell == ws_cell.value


def test_open_too_many_files(temp_file):
    wb = Workbook(optimized_write=True)
    for i in range(200): # over 200 worksheets should raise an OSError ('too many open files')
        wb.create_sheet()
    wb.save(temp_file)


def test_dump_with_font(temp_file):
    from openpyxl.writer.dump_worksheet import WriteOnlyCell

    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    user_font = Font(name='Courrier', size=36)
    cell = WriteOnlyCell(ws, value='hello')
    cell.font = user_font

    ws.append([cell, 3.14, None])
    assert user_font in wb._fonts


@pytest.mark.parametrize("method",
                         [
                             '__getitem__',
                             '__setitem__',
                             'cell',
                             'range',
                             'merge_cells'
                         ]
                         )
def test_illegal_method(method):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    fn = getattr(ws, method)
    with pytest.raises(NotImplementedError):
        fn()

def test_save_empty_workbook(temp_file):
    wb = Workbook(write_only=True)
    assert len(wb.worksheets) == 0
    wb.save(temp_file)

    wb = load_workbook(temp_file)
    assert len(wb.worksheets) == 1
