# Copyright (c) 2010-2014 openpyxl

import datetime
import os.path

import pytest

from openpyxl.tests.helper import DATADIR
from openpyxl.worksheet.iter_worksheet import read_dimension
from openpyxl.reader.excel import load_workbook
from openpyxl.compat import range, zip


def test_open_many_sheets(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook("bigfoot.xlsx", True) # if
    assert len(wb.worksheets) == 1024


@pytest.mark.parametrize("filename, expected",
                         [
                             ("sheet2.xml", ('D', 1, 'AA', 30)),
                             ("sheet2_no_dimension.xml", None),
                             ("sheet2_no_span.xml", None),
                          ]
                         )
def test_read_dimension(datadir, filename, expected):
    datadir.join("reader").chdir()
    with open(filename) as handle:
        dimension = read_dimension(handle)
    assert dimension == expected


def test_calculate_dimension(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook("empty.xlsx", use_iterators=True)
    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    dimensions = sheet2.calculate_dimension()
    assert '%s%s:%s%s' % ('D', 1, 'AA', 30) == dimensions


def test_get_highest_row(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook("empty.xlsx", use_iterators=True)
    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    max_row = sheet2.get_highest_row()
    assert 30 == max_row


@pytest.fixture
def sample_workbook(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook(filename="empty.xlsx", use_iterators=True,
                       data_only=True)
    return wb


class TestWorksheet(object):

    workbook_name = os.path.join(DATADIR, 'genuine', 'empty.xlsx')

    def _open_wb(self, data_only=False):
        return load_workbook(filename=self.workbook_name,
                             use_iterators=True,
                             data_only=data_only)

def test_getitem(sample_workbook):
    wb = sample_workbook
    ws = wb['Sheet1 - Text']
    assert list(ws.iter_rows("A1"))[0][0] == ws['A1']
    assert list(ws.iter_rows("A1:D30")) == list(ws["A1:D30"])
    assert list(ws.iter_rows("A1:D30")) == list(ws["A1":"D30"])

    ws = wb['Sheet2 - Numbers']
    assert ws['A1'] is None


expected = [
    ("Sheet1 - Text", 'A1:G5'),
    ("Sheet2 - Numbers", 'D1:AA30'),
    ("Sheet3 - Formulas", 'D2:D2'),
    ("Sheet4 - Dates", 'A1:C1')
                 ]
@pytest.mark.parametrize("sheetname, dims", expected)
def test_get_dimensions(sample_workbook, sheetname, dims):
    wb = sample_workbook
    ws = wb[sheetname]
    assert ws.dimensions == dims

expected = [
    ("Sheet1 - Text", 7),
    ("Sheet2 - Numbers", 27),
    ("Sheet3 - Formulas", 4),
    ("Sheet4 - Dates", 3)
             ]
@pytest.mark.parametrize("sheetname, col", expected)
def test_get_highest_column_iter(sample_workbook, sheetname, col):
    wb = sample_workbook
    ws = wb[sheetname]
    assert ws.get_highest_column() == col


expected = [['This is cell A1 in Sheet 1', None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, 'This is cell G5'], ]
def test_read_fast_integrated(sample_workbook):
    wb = sample_workbook
    ws = wb.get_sheet_by_name('Sheet1 - Text')
    for row, expected_row in zip(ws.iter_rows(), self.expected):
        row_values = [x.value for x in row]
        assert row_values == expected_row


def test_read_single_cell_range(sample_workbook):
    wb = sample_workbook
    ws = wb.get_sheet_by_name('Sheet1 - Text')
    assert 'This is cell A1 in Sheet 1' == list(ws.iter_rows('A1'))[0][0].value


def test_read_fast_integrated(sample_workbook):
    wb = sample_workbook
    sheet_name = 'Sheet2 - Numbers'
    expected = [[x + 1] for x in range(30)]
    query_range = 'D1:D30'
    ws = wb.get_sheet_by_name(name = sheet_name)
    for row, expected_row in zip(ws.iter_rows(query_range), expected):
        row_values = [x.value for x in row]
        assert row_values == expected_row


def test_read_fast_integrated(sample_workbook):
    wb = sample_workbook
    sheet_name = 'Sheet2 - Numbers'
    query_range = 'K1:K30'
    expected = expected = [[(x + 1) / 100.0] for x in range(30)]
    ws = wb.get_sheet_by_name(name = sheet_name)
    for row, expected_row in zip(ws.iter_rows(query_range), expected):
        row_values = [x.value for x in row]
        assert row_values == expected_row


@pytest.mark.parametrize("cell, value",
    [
    ("A1", datetime.datetime(1973, 5, 20)),
    ("C1", datetime.datetime(1973, 5, 20, 9, 15, 2))
    ]
    )
def test_read_single_cell_date(sample_workbook, cell, value):
    wb = sample_workbook
    ws = wb.get_sheet_by_name('Sheet4 - Dates')
    rows = ws.iter_rows(cell)
    cell = list(rows)[0][0]
    assert cell.value == value


@pytest.mark.parametrize("data_only, expected",
    [
    (True, 5),
    (False, "='Sheet2 - Numbers'!D5")
    ]
    )
def test_read_single_cell_formula(datadir, data_only, expected):
    datadir.join("genuine").chdir()
    wb = load_workbook("empty.xlsx", read_only=True, data_only=data_only)
    ws = wb.get_sheet_by_name("Sheet3 - Formulas")
    rows = ws.iter_rows("D2")
    cell = list(rows)[0][0]
    assert ws.parent.data_only == data_only
    assert cell.value == expected


@pytest.mark.parametrize("cell, expected",
    [
    ("G9", True),
    ("G10", False)
    ]
    )
def test_read_boolean(sample_workbook, cell, expected):
    wb = sample_workbook
    ws = wb["Sheet2 - Numbers"]
    row = list(ws.iter_rows(cell))
    assert row[0][0].coordinate == cell
    assert row[0][0].data_type == 'b'
    assert row[0][0].value == expected


def test_read_style_iter():
    '''
    Test if cell styles are read properly in iter mode.
    '''
    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Style, Font

    FONT_NAME = "Times New Roman"
    FONT_SIZE = 15

    wb = Workbook()
    ws = wb.worksheets[0]
    cell = ws.cell('A1')
    cell.style = Style(font=Font(name=FONT_NAME, size=FONT_SIZE))

    xlsx_file = tempfile.NamedTemporaryFile()
    wb.save(xlsx_file)

    # Passes as of 1.6.1
    wb_regular = load_workbook(xlsx_file)
    ws_regular = wb_regular.worksheets[0]
    cell_style_regular = ws_regular.cell('A1').style
    assert cell_style_regular.font.name == FONT_NAME
    assert cell_style_regular.font.size == FONT_SIZE

    # Fails as of 1.6.1
    # perhaps not correct
    # but would work if style_table was not ignored and styles
    # would still be present in ws_iter._styles
    wb_iter = load_workbook(xlsx_file, use_iterators=True)
    ws_iter = wb_iter.worksheets[0]
    cell = ws_iter['A1']

    assert cell.style.font.name == FONT_NAME
    assert cell.style.font.size == FONT_SIZE
