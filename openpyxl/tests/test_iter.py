# Copyright (c) 2010-2014 openpyxl

import datetime

import pytest

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


@pytest.mark.parametrize("filename",
                         ["sheet2.xml",
                          "sheet2_no_dimension.xml"
                         ]
                         )
def test_get_max_cell(datadir, filename):
    datadir.join("reader").chdir()

    class Workbook:
        excel_base_date = None
        _cell_styles = [None]

        def get_sheet_names(self):
            return []

    from openpyxl.worksheet.iter_worksheet import IterableWorksheet
    ws = IterableWorksheet(Workbook(), "Sheet", "", filename, [], [])
    rows = tuple(ws.rows)
    assert rows[-1][-1].coordinate == "AA30"


@pytest.fixture(params=[False, True])
def sample_workbook(request, datadir):
    """Standard and read-only workbook"""
    datadir.join("genuine").chdir()
    wb = load_workbook(filename="empty.xlsx", read_only=request.param, data_only=True)
    return wb


def test_calculate_dimension(datadir):
    """
    Behaviour differs between implementations
    """
    datadir.join("genuine").chdir()
    wb = load_workbook(filename="empty.xlsx", read_only=True)
    sheet2 = wb['Sheet2 - Numbers']
    assert sheet2.calculate_dimension() == 'D1:AA30'


@pytest.mark.parametrize("read_only",
                         [
                             False,
                             True
                         ]
                         )
def test_get_missing_cell(read_only, datadir):
    """
    Behaviour differs between implementations
    """
    datadir.join("genuine").chdir()
    wb = load_workbook(filename="empty.xlsx", read_only=read_only)
    ws = wb['Sheet2 - Numbers']
    assert (ws['A1'] is None) is read_only


def test_getitem(sample_workbook):
    wb = sample_workbook
    ws = wb['Sheet1 - Text']
    assert list(ws.iter_rows("A1"))[0][0] == ws['A1']
    assert list(ws.iter_rows("A1:D30")) == list(ws["A1:D30"])
    assert list(ws.iter_rows("A1:D30")) == list(ws["A1":"D30"])


def test_max_row(sample_workbook):
    wb = sample_workbook
    sheet2 = wb['Sheet2 - Numbers']
    assert sheet2.max_row == 30


expected = [
    ("Sheet1 - Text", 7),
    ("Sheet2 - Numbers", 27),
    ("Sheet3 - Formulas", 4),
    ("Sheet4 - Dates", 3)
             ]
@pytest.mark.parametrize("sheetname, col", expected)
def test_max_column(sample_workbook, sheetname, col):
    wb = sample_workbook
    ws = wb[sheetname]
    assert ws.max_column == col


def test_read_single_cell_range(sample_workbook):
    wb = sample_workbook
    ws = wb['Sheet1 - Text']
    assert 'This is cell A1 in Sheet 1' == list(ws.iter_rows('A1'))[0][0].value


expected = [['This is cell A1 in Sheet 1', None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, None],
            [None, None, None, None, None, None, 'This is cell G5'], ]

def test_read_fast_integrated_text(sample_workbook):
    wb = sample_workbook
    ws = wb['Sheet1 - Text']
    for row, expected_row in zip(ws.iter_rows(), expected):
        row_values = [x.value for x in row]
        assert row_values == expected_row


def test_read_single_cell_range(sample_workbook):
    wb = sample_workbook
    ws = wb['Sheet1 - Text']
    assert 'This is cell A1 in Sheet 1' == list(ws.iter_rows('A1'))[0][0].value


def test_read_fast_integrated_numbers(sample_workbook):
    wb = sample_workbook
    expected = [[x + 1] for x in range(30)]
    query_range = 'D1:D30'
    ws = wb['Sheet2 - Numbers']
    for row, expected_row in zip(ws.iter_rows(query_range), expected):
        row_values = [x.value for x in row]
        assert row_values == expected_row


def test_read_fast_integrated_numbers_2(sample_workbook):
    wb = sample_workbook
    query_range = 'K1:K30'
    expected = expected = [[(x + 1) / 100.0] for x in range(30)]
    ws = wb['Sheet2 - Numbers']
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
    ws = wb['Sheet4 - Dates']
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
    ws = wb["Sheet3 - Formulas"]
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


def test_read_style_iter(tmpdir):
    '''
    Test if cell styles are read properly in iter mode.
    '''
    tmpdir.chdir()
    from openpyxl import Workbook
    from openpyxl.styles import Style, Font

    FONT_NAME = "Times New Roman"
    FONT_SIZE = 15
    ft = Font(name=FONT_NAME, size=FONT_SIZE)

    wb = Workbook()
    ws = wb.worksheets[0]
    cell = ws.cell('A1')
    cell.font = ft

    xlsx_file = "read_only_styles.xlsx"
    wb.save(xlsx_file)

    wb_iter = load_workbook(xlsx_file, read_only=True)
    ws_iter = wb_iter.worksheets[0]
    cell = ws_iter['A1']

    assert cell.style.font == ft


def test_read_hyperlinks_read_only(datadir, Workbook):
    from openpyxl.worksheet.iter_worksheet import IterableWorksheet

    datadir.join("reader").chdir()
    filename = 'bug328_hyperlinks.xml'
    ws = IterableWorksheet(Workbook(data_only=True, read_only=True), "Sheet",
                           "", filename, ['SOMETEXT'], [])
    assert ws['F2'].value is None


def test_read_with_missing_cells(datadir):
    datadir.join("reader").chdir()

    class Workbook:
        excel_base_date = None
        _cell_styles = [None]

        def get_sheet_names(self):
            return []

    filename = "bug393-worksheet.xml"

    from openpyxl.worksheet.iter_worksheet import IterableWorksheet
    ws = IterableWorksheet(Workbook(), "Sheet", "", filename, [], [])
    row = tuple(ws._get_cells(2, 1, 2, None))
    values = [c.value for c in row]
    assert values == [None, None, 1, 2, 3]

    row = tuple(ws._get_cells(4, 1, 4, None))
    values = [c.value for c in row]
    assert values == [1, 2, None, None, 3]
