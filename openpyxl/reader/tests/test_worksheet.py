# Copyright (c) 2010-2014 openpyxl

import pytest

from lxml.etree import iterparse

from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.cell import Cell


@pytest.fixture
def Worksheet(Workbook):
    class DummyWorkbook:

        _guess_types = False
        data_only = False

    class DummyWorksheet:

        def __init__(self):
            self.parent = DummyWorkbook()
            self.column_dimensions = {}
            self.row_dimensions = {}
            self._styles = {}

        def __getitem__(self, value):
            return Cell(self, 'A', 1)

    return DummyWorksheet()


@pytest.fixture
def WorkSheetParser(Worksheet):
    """Setup a parser instance with an empty source"""
    from .. worksheet import WorkSheetParser
    return WorkSheetParser(Worksheet, None, {0:'a'}, {})


def test_col_width(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("complex-styles-worksheet.xml", "rb") as src:
        cols = iterparse(src, tag='{%s}col' % SHEET_MAIN_NS)
        for _, col in cols:
            parser.parse_column_dimensions(col)
    assert set(ws.column_dimensions.keys()) == set(['A', 'C', 'E', 'I', 'G'])
    assert dict(ws.column_dimensions['A']) == {'max': '1', 'min': '1',
                                               'customWidth': '1',
                                               'width': '31.1640625'}


def test_hidden_col(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("hidden_rows_cols.xml", "rb") as src:
        cols = iterparse(src, tag='{%s}col' % SHEET_MAIN_NS)
        for _, col in cols:
            parser.parse_column_dimensions(col)
    assert 'D' in ws.column_dimensions
    assert dict(ws.column_dimensions['D']) == {'customWidth': '1', 'hidden': '1', 'max': '4', 'min': '4'}


def test_hidden_col(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("hidden_rows_cols.xml", "rb") as src:
        rows = iterparse(src, tag='{%s}row' % SHEET_MAIN_NS)
        for _, row in rows:
            parser.parse_row_dimensions(row)
    assert 2 in ws.row_dimensions
    #assert dict(ws.row_dimensions[2]) == {'hidden': '1'}


def test_sheet_protection(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("protected_sheet.xml", "rb") as src:
        tree = iterparse(src, tag='{%s}sheetProtection' % SHEET_MAIN_NS)
        for _, tag in tree:
            parser.parse_sheet_protection(tag)
    assert dict(ws.protection) == {
        'autoFilter': '1', 'deleteColumns': '1',
        'deleteRows': '1', 'formatCells': '1', 'formatColumns': '1', 'formatRows':
        '1', 'insertColumns': '1', 'insertHyperlinks': '1', 'insertRows': '1',
        'objects': '0', 'password': 'DAA7', 'pivotTables': '1', 'scenarios': '0',
        'selectLockedCells': '0', 'selectUnlockedCells': '0', 'sheet': '1', 'sort':
        '1'
    }
