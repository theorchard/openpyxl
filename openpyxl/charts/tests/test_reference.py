# Copyright (c) 2010-2014 openpyxl

import pytest

@pytest.fixture
def Workbook():
    from openpyxl.workbook import Workbook
    return Workbook


@pytest.fixture
def Reference():
    from openpyxl.charts.reference import Reference
    return Reference

@pytest.fixture
def ws(Workbook):
    """Empty worksheet titled 'data'"""
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def ten_row_sheet(ws):
    """Worksheet with values 0-9 in the first column"""
    for i in range(10):
        ws.append([i])
    return ws


@pytest.fixture
def sheet(ten_row_sheet):
    ten_row_sheet.title = "reference"
    return ten_row_sheet


@pytest.fixture
def cell(sheet, Reference):
    return Reference(sheet, (1, 1))


@pytest.fixture
def cell_range(sheet, Reference):
    return Reference(sheet, (1, 1), (10, 1))


@pytest.fixture()
def empty_range(sheet, Reference):
    for i in range(10):
        sheet.cell(row=i+1, column=2).value = None
    return Reference(sheet, (1, 2), (10, 2))


@pytest.fixture()
def missing_values(sheet, Reference):
    vals = [None, None, 1, 2, 3, 4, 5, 6, 7, 8]
    for idx, val in enumerate(vals, 1):
        sheet.cell(row=idx, column=3).value = val
    return Reference(sheet, (1, 3), (10, 3))


@pytest.fixture
def column_of_letters(sheet, Reference):
    for idx, l in enumerate("ABCDEFGHIJ", 1):
        sheet.cell(row=idx, column=2).value = l
    return Reference(sheet, (1, 2), (10, 2))


class TestReference(object):

    def test_single_cell_ctor(self, cell):
        assert cell.pos1 == (1, 1)
        assert cell.pos2 == None

    def test_range_ctor(self, cell_range):
        assert cell_range.pos1 == (1, 1)
        assert cell_range.pos2 == (10, 1)

    def test_single_cell_ref(self, cell):
        assert cell.values == [0]
        assert str(cell) == "'reference'!$A$1"

    def test_cell_range_ref(self, cell_range):
        assert cell_range.values == [0, 1, 2, 3, 4, 5, 6, 7, 8 , 9]
        assert str(cell_range) == "'reference'!$A$1:$A$10"

    def test_data_type(self, cell):
        with pytest.raises(ValueError):
            cell.data_type = 'f'
            cell.data_type = None

    def test_type_inference(self, cell, cell_range, column_of_letters,
                            missing_values):
        assert cell.values == [0]
        assert cell.data_type == 'n'

        assert cell_range.values == [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
        assert cell_range.data_type == 'n'

        assert column_of_letters.values == list("ABCDEFGHIJ")
        assert column_of_letters.data_type == "s"

        assert missing_values.values == ['', '', 1, 2, 3, 4, 5, 6, 7, 8]
        missing_values.values
        assert missing_values.data_type == 'n'

    def test_number_format(self, cell):
        with pytest.raises(ValueError):
            cell.number_format = 'YYYY'
        cell.number_format = 'd-mmm'
        assert cell.number_format == 'd-mmm'

