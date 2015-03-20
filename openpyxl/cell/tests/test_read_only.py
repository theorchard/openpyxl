from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import datetime
import pytest

from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.styles.styleable import StyleId


@pytest.fixture(scope='module')
def dummy_sheet():
    class DummyWorkbook(object):
        shared_styles = IndexedList()
        shared_styles.add(None) # Workbooks always have a default style
        _cell_styles = IndexedList()
        _cell_styles.add(None)
        _number_formats = IndexedList()


    class DummySheet(object):
        base_date = 2415018.5
        style_table = {}
        shared_strings = ['Hello world']
        parent = DummyWorkbook()
    return DummySheet()


def test_ctor(dummy_sheet):
    cell = ReadOnlyCell(dummy_sheet, None, None, 10, 'n')
    assert cell.value == 10


def test_empty_cell(dummy_sheet):
    from openpyxl.cell.read_only import EMPTY_CELL
    assert EMPTY_CELL.value is None
    assert EMPTY_CELL.data_type == 'n'


def test_base_date(dummy_sheet):
    cell = ReadOnlyCell(dummy_sheet, None, None, 10, 'n')
    assert cell.base_date == 2415018.5


def test_string_table(dummy_sheet):
    cell = ReadOnlyCell(dummy_sheet, None, None, 0, 's')
    assert cell.shared_strings == ['Hello world']
    assert cell.value == 'Hello world'


def test_coordinate(dummy_sheet):
    cell = ReadOnlyCell(dummy_sheet, 1, "A", 10, None)
    assert cell.coordinate == "A1"
    cell = ReadOnlyCell(dummy_sheet, None, None, 1, None)
    with pytest.raises(AttributeError):
        cell.coordinate


@pytest.mark.parametrize("value, expected",
                         [
                         ('1', True),
                         ('0', False),
                         ])
def test_bool(dummy_sheet, value, expected):
    cell = ReadOnlyCell(dummy_sheet, None, None, value, 'b')
    assert cell.value is expected


def test_inline_String(dummy_sheet):
    cell = ReadOnlyCell(dummy_sheet, None, None, "Hello World!", 'inlineStr')
    assert cell.value == "Hello World!"


@pytest.mark.parametrize("value, expected",
                         [
                             ("24555", 24555),
                             ("1.5", 1.5),
                             (None, None),
                         ])
def test_numeric(dummy_sheet, value, expected):
    cell = ReadOnlyCell(dummy_sheet, None, None, value, 'n')
    v = cell.value
    assert v == expected
    assert hasattr(v, 'is_integer') == hasattr(expected, 'is_integer'),\
           "Expected {0}, {1}".format(type(expected), type(v))


@pytest.fixture(scope="class")
def DummyCell(dummy_sheet):

    dummy_sheet.parent._number_formats.add('d-mmm-yy')
    style = StyleId(numFmtId=164)
    dummy_sheet.parent._cell_styles.add(style)
    cell = ReadOnlyCell(dummy_sheet, None, None, "23596", 'n', 1)
    return cell


class TestDateTime:

    def test_number_format(self, DummyCell):
        assert DummyCell.number_format == 'd-mmm-yy'

    def test_is_date(self, DummyCell):
        assert DummyCell.is_date is True

    def test_conversion(self, DummyCell):
        assert DummyCell.value == datetime.datetime(1964, 8, 7, 0, 0, 0)

    def test_interal_value(self, DummyCell):
        assert DummyCell.internal_value == 23596


def test_read_only():
    cell = ReadOnlyCell(sheet=None, row=None, column=None, value=1)
    assert cell.value == 1
    with pytest.raises(AttributeError):
        cell.value = 10
    with pytest.raises(AttributeError):
        cell.style_id = 1


def test_equality():
    c1 = ReadOnlyCell(None, None, 10, None)
    c2 = ReadOnlyCell(None, None, 10, None)
    assert c1 is not c2
    assert c1 == c2
    c3 = ReadOnlyCell(None, None, 5, None)
    assert c3 != c1
