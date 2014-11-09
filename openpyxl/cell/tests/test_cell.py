from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#

# Python stdlib imports
from datetime import time, datetime, timedelta, date
import decimal

# 3rd party imports
import pytest

# package imports

from openpyxl.comments import Comment
from ..cell import Cell


@pytest.fixture
def build_dummy_worksheet():
    from openpyxl.utils.collections import IndexedList
    from openpyxl.utils.datetime  import CALENDAR_WINDOWS_1900
    from openpyxl.cell import Cell
    from openpyxl.styles import Style

    class Wb(object):
        excel_base_date = CALENDAR_WINDOWS_1900
        shared_styles = IndexedList([Style()])


    class Ws(object):

        encoding = 'utf-8'
        parent = Wb()
        title = "Dummy Worksheet"
        _comment_count = 0

        def cell(self, column, row):
            return Cell(self, column, row)

    return Ws()


@pytest.fixture
def dummy_cell(request):
    ws = build_dummy_worksheet()
    cell = ws.cell('A', 1)
    return cell


@pytest.fixture(params=[True, False])
def guess_types(request):
    return request.param


@pytest.mark.parametrize("value, expected",
                         [
                             ('4.2', 4.2),
                             ('-42.000', -42),
                             ( '0', 0),
                             ('0.9999', 0.9999),
                             ('99E-02', 0.99),
                             ('4', 4),
                             ('-1E3', -1000),
                             ('2e+2', 200),
                             ('3.1%', 0.031),
                             ('03:40:16', time(3, 40, 16)),
                             ('03:40', time(3, 40)),
                             ('30:33.865633336', time(0, 30, 33, 865633))
                         ]
                        )
def test_infer_numeric(dummy_cell, guess_types, value, expected):
    cell = dummy_cell
    cell.parent.parent._guess_types = guess_types
    cell.value = value
    if cell.guess_types:
        assert cell.value == expected
    else:
        cell.value == value


def test_ctor(dummy_cell):
    cell = dummy_cell
    assert cell.data_type == 'n'
    assert cell.column == 'A'
    assert cell.row == 1
    assert cell.coordinate == "A1"
    assert cell.value is None
    assert cell.xf_index == 0
    assert cell.comment is None


@pytest.mark.parametrize("datatype", ['n', 'd', 's', 'b', 'f', 'e'])
def test_null(dummy_cell, datatype):
    cell = dummy_cell
    cell.data_type = datatype
    assert cell.data_type == datatype
    cell.value = None
    assert cell.data_type == 'n'


@pytest.mark.parametrize("value", ['hello', ".", '0800'])
def test_string(dummy_cell, value):
    cell = dummy_cell
    cell.value = 'hello'
    assert cell.data_type == 's'


@pytest.mark.parametrize("value", ['=42', '=if(A1<4;-1;1)'])
def test_formula(dummy_cell, value):
    cell = dummy_cell
    cell.value = value
    assert cell.data_type == 'f'


@pytest.mark.parametrize("value", [True, False])
def test_boolean(dummy_cell, value):
    cell = dummy_cell
    cell.value = True
    assert cell.data_type == 'b'


@pytest.mark.parametrize("error_string", Cell.ERROR_CODES)
def test_error_codes(dummy_cell, error_string):
    cell = dummy_cell
    cell.value = error_string
    assert cell.data_type == 'e'


@pytest.mark.parametrize("value, internal, number_format",
                         [
                             (
                                 datetime(2010, 7, 13, 6, 37, 41),
                                 40372.27616898148,
                                 "yyyy-mm-dd h:mm:ss"
                             ),
                             (
                                 date(2010, 7, 13),
                                 40372,
                                 "yyyy-mm-dd"
                             ),
                             (
                                 time(1, 3),
                                 0.04375,
                                 "h:mm:ss",
                             )
                         ]
                         )
def test_insert_date(dummy_cell, value, internal, number_format):
    cell = dummy_cell
    cell.value = value
    assert cell.data_type == 'n'
    assert cell.internal_value == internal
    assert cell.is_date
    assert cell.number_format == number_format


@pytest.mark.parametrize("value, is_date",
                         [
                             (None, True,),
                             ("testme", False),
                             (True, False),
                         ]
                         )
def test_cell_formatted_as_date(dummy_cell, value, is_date):
    cell = dummy_cell
    cell.value = datetime.today()
    cell.value = value
    assert cell.is_date == is_date
    assert cell.value == value


def test_set_bad_type(dummy_cell):
    cell = dummy_cell
    with pytest.raises(ValueError):
        cell.set_explicit_value(1, 'q')


def test_illegal_chacters(dummy_cell):
    from openpyxl.exceptions import IllegalCharacterError
    from openpyxl.compat import range
    from itertools import chain
    cell = dummy_cell
    #ws = build_dummy_worksheet()
    #cell = Cell(ws, 'A', 1)

    # The bytes 0x00 through 0x1F inclusive must be manually escaped in values.

    illegal_chrs = chain(range(9), range(11, 13), range(14, 32))
    for i in illegal_chrs:
        with pytest.raises(IllegalCharacterError):
            cell.value = chr(i)

        with pytest.raises(IllegalCharacterError):
            cell.value = "A {0} B".format(chr(i))

    cell.value = chr(33)
    cell.value = chr(9)  # Tab
    cell.value = chr(10)  # Newline
    cell.value = chr(13)  # Carriage return
    cell.value = " Leading and trailing spaces are legal "


values = (
    ('30:33.865633336', [('', '', '', '30', '33', '865633')]),
    ('03:40:16', [('03', '40', '16', '', '', '')]),
    ('03:40', [('03', '40', '',  '', '', '')]),
    ('55:72:12', []),
    )
@pytest.mark.parametrize("value, expected",
                             values)
def test_time_regex(value, expected):
    from openpyxl.cell.cell import TIME_REGEX
    m = TIME_REGEX.findall(value)
    assert m == expected


def test_timedelta(dummy_cell):
    cell = dummy_cell
    cell.value = timedelta(days=1, hours=3)
    assert cell.value == 1.125
    assert cell.data_type == 'n'
    assert cell.is_date is False
    assert cell.number_format == "[hh]:mm:ss"


def test_repr(dummy_cell):
    cell = dummy_cell
    assert repr(cell), '<Cell Sheet1.A1>' == 'Got bad repr: %s' % repr(cell)


def test_comment_count(dummy_cell):
    cell = dummy_cell
    ws = cell.parent
    assert ws._comment_count == 0
    cell.comment = Comment("text", "author")
    assert ws._comment_count == 1
    cell.comment = Comment("text", "author")
    assert ws._comment_count == 1
    cell.comment = None
    assert ws._comment_count == 0
    cell.comment = None
    assert ws._comment_count == 0

def test_comment_assignment(dummy_cell):
    cell = dummy_cell
    ws = cell.parent
    c = Comment("text", "author")
    cell.comment = c
    with pytest.raises(AttributeError):
        ws.cell(column='A', row=2).commment = c
    ws.cell(column='A', row=2).comment = Comment("text2", "author2")
    # this should orphan c, so that assigning it to A2 does not raise AttributeError
    cell.comment = None
    ws.cell(column='A', row=2).comment = c

def test_cell_offset(dummy_cell):
    cell = dummy_cell
    ws = cell.parent
    assert cell.offset(2, 1).coordinate == 'B3'


class TestEncoding:

    try:
        # Python 2
        pound = unichr(163)
    except NameError:
        # Python 3
        pound = chr(163)
    test_string = ('Compound Value (' + pound + ')').encode('latin1')

    def test_bad_encoding(self):
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        cell = ws['A1']
        with pytest.raises(UnicodeDecodeError):
            cell.check_string(self.test_string)
        with pytest.raises(UnicodeDecodeError):
            cell.value = self.test_string

    def test_good_encoding(self):
        from openpyxl import Workbook

        wb = Workbook(encoding='latin1')
        ws = wb.active
        cell = ws['A1']
        cell.value = self.test_string


def test_style(dummy_cell):
    from openpyxl.styles import Font, Style
    cell = dummy_cell
    new_style = Style(font=Font(bold=True))
    cell.style = new_style
    assert new_style in cell.parent.parent.shared_styles
