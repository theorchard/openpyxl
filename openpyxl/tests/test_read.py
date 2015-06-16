from __future__ import absolute_import
# coding=utf8
# Copyright (c) 2010-2015 openpyxl

# Python stdlib imports
from datetime import datetime
from io import BytesIO

import pytest

# compatibility imports
from openpyxl.compat import unicode

# package imports
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet import Worksheet
from openpyxl.workbook import Workbook
from openpyxl.styles import numbers, Style
from openpyxl.reader.worksheet import read_worksheet
from openpyxl.reader.excel import load_workbook
from openpyxl.utils.datetime  import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904


def test_read_standalone_worksheet(datadir):

    class DummyWb(object):

        encoding = 'utf-8'

        excel_base_date = CALENDAR_WINDOWS_1900
        _guess_types = True
        data_only = False
        vba_archive = None

        def __init__(self):
            self.shared_styles = [Style()]
            self._cell_styles = IndexedList()

        def get_sheet_by_name(self, value):
            return None

        def get_sheet_names(self):
            return []

    datadir.join("reader").chdir()
    ws = None
    shared_strings = IndexedList(['hello'])

    with open('sheet2.xml') as src:
        ws = read_worksheet(src.read(), DummyWb(), 'Sheet 2', shared_strings,
                            {1: Style()})
        assert isinstance(ws, Worksheet)
        assert ws.cell('G5').value == 'hello'
        assert ws.cell('D30').value == 30
        assert ws.cell('K9').value == 0.09


@pytest.fixture
def standard_workbook(datadir):
    datadir.join("genuine").chdir()
    return load_workbook("empty.xlsx")


def test_read_worksheet(standard_workbook):
    wb = standard_workbook
    sheet2 = wb['Sheet2 - Numbers']
    assert isinstance(sheet2, Worksheet)
    assert 'This is cell G5' == sheet2['G5'].value
    assert 18 == sheet2['D18'].value
    assert sheet2['G9'].value is True
    assert sheet2['G10'].value is False


@pytest.mark.parametrize("cell, number_format",
                    [
                        ('A1', numbers.FORMAT_GENERAL),
                        ('A2', numbers.FORMAT_DATE_XLSX14),
                        ('A3', numbers.FORMAT_NUMBER_00),
                        ('A4', numbers.FORMAT_DATE_TIME3),
                        ('A5', numbers.FORMAT_PERCENTAGE_00),
                    ]
                    )
def test_read_general_style(datadir, cell, number_format):
    datadir.join("genuine").chdir()
    wb = load_workbook('empty-with-styles.xlsx')
    ws = wb["Sheet1"]
    assert ws[cell].number_format == number_format


def test_read_no_theme(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook('libreoffice_nrt.xlsx')
    assert wb


def test_read_cell_formulae(datadir):
    from openpyxl.reader.worksheet import fast_parse
    datadir.join("reader").chdir()
    wb = Workbook()
    ws = wb.active
    fast_parse(ws, open( "worksheet_formula.xml"), ['', ''], {}, None)
    b1 = ws['B1']
    assert b1.data_type == 'f'
    assert b1.value == '=CONCATENATE(A1,A2)'
    a6 = ws['A6']
    assert a6.data_type == 'f'
    assert a6.value == '=SUM(A4:A5)'


def test_read_complex_formulae(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook('formulae.xlsx')
    ws = wb.get_active_sheet()

    # Test normal forumlae
    assert ws.cell('A1').data_type != 'f'
    assert ws.cell('A2').data_type != 'f'
    assert ws.cell('A3').data_type == 'f'
    assert 'A3' not in ws.formula_attributes
    assert ws.cell('A3').value == '=12345'
    assert ws.cell('A4').data_type == 'f'
    assert 'A4' not in ws.formula_attributes
    assert ws.cell('A4').value == '=A2+A3'
    assert ws.cell('A5').data_type == 'f'
    assert 'A5' not in ws.formula_attributes
    assert ws.cell('A5').value == '=SUM(A2:A4)'

    # Test unicode
    expected = '=IF(ISBLANK(B16), "Düsseldorf", B16)'
    # Hack to prevent pytest doing it's own unicode conversion
    try:
        expected = unicode(expected, "UTF8")
    except TypeError:
        pass
    assert ws['A16'].value == expected

    # Test shared forumlae
    assert ws.cell('B7').data_type == 'f'
    assert ws.formula_attributes['B7']['t'] == 'shared'
    assert ws.formula_attributes['B7']['si'] == '0'
    assert ws.formula_attributes['B7']['ref'] == 'B7:E7'
    assert ws.cell('B7').value == '=B4*2'
    assert ws.cell('C7').data_type == 'f'
    assert ws.formula_attributes['C7']['t'] == 'shared'
    assert ws.formula_attributes['C7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['C7']
    assert ws.cell('C7').value == '='
    assert ws.cell('D7').data_type == 'f'
    assert ws.formula_attributes['D7']['t'] == 'shared'
    assert ws.formula_attributes['D7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['D7']
    assert ws.cell('D7').value == '='
    assert ws.cell('E7').data_type == 'f'
    assert ws.formula_attributes['E7']['t'] == 'shared'
    assert ws.formula_attributes['E7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['E7']
    assert ws.cell('E7').value == '='

    # Test array forumlae
    assert ws.cell('C10').data_type == 'f'
    assert 'ref' not in ws.formula_attributes['C10']['ref']
    assert ws.formula_attributes['C10']['t'] == 'array'
    assert 'si' not in ws.formula_attributes['C10']
    assert ws.formula_attributes['C10']['ref'] == 'C10:C14'
    assert ws.cell('C10').value == '=SUM(A10:A14*B10:B14)'
    assert ws.cell('C11').data_type != 'f'


def test_data_only(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook('formulae.xlsx', data_only=True)
    ws = wb.active
    # Test cells returning values only, not formulae
    assert ws['A2'].value == 12345
    assert ws['A3'].value == 12345
    assert ws['A4'].value == 24690
    assert ws['A5'].value == 49380


@pytest.mark.parametrize("guess_types, dtype",
                         (
                             (True, float),
                             (False, unicode),
                         )
                        )
def test_guess_types(datadir, guess_types, dtype):
    datadir.join("genuine").chdir()
    wb = load_workbook('guess_types.xlsx', guess_types=guess_types)
    ws = wb.active
    assert isinstance(ws['D2'].value, dtype)
