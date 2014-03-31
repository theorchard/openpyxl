# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

# Python stdlib imports
from datetime import datetime, time, timedelta
from tempfile import NamedTemporaryFile
import os
import os.path

import pytest

from openpyxl.workbook import Workbook
from openpyxl.writer import dump_worksheet
from openpyxl.cell import get_column_letter
from openpyxl.reader.excel import load_workbook
from openpyxl.compat import xrange
from openpyxl.exceptions import WorkbookAlreadySaved


def _get_test_filename():

    test_file = NamedTemporaryFile(mode='w', prefix='openpyxl.', suffix='.xlsx', delete=False)
    test_file.close()
    return test_file.name

def test_dump_sheet_title():

    test_filename = _get_test_filename()
    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet(title='Test1')
    wb.save(test_filename)
    wb2 = load_workbook(test_filename)
    ws = wb2.get_sheet_by_name('Test1')
    assert 'Test1' == ws.title


def test_dump_string_table():
    test_filename = _get_test_filename()
    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    letters = [get_column_letter(x + 1) for x in xrange(10)]
    expected_rows = []

    for row in xrange(5):
        ws.append(['%s%d' % (letter, row + 1) for letter in letters])
    table = list(wb.shared_strings)
    assert table == ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1',
                     'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2',
                     'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3',
                     'A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4',
                     'A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5', 'J5',
                     ]


def test_dump_sheet_with_styles():
    test_filename = _get_test_filename()
    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    letters = [get_column_letter(x + 1) for x in xrange(20)]
    expected_rows = []

    for row in xrange(20):
        expected_rows.append(['%s%d' % (letter, row + 1) for letter in letters])

    for row in xrange(20):
        expected_rows.append([(row + 1) for letter in letters])

    for row in xrange(10):
        expected_rows.append([datetime(2010, ((x % 12) + 1), row + 1) for x in range(len(letters))])

    for row in xrange(20):
        expected_rows.append(['=%s%d' % (letter, row + 1) for letter in letters])

    for row in expected_rows:
        ws.append(row)

    wb.save(test_filename)
    wb2 = load_workbook(test_filename)
    ws = wb2.worksheets[0]

    for ex_row, ws_row in zip(expected_rows[:-20], ws.rows):
        for ex_cell, ws_cell in zip(ex_row, ws_row):
            assert ex_cell == ws_cell.value
    os.remove(test_filename)


def test_open_too_many_files():
    test_filename = _get_test_filename()
    wb = Workbook(optimized_write=True)
    for i in range(200): # over 200 worksheets should raise an OSError ('too many open files')
        wb.create_sheet()
    wb.save(test_filename)
    os.remove(test_filename)

def test_create_temp_file():
    f = dump_worksheet.create_temporary_file()
    if not os.path.isfile(f):
        raise Exception("The file %s does not exist" % f)

def test_dump_twice():
    test_filename = _get_test_filename()

    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    ws.append(['hello'])

    wb.save(test_filename)
    os.remove(test_filename)
    with pytest.raises(WorkbookAlreadySaved):
        wb.save(test_filename)

def test_append_after_save():
    test_filename = _get_test_filename()

    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    ws.append(['hello'])

    wb.save(test_filename)
    os.remove(test_filename)
    with pytest.raises(WorkbookAlreadySaved):
        ws.append(['hello'])


@pytest.mark.parametrize("value",
                         [
                             time(12, 15, 30),
                             datetime.now(),
                             timedelta(1),
                         ]
                         )
def test_datetime(value):
    wb = Workbook(optimized_write=True)
    ws = wb.create_sheet()
    row = [value]
    ws.append(row)
