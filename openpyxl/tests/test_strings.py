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
import os.path

# package imports
from openpyxl.workbook import Workbook
from openpyxl.writer.strings import create_string_table
from openpyxl.reader.strings import read_string_table

import pytest

def test_create_string_table():
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('B12').value = 'hello'
    ws.cell('B13').value = 'world'
    ws.cell('D28').value = 'hello'
    table = create_string_table(wb)
    assert table == ['hello', 'world']


@pytest.fixture
def reader_dir(datadir):
    return datadir.join('reader')


def test_read_string_table(reader_dir):
    src = str(reader_dir.join('sharedStrings.xml'))
    with open(src) as content:
        assert read_string_table(content.read()) == [
            'This is cell A1 in Sheet 1', 'This is cell G5']


def test_empty_string(reader_dir):
    src = str(reader_dir.join('sharedStrings-emptystring.xml'))
    with open(src) as content:
        assert read_string_table(content.read()) == ['Testing empty cell', '']


def test_formatted_string_table(reader_dir):
    src = str(reader_dir.join('shared-strings-rich.xml'))
    with open(src) as content:
        assert read_string_table(content.read()) == [
            'Welcome', 'to the best shop in town', "     let's play "]
