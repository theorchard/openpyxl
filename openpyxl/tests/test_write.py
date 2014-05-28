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

# stdlib imports
import decimal
from io import BytesIO
import os.path
import zipfile

import pytest

# package imports
from openpyxl.tests.helper import (
    DATADIR,
    compare_xml,
    )
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import (
    save_workbook,
    save_virtual_workbook,
    ExcelWriter
    )
from openpyxl.writer.workbook import write_workbook, write_workbook_rels
from openpyxl.writer.worksheet import write_worksheet, write_worksheet_rels
from openpyxl.writer.strings import write_string_table, create_string_table
from openpyxl.writer.styles import StyleWriter

def test_write_empty_workbook(tmpdir):
    tmpdir.chdir()
    wb = Workbook()
    dest_filename = 'empty_book.xlsx'
    save_workbook(wb, dest_filename)
    assert os.path.isfile(dest_filename)


def test_write_virtual_workbook():
    old_wb = Workbook()
    saved_wb = save_virtual_workbook(old_wb)
    new_wb = load_workbook(BytesIO(saved_wb))
    assert new_wb


def test_write_workbook_rels():
    wb = Workbook()
    content = write_workbook_rels(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'workbook.xml.rels')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_workbook():
    wb = Workbook()
    content = write_workbook(wb)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'workbook.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_string_table():
    table = ['hello', 'world', 'nice']
    content = write_string_table(table)
    reference_file = os.path.join(DATADIR, 'writer', 'expected', 'sharedStrings.xml')
    with open(reference_file) as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff

