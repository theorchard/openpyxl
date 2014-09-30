# coding: utf-8

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

from io import BytesIO

from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import save_virtual_workbook


def test_write_workbook_code_name(datadir):
    datadir.join('genuine').chdir()

    wb = load_workbook('empty.xlsx')
    wb = load_workbook(BytesIO(save_virtual_workbook(wb)))
    assert wb.code_name == u'ThisWorkbook'

    # This file contains a macros that should run when you open a workbook
    wb = load_workbook('empty_wb_russian_code_name.xlsm', keep_vba=True)
    wb = load_workbook(BytesIO(save_virtual_workbook(wb)), keep_vba=True)
    assert wb.code_name == u'\u042d\u0442\u0430\u041a\u043d\u0438\u0433\u0430'