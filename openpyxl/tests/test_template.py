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
from zipfile import ZipFile

from openpyxl.reader.workbook import read_content_types
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.reader.excel import load_workbook


def check_content_type(workbook_type, archive):
    cts = {name: ct for ct, name in read_content_types(archive)}

    assert workbook_type in cts['/xl/workbook.xml']


def check_sheet_content_type(archive):
    check_content_type('.sheet.', archive)


def check_template_content_type(archive):
    check_content_type('.template.', archive)


def test_xl_content_type(datadir):
    datadir.join('genuine').chdir()

    check_sheet_content_type(ZipFile('empty.xlsx'))
    check_sheet_content_type(ZipFile('empty.xlsm'))

    check_template_content_type(ZipFile('empty.xltx'))
    check_template_content_type(ZipFile('empty.xltm'))


def test_save_xl_as_no_template(datadir):
    datadir.join('genuine').chdir()

    wb = save_virtual_workbook(load_workbook('empty.xlsx'), as_template=False)
    check_sheet_content_type(ZipFile(BytesIO(wb)))

    wb = save_virtual_workbook(load_workbook('empty.xlsm'), as_template=False)
    check_sheet_content_type(ZipFile(BytesIO(wb)))

    wb = save_virtual_workbook(load_workbook('empty.xltx'), as_template=False)
    check_sheet_content_type(ZipFile(BytesIO(wb)))

    wb = save_virtual_workbook(load_workbook('empty.xltm'), as_template=False)
    check_sheet_content_type(ZipFile(BytesIO(wb)))


def test_save_xl_as_template(datadir):
    datadir.join('genuine').chdir()

    wb = save_virtual_workbook(load_workbook('empty.xlsx'), as_template=True)
    check_template_content_type(ZipFile(BytesIO(wb)))

    wb = save_virtual_workbook(load_workbook('empty.xlsm'), as_template=True)
    check_template_content_type(ZipFile(BytesIO(wb)))

    wb = save_virtual_workbook(load_workbook('empty.xltx'), as_template=True)
    check_template_content_type(ZipFile(BytesIO(wb)))

    wb = save_virtual_workbook(load_workbook('empty.xltm'), as_template=True)
    check_template_content_type(ZipFile(BytesIO(wb)))
