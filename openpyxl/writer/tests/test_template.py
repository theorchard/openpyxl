# Copyright (c) 2010-2015 openpyxl


from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.reader.workbook import read_content_types
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.xml.constants import XLTM, XLTX, XLSM, XLSX


def check_content_type(workbook_type, archive):
    assert workbook_type in dict(read_content_types(archive))


@pytest.mark.parametrize('tmpl, is_template', [
    ('empty.xlsx', False),
    ('empty.xlsm', False),
    ('empty.xltx', True),
    ('empty.xltm', True)
])
def test_workbook_is_template(datadir, tmpl, is_template):
    datadir.chdir()

    wb = load_workbook(tmpl)
    assert wb.is_template is is_template


@pytest.mark.parametrize('tmpl, wb_type', [
    ('empty.xlsx', XLSX),
    ('empty.xlsm', XLSM),
    ('empty.xltx', XLTX),
    ('empty.xltm', XLTM)
])
def test_xl_content_type(datadir, tmpl, wb_type):
    datadir.chdir()

    check_content_type(wb_type, ZipFile(tmpl))


@pytest.mark.parametrize('tmpl, keep_vba, wb_type', [
    ('empty.xlsx', False, XLSX),
    ('empty.xlsm', True, XLSM),
    ('empty.xltx', False, XLSX),
    ('empty.xltm', True, XLSM)
])
def test_save_xl_as_no_template(datadir, tmpl, keep_vba, wb_type):
    datadir.chdir()

    wb = save_virtual_workbook(load_workbook(tmpl, keep_vba=keep_vba),
                               as_template=False)
    check_content_type(wb_type, ZipFile(BytesIO(wb)))


@pytest.mark.parametrize('tmpl, keep_vba, wb_type', [
    ('empty.xlsx', False, XLTX),
    ('empty.xlsm', True, XLTM),
    ('empty.xltx', False, XLTX),
    ('empty.xltm', True, XLTM)
])
def test_save_xl_as_template(datadir, tmpl, keep_vba, wb_type):
    datadir.chdir()

    wb = save_virtual_workbook(load_workbook(tmpl, keep_vba=keep_vba),
                               as_template=True)
    check_content_type(wb_type, ZipFile(BytesIO(wb)))
