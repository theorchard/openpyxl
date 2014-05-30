# Copyright (c) 2010-2014 openpyxl

# Python stdlib imports
from zipfile import ZipFile
from datetime import datetime

# test imports
import pytest

# package imports
from openpyxl.tests.helper import compare_xml
from openpyxl.reader.workbook import read_properties_core
from openpyxl.writer.workbook import (
    write_properties_core,
    write_properties_app
)
from openpyxl.xml.constants import ARC_CORE
from openpyxl.date_time import CALENDAR_WINDOWS_1900
from openpyxl.workbook import DocumentProperties, Workbook


def test_read_properties_core(datadir):
    datadir.join("genuine").chdir()
    archive = ZipFile("empty.xlsx")
    content = archive.read(ARC_CORE)
    prop = read_properties_core(content)
    assert prop.creator == '*.*'
    assert prop.excel_base_date == CALENDAR_WINDOWS_1900
    assert prop.last_modified_by == 'Charlie Clark'
    assert prop.created == datetime(2010, 4, 9, 20, 43, 12)
    assert prop.modified ==  datetime(2014, 1, 2, 14, 53, 6)


def test_read_properties_libreeoffice(datadir):
    datadir.join("genuine").chdir()
    archive = ZipFile("empty_libre.xlsx")
    content = archive.read(ARC_CORE)
    prop = read_properties_core(content)
    assert prop.excel_base_date == CALENDAR_WINDOWS_1900
    assert prop.creator == ''
    assert prop.last_modified_by == ''


@pytest.mark.parametrize("filename", ['empty.xlsx', 'empty_libre.xlsx'])
def test_read_sheets_titles(datadir, filename):
    from openpyxl.reader.workbook import read_sheets

    datadir.join("genuine").chdir()
    archive = ZipFile(filename)
    sheet_titles = [s[1] for s in read_sheets(archive)]
    assert sheet_titles == ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas', 'Sheet4 - Dates']


def test_write_properties_core(datadir):
    datadir.join("writer").chdir()
    prop = DocumentProperties()
    prop.creator = 'TEST_USER'
    prop.last_modified_by = 'SOMEBODY'
    prop.created = datetime(2010, 4, 1, 20, 30, 00)
    prop.modified = datetime(2010, 4, 5, 14, 5, 30)
    content = write_properties_core(prop)
    with open('core.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None

def test_write_properties_app(datadir):
    datadir.join("writer").chdir()
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_properties_app(wb)
    with open('app.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None
