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
    assert prop.last_modified_by == 'Charlie Clark'
    assert prop.created == datetime(2010, 4, 9, 20, 43, 12)
    assert prop.modified ==  datetime(2014, 1, 2, 14, 53, 6)


def test_read_properties_libreeoffice(datadir):
    datadir.join("genuine").chdir()
    archive = ZipFile("empty_libre.xlsx")
    content = archive.read(ARC_CORE)
    prop = read_properties_core(content)
    assert prop.creator == ''
    assert prop.last_modified_by == ''


@pytest.mark.parametrize("filename", ['empty.xlsx', 'empty_libre.xlsx'])
def test_read_sheets_titles(datadir, filename):
    from openpyxl.reader.workbook import read_sheets

    datadir.join("genuine").chdir()
    archive = ZipFile(filename)
    sheet_titles = [s['name'] for s in read_sheets(archive)]
    assert sheet_titles == ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas', 'Sheet4 - Dates']


def test_write_properties_app(datadir):
    datadir.join("writer").chdir()
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_properties_app(wb)
    with open('app.xml') as expected:
        diff = compare_xml(content, expected.read())
    assert diff is None, diff


def test_read_workbook_with_no_core_properties(datadir):
    from openpyxl.workbook import DocumentProperties
    from openpyxl.reader.excel import _load_workbook

    datadir.join('genuine').chdir()
    archive = ZipFile('empty_with_no_properties.xlsx')
    wb = Workbook()
    default_props = DocumentProperties()
    _load_workbook(wb, archive, None, False, False)
    prop = wb.properties
    assert prop.creator == default_props.creator
    assert prop.last_modified_by == default_props.last_modified_by
    assert prop.title == default_props.title
    assert prop.subject == default_props.subject
    assert prop.description == default_props.description
    assert prop.category == default_props.category
    assert prop.keywords == default_props.keywords
    assert prop.created.timetuple()[:9] == default_props.created.timetuple()[:9] # might break if tests run on the stoke of midnight
    assert prop.modified.timetuple()[:9] == prop.created.timetuple()[:9]
