# Copyright (c) 2010-2014 openpyxl

# Python stdlib imports

# package imports
from openpyxl.tests.helper import compare_xml
from openpyxl.writer.workbook import write_content_types, write_root_rels
from openpyxl.workbook import Workbook


def test_write_content_types(datadir):
    datadir.join("writer").chdir()
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_content_types(wb)
    with open('[Content_Types].xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_root_rels(datadir):
    datadir.join("writer").chdir()
    wb = Workbook()
    content = write_root_rels(wb)
    with open('.rels') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
