from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# package imports
from openpyxl.tests.helper import compare_xml
from openpyxl.writer.workbook import write_content_types, write_root_rels
from openpyxl.workbook import Workbook


def test_write_content_types(datadir):
    datadir.chdir()
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_content_types(wb)
    with open('[Content_Types].xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_root_rels(datadir):
    datadir.chdir()
    wb = Workbook()
    content = write_root_rels(wb)
    with open('.rels') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
