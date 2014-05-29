from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# package
from openpyxl import Workbook

# test
from openpyxl.tests.helper import compare_xml


def test_write_auto_filter(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.active
    ws.cell('F42').value = 'hello'
    ws.auto_filter.ref = 'A1:F1'

    from .. workbook import write_workbook

    content = write_workbook(wb)
    with open('workbook_auto_filter.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
