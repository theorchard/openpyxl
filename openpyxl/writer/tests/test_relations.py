from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


# package
from openpyxl import Workbook
from openpyxl.xml.functions import tostring

# test imports
from openpyxl.tests.helper import compare_xml


def test_write_hyperlink_rels(datadir):
    from .. relations import write_rels
    wb = Workbook()
    ws = wb.active

    assert 0 == len(ws.relationships)
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com/"
    assert 1 == len(ws.relationships)
    ws.cell('A2').value = "test"
    ws.cell('A2').hyperlink = "http://test2.com/"
    assert 2 == len(ws.relationships)

    el = write_rels(ws, 1, 1, 1)
    xml = tostring(el)

    datadir.chdir()
    with open('sheet1_hyperlink.xml.rels') as expected:
        diff = compare_xml(xml, expected.read())
        assert diff is None, diff
