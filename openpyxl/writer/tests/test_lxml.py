from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# stdlib
import decimal
from io import BytesIO

# package
from openpyxl import Workbook
from lxml.etree import xmlfile

# test imports
import pytest
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def out():
    return BytesIO()


@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                         ])
def test_write_cell(out, value, expected):
    from .. lxml_worksheet import write_cell

    wb = Workbook()
    ws = wb.active
    ws['A1'] = value
    with xmlfile(out) as xf:
        write_cell(xf, ws, ws['A1'], [])
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_sheetdata(out):
    from .. lxml_worksheet import write_worksheet_data

    wb = Workbook()
    ws = wb.active
    ws['A1'] = 10
    write_worksheet_data(out, ws, [], None)
    xml = out.getvalue()
    expected = """<sheetData><row r="1" spans="1:1"><c t="n" r="A1"><v>10</v></c></row></sheetData>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff
