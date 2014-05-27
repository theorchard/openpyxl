from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from io import BytesIO
import pytest

from openpyxl.xml.functions import XMLGenerator

from openpyxl.tests.helper import compare_xml


class DummyWorksheet:

    _styles = {}
    column_dimensions = {}


@pytest.fixture
def out():
    return BytesIO()


@pytest.fixture
def doc(out):
    doc = XMLGenerator(out)
    return doc


@pytest.fixture
def write_cols():
    from .. worksheet import write_worksheet_cols
    return write_worksheet_cols


@pytest.fixture
def ColumnDimension():
    from openpyxl.worksheet.dimensions import ColumnDimension
    return ColumnDimension


def test_write_no_cols(out, doc, write_cols):
    write_cols(doc, DummyWorksheet())
    doc.endDocument()
    assert out.getvalue() == b""


def test_write_col_widths(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    worksheet.column_dimensions['A'] = ColumnDimension(width=4)
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols><col width="4" min="1" max="1" customWidth="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_cols_style(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    worksheet.column_dimensions['A'] = ColumnDimension()
    worksheet._styles['A'] = 1
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols><col max="1" min="1" style="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_lots_cols(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    from openpyxl.cell import get_column_letter
    for i in range(1, 15):
        label = get_column_letter(i)
        worksheet._styles[label] = i
        worksheet.column_dimensions[label] = ColumnDimension()
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols>
   <col max="1" min="1" style="1"></col>
   <col max="2" min="2" style="2"></col>
   <col max="3" min="3" style="3"></col>
   <col max="4" min="4" style="4"></col>
   <col max="5" min="5" style="5"></col>
   <col max="6" min="6" style="6"></col>
   <col max="7" min="7" style="7"></col>
   <col max="8" min="8" style="8"></col>
   <col max="9" min="9" style="9"></col>
   <col max="10" min="10" style="10"></col>
   <col max="11" min="11" style="11"></col>
   <col max="12" min="12" style="12"></col>
   <col max="13" min="13" style="13"></col>
   <col max="14" min="14" style="14"></col>
 </cols>
"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff
