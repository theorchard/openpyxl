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
    for i in range(1, 50):
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
   <col max="15" min="15" style="15"></col>
   <col max="16" min="16" style="16"></col>
   <col max="17" min="17" style="17"></col>
   <col max="18" min="18" style="18"></col>
   <col max="19" min="19" style="19"></col>
   <col max="20" min="20" style="20"></col>
   <col max="21" min="21" style="21"></col>
   <col max="22" min="22" style="22"></col>
   <col max="23" min="23" style="23"></col>
   <col max="24" min="24" style="24"></col>
   <col max="25" min="25" style="25"></col>
   <col max="26" min="26" style="26"></col>
   <col max="27" min="27" style="27"></col>
   <col max="28" min="28" style="28"></col>
   <col max="29" min="29" style="29"></col>
   <col max="30" min="30" style="30"></col>
   <col max="31" min="31" style="31"></col>
   <col max="32" min="32" style="32"></col>
   <col max="33" min="33" style="33"></col>
   <col max="34" min="34" style="34"></col>
   <col max="35" min="35" style="35"></col>
   <col max="36" min="36" style="36"></col>
   <col max="37" min="37" style="37"></col>
   <col max="38" min="38" style="38"></col>
   <col max="39" min="39" style="39"></col>
   <col max="40" min="40" style="40"></col>
   <col max="41" min="41" style="41"></col>
   <col max="42" min="42" style="42"></col>
   <col max="43" min="43" style="43"></col>
   <col max="44" min="44" style="44"></col>
   <col max="45" min="45" style="45"></col>
   <col max="46" min="46" style="46"></col>
   <col max="47" min="47" style="47"></col>
   <col max="48" min="48" style="48"></col>
   <col max="49" min="49" style="49"></col>
 </cols>
"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff
