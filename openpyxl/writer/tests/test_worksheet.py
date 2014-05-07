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
    doc = XMLGenerator(out, encoding="utf-8")
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
    assert out.getvalue() == ""


def test_write_col_widths(out, doc, write_cols, ColumnDimension):
    worksheet = DummyWorksheet()
    worksheet.column_dimensions['A'] = ColumnDimension(width=4)
    write_cols(doc, worksheet)
    doc.endDocument()
    xml = out.getvalue()
    expected = """<cols><col width="4" min="1" max="1"></col></cols>"""
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
  <col max="1" style="1" min="1"/>
  <col max="27" style="27" min="27"/>
  <col max="28" style="28" min="28"/>
  <col max="29" style="29" min="29"/>
  <col max="30" style="30" min="30"/>
  <col max="31" style="31" min="31"/>
  <col max="32" style="32" min="32"/>
  <col max="33" style="33" min="33"/>
  <col max="34" style="34" min="34"/>
  <col max="35" style="35" min="35"/>
  <col max="36" style="36" min="36"/>
  <col max="37" style="37" min="37"/>
  <col max="38" style="38" min="38"/>
  <col max="39" style="39" min="39"/>
  <col max="40" style="40" min="40"/>
  <col max="41" style="41" min="41"/>
  <col max="42" style="42" min="42"/>
  <col max="43" style="43" min="43"/>
  <col max="44" style="44" min="44"/>
  <col max="45" style="45" min="45"/>
  <col max="46" style="46" min="46"/>
  <col max="47" style="47" min="47"/>
  <col max="48" style="48" min="48"/>
  <col max="49" style="49" min="49"/>
  <col max="2" style="2" min="2"/>
  <col max="3" style="3" min="3"/>
  <col max="4" style="4" min="4"/>
  <col max="5" style="5" min="5"/>
  <col max="6" style="6" min="6"/>
  <col max="7" style="7" min="7"/>
  <col max="8" style="8" min="8"/>
  <col max="9" style="9" min="9"/>
  <col max="10" style="10" min="10"/>
  <col max="11" style="11" min="11"/>
  <col max="12" style="12" min="12"/>
  <col max="13" style="13" min="13"/>
  <col max="14" style="14" min="14"/>
  <col max="15" style="15" min="15"/>
  <col max="16" style="16" min="16"/>
  <col max="17" style="17" min="17"/>
  <col max="18" style="18" min="18"/>
  <col max="19" style="19" min="19"/>
  <col max="20" style="20" min="20"/>
  <col max="21" style="21" min="21"/>
  <col max="22" style="22" min="22"/>
  <col max="23" style="23" min="23"/>
  <col max="24" style="24" min="24"/>
  <col max="25" style="25" min="25"/>
  <col max="26" style="26" min="26"/>
</cols>
"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff
