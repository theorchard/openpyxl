from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from io import BytesIO

import pytest


from openpyxl.workbook import Workbook
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

# There are already unit-tests in test_cell.py that test out the
# coordinate_from_string method.  This should be the only way the
# collapse_cell_addresses method can throw, so we don't bother using invalid
# cell coordinates in the test-data here.
COLLAPSE_TEST_DATA = [
    (["A1"], "A1"),
    (["A1", "B1"], "A1 B1"),
    (["A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4"], "A1:A4 B1:B4"),
    (["A2", "A4", "A3", "A1", "A5"], "A1:A5"),
]
@pytest.mark.parametrize("cells, expected",
                         COLLAPSE_TEST_DATA)
def test_collapse_cell_addresses(cells, expected):
    from .. datavalidation import collapse_cell_addresses
    assert collapse_cell_addresses(cells) == expected


def test_expand_cell_ranges():
    from .. datavalidation import expand_cell_ranges
    rs = "A1:A3 B1:B3"
    assert expand_cell_ranges(rs) == ["A1", "A2", "A3", "B1", "B2", "B3"]


@pytest.fixture
def DataValidation():
    from .. datavalidation import DataValidation
    return DataValidation


def test_list_validation(DataValidation):
    dv = DataValidation(type="list", formula1='"Dog,Cat,Fish"')
    assert dv.formula1, '"Dog,Cat == Fish"'
    dv_dict = dict(dv)
    assert dv_dict['type'] == 'list'
    assert dv_dict['allowBlank'] == '0'
    assert dv_dict['showErrorMessage'] == '1'
    assert dv_dict['showInputMessage'] == '1'


def test_error_message(DataValidation):
    dv = DataValidation("list", formula1='"Dog,Cat,Fish"')
    dv.set_error_message('You done bad')
    dv_dict = dict(dv)
    assert dv_dict['errorTitle'] == 'Validation Error'
    assert dv_dict['error'] == 'You done bad'


def test_prompt_message(DataValidation):
    dv = DataValidation(type="list", formula1='"Dog,Cat,Fish"')
    dv.set_prompt_message('Please enter a value')
    dv_dict = dict(dv)
    assert dv_dict['promptTitle'] == 'Validation Prompt'
    assert dv_dict['prompt'] == 'Please enter a value'


def test_writer_validation(DataValidation):
    from .. datavalidation import writer
    wb = Workbook()
    ws = wb.active
    dv = DataValidation(type="list", formula1='"Dog,Cat,Fish"')
    dv.add_cell(ws['A1'])

    xml = tostring(writer(dv))
    expected = """
    <dataValidation xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" type="list">
      <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
    </dataValidation>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_sqref(DataValidation):
    from .. datavalidation import DataValidation
    dv = DataValidation()
    dv.sqref = "A1"
    assert dv.cells == ["A1"]


def test_ctor(DataValidation):
    from .. datavalidation import DataValidation
    dv = DataValidation()
    assert dict(dv) == {'allowBlank': '0', 'showErrorMessage': '1',
                        'showInputMessage': '1', 'sqref': ''}


def test_with_formula():
    from .. datavalidation import parser
    xml = """
    <dataValidation xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" type="list">
      <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
    </dataValidation>
    """
    dv = parser(fromstring(xml))
    assert dv.cells == ['A1']
    assert dv.type == "list"
    assert dv.formula1 == '"Dog,Cat,Fish"'


def test_parser():
    from .. datavalidation import parser
    xml = """
    <dataValidation xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" type="list" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" error="Value must be between 1 and 3!" errorTitle="An Error Message" promptTitle="Multiplier" prompt="for monthly or quartely reports" sqref="H6">
    </dataValidation>
"""
    dv = parser(fromstring(xml))
    assert dict(dv) == {"error":"Value must be between 1 and 3!",
                        "errorStyle":"warning",
                        "errorTitle":"An Error Message",
                        "prompt":"for monthly or quartely reports",
                        "promptTitle":"Multiplier",
                        "type":"list",
                        "allowBlank":"1",
                        "sqref":"H6",
                        "showErrorMessage":"1",
                        "showInputMessage":"1"}

    from ..datavalidation import writer
    tag = writer(dv)
    output = tostring(tag)
    diff = compare_xml(output, xml)
    assert diff is None,diff
