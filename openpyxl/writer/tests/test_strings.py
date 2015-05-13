from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.tests.helper import compare_xml


def test_write_string_table(datadir):
    from ..strings import write_string_table

    datadir.chdir()
    table = ['This is cell A1 in Sheet 1', 'This is cell G5']
    content = write_string_table(table)
    with open('sharedStrings.xml') as src:
        expected = src.read()
    diff = compare_xml(content, expected)
    assert diff is None, diff


def test_preseve_space():
    from ..strings import write_string_table
    table = ['String with trailing space   ']
    content = write_string_table(table)
    expected = """
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="1">
    <si>
      <t xml:space="preserve">String with trailing space</t>
    </si>
    </sst>
    """
    diff = compare_xml(content, expected)
    assert diff is None, diff
