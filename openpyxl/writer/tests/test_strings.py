from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.tests.helper import compare_xml


def test_write_string_table(datadir):
    from openpyxl.writer.strings import write_string_table

    datadir.chdir()
    table = ['This is cell A1 in Sheet 1', 'This is cell G5']
    content = write_string_table(table)
    with open('sharedStrings.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
