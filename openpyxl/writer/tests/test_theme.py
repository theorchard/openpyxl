# Copyright (c) 2010-2014 openpyxl

# Python stdlib imports

# package imports
from openpyxl.tests.helper import compare_xml
from openpyxl.writer.theme import write_theme


def test_write_theme(datadir):
    datadir.chdir()
    content = write_theme()
    with open( 'theme1.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
