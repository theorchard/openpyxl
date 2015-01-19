from __future__ import absolute_import
# copyright 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Rule():
    from ..rule import Rule
    return Rule


def test_create(Rule, datadir):
    datadir.chdir()
    with open("worksheet.xml") as src:
        xml = fromstring(src.read())
    for el in xml.findall("{%s}conditionalFormatting/{%s}cfRule" % (SHEET_MAIN_NS, SHEET_MAIN_NS)):
        rule = Rule.create(el)
