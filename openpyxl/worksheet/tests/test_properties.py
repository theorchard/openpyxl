from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring

from openpyxl.styles.colors import Color
from openpyxl.tests.schema import sheet_schema
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.functions import safe_iterator, tostring

def test_ctor():
    from .. properties import WorksheetProperties, Outline
    color_test = 'F0F0F0'
    outline_pr = Outline(summaryBelow=True, summaryRight=True)
    wsprops = WorksheetProperties(tabColor=color_test, outlinePr=outline_pr)
    assert dict(wsprops) == {}
    assert dict(wsprops.outlinePr) == {'summaryBelow': '1', 'summaryRight': '1'}
    assert dict(wsprops.tabColor) == {'rgb': '00F0F0F0'}


@pytest.fixture
def SimpleTestProps():
    from .. properties import WorksheetProperties, PageSetupPr
    wsp = WorksheetProperties()
    wsp.filterMode = False
    wsp.tabColor = 'FF123456'
    wsp.pageSetUpPr = PageSetupPr(fitToPage=False)
    return wsp


def test_write_properties(SimpleTestProps):

    xml = tostring(SimpleTestProps.to_tree())
    expected = """
    <sheetPr filterMode="0">
      <pageSetUpPr fitToPage="0" />
      <tabColor rgb="FF123456"/>
    </sheetPr>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_parse_properties(datadir, SimpleTestProps):
    from .. properties import WorksheetProperties
    datadir.chdir()

    with open("sheetPr2.xml") as src:
        content = src.read()

    xml = fromstring(content)
    parseditem = WorksheetProperties.from_tree(xml)
    assert dict(parseditem) == dict(SimpleTestProps)
    assert parseditem.tabColor == SimpleTestProps.tabColor
    assert dict(parseditem.pageSetUpPr) == dict(SimpleTestProps.pageSetUpPr)
