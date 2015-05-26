from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest
from lxml.etree import fromstring
from openpyxl.styles.colors import Color
from openpyxl.tests.schema import sheet_schema
from openpyxl.tests.helper import compare_xml
from _pytest.main import Node

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
    from .. properties import WorksheetProperties, PageSetupProperties
    wsp = WorksheetProperties()
    wsp.filterMode = False
    wsp.tabColor = 'FF123456'
    wsp.pageSetUpPr = PageSetupProperties(fitToPage=False)
    return wsp


def test_write_properties(SimpleTestProps):
    from .. properties import write_sheetPr

    content = write_sheetPr(SimpleTestProps)
    expected = """ <s:sheetPr xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" filterMode="0"><s:pageSetUpPr fitToPage="0" /><s:tabColor rgb="FF123456"/></s:sheetPr>"""
    diff = compare_xml(tostring(content), expected)
    assert diff is None, diff


def test_parse_properties(datadir, SimpleTestProps):
    from .. properties import parse_sheetPr
    datadir.chdir()

    with open("sheetPr2.xml") as src:
        content = src.read()

    parseditem = parse_sheetPr(fromstring(content))
    assert dict(parseditem) == dict(SimpleTestProps)
    assert parseditem.tabColor == SimpleTestProps.tabColor
    assert dict(parseditem.pageSetUpPr) == dict(SimpleTestProps.pageSetUpPr)
