# Copyright (c) 2010-2014 openpyxl

import pytest
from lxml.etree import fromstring
from openpyxl.styles.colors import Color
from openpyxl.tests.schema import sheet_schema
from openpyxl.tests.helper import compare_xml
from _pytest.main import Node

from openpyxl.xml.functions import safe_iterator

def test_ctor():
    from .. properties import WorksheetProperties
    color1 = 'F0F0F0'
    wsprops = WorksheetProperties(tabColor=color1)
    assert dict(wsprops) == {}

@pytest.fixture
def TabColorProps():
    from .. properties import WorksheetProperties
    wsp = WorksheetProperties()
    wsp.filterMode = False
    wsp.tabColor = '1072BA'
    return wsp

def test_TabColorProps(TabColorProps):
    assert dict(TabColorProps) == {'filterMode':'false'}
    assert TabColorProps.filterMode == False
    assert TabColorProps.tabColor.value == '001072BA'
    assert TabColorProps.tabColor.type == 'rgb'
    assert TabColorProps.tabColor.tint == 0.0

def test_write_properties(TabColorProps):
    from .. properties import write_sheetPr

    content = write_sheetPr(TabColorProps)
    expected = """ <s:sheetPr xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" filterMode="false"><s:tabColor rgb="001072BA"/></s:sheetPr>"""
    diff = compare_xml(content, expected)
    assert diff is None, diff

@pytest.fixture
def SimpleTestProps():
    from .. properties import WorksheetProperties, PageSetup
    wsp = WorksheetProperties()
    wsp.filterMode = False
    wsp.tabColor = 'FF123456'
    wsp.pageSetUpPr = PageSetup(fitToPage=False)
    return wsp

def test_parse_properties(datadir, SimpleTestProps):
    from .. properties import parse_sheetPr
    datadir.chdir()

    with open("sheetPr2.xml") as src:
        content = src.read()

    parseditem = parse_sheetPr(fromstring(content))
    assert dict(parseditem) == dict(SimpleTestProps)
    assert parseditem.tabColor == SimpleTestProps.tabColor
    assert dict(parseditem.pageSetUpPr) == dict(SimpleTestProps.pageSetUpPr)

    with open("fullsheet.xml") as src:
        content = src.read()

    root = fromstring(content)
    for node in safe_iterator(root):
        if node.tag == SimpleTestProps.tag:
            parseditem = parse_sheetPr(node)

    assert dict(parseditem) == dict(SimpleTestProps)
    assert parseditem.tabColor == SimpleTestProps.tabColor
    assert dict(parseditem.pageSetUpPr) == dict(SimpleTestProps.pageSetUpPr)
