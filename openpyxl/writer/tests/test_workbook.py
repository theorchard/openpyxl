from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# package
from openpyxl import Workbook

# test
import pytest
from openpyxl.tests.helper import compare_xml


from .. workbook import write_workbook


def test_write_auto_filter(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.active
    ws.cell('F42').value = 'hello'
    ws.auto_filter.ref = 'A1:F1'

    content = write_workbook(wb)
    with open('workbook_auto_filter.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


@pytest.mark.xfail
# This should actually fail because you cannot hide the only sheet of a workbook
def test_write_hidden_worksheet():
    wb = Workbook()
    ws = wb.active
    ws.sheet_state = ws.SHEETSTATE_HIDDEN
    xml = write_workbook(wb)
    expected = """
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
    <workbookPr codeName="ThisWorkbook" defaultThemeVersion="124226"/>
    <bookViews>
      <workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
    </bookViews>
    <sheets>
      <sheet name="Sheet" sheetId="1" state="hidden" r:id="rId1"/>
    </sheets>
      <definedNames/>
      <calcPr calcId="124519" calcMode="auto" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
