# Copyright (c) 2010-2014 openpyxl

import pytest
from openpyxl.styles.colors import Color
from ..properties import Outline, PageSetup


@pytest.fixture
def WorksheetProperties():
    from .. properties import WorksheetProperties
    return WorksheetProperties

def test_empty_wsproperties(WorksheetProperties):
    wsp = WorksheetProperties()
    assert wsp.codeName == None
    assert wsp.enableFormatConditionsCalculation == None
    assert wsp.filterMode == None
    assert wsp.published == None
    assert wsp.syncHorizontal == None
    assert wsp.syncRef == None
    assert wsp.syncVertical == None
    assert wsp.transitionEvaluation == None
    assert wsp.transitionEntry == None
    assert wsp.tabColor == None
    assert isinstance(wsp.outlinePr, Outline)
    assert isinstance(wsp.pageSetUpPr,PageSetup)

def test_tab_color(WorksheetProperties):
    wsp = WorksheetProperties(tabColor = '1072BA')
    assert wsp.tabColor.value == '001072BA'
    wsp.tabColor = "001073AB"
    assert wsp.tabColor == Color(rgb='001073AB')
    with pytest.raises(ValueError):
        wsp.tabColor = "01072BA"
    with pytest.raises(ValueError):
        wsp.tabColor = ""
    with pytest.raises(ValueError):
        wsp.tabColor = "12345"
    wsp.tabColor = "123456"
    assert wsp.tabColor == Color(rgb='00123456')
    wsp.tabColor = "12345600"
    assert wsp.tabColor == Color(rgb='12345600')
    with pytest.raises(ValueError):
        wsp.tabColor = "#1234567"
    with pytest.raises(ValueError):
        wsp.tabColor = "ABCDEFGHI"
    with pytest.raises(TypeError):
        wsp.tabColor = 123456
        