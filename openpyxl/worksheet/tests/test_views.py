from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.xml.functions import fromstring, tostring

import pytest

from openpyxl.tests.helper import compare_xml


@pytest.fixture
def SheetView():
    from ..views import SheetView
    return SheetView


@pytest.mark.parametrize("value, result",
                         [
                             (True, {'workbookViewId': '0'}),
                             (False, {'workbookViewId': '0', 'showGridLines':'0'})
                         ]
                         )
def test_show_gridlines(SheetView, value, result):
    view = SheetView(showGridLines=value)
    assert dict(view) == result


def test_parse(SheetView):
    src = """
     <sheetView xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" tabSelected="1" zoomScale="200" zoomScaleNormal="200" zoomScalePageLayoutView="200" workbookViewId="0">
      <pane xSplit="5" ySplit="19" topLeftCell="F20" activePane="bottomRight" state="frozenSplit"/>
      <selection pane="topRight" activeCell="F1" sqref="F1"/>
      <selection pane="bottomLeft" activeCell="A20" sqref="A20"/>
      <selection pane="bottomRight" activeCell="E22" sqref="E22"/>
    </sheetView>
    """
    xml = fromstring(src)
    view = SheetView.from_tree(xml)
    assert dict(view) == {'tabSelected': '1', 'zoomScale': '200', 'workbookViewId':"0",
                          'zoomScaleNormal': '200', 'zoomScalePageLayoutView': '200'}
    assert len(view.selection) == 3


def test_serialise(SheetView):
    view = SheetView()

    xml = tostring(view.to_tree())
    expected = """
    <sheetView workbookViewId="0">
       <selection activeCell="A1" sqref="A1"></selection>
    </sheetView>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
