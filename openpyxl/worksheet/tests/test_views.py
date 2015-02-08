from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

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
