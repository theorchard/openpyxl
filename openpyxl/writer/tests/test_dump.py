from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest


class DummyWorkbook:

    def __init__(self):
        self.shared_strings = []
        self.shared_styles = []

    def get_sheet_names(self):
        return []


@pytest.fixture
def DumpWorksheet():
    from .. dump_worksheet import DumpWorksheet
    return DumpWorksheet(DummyWorkbook(), "TestWorkSheet")


def test_ctor(DumpWorksheet):
    ws = DumpWorksheet
    assert isinstance(ws._parent, DummyWorkbook)
    assert ws.title == "TestWorkSheet"
    assert ws._max_col == 0
    assert ws._max_row == 0
    assert hasattr(ws, '_fileobj_header_name')
    assert hasattr(ws, '_fileobj_content_name')
    assert hasattr(ws, '_fileobj_name')
    assert ws._comments == []
