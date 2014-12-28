# Copyright (c) 2010-2014 openpyxl

import pytest


@pytest.fixture
def HeaderFooterItem():
    from .. header_footer import HeaderFooterItem
    return HeaderFooterItem


@pytest.fixture
def HeaderFooter():
    from .. header_footer import HeaderFooter
    return HeaderFooter


def test_ctor_item(HeaderFooterItem):
    hf = HeaderFooterItem("L")
    assert hf.font_size == None
    assert hf.font_name == "Calibri,Regular"
    assert hf.font_color == "000000"
    assert hf.type == "L"
