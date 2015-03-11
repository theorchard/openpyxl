# Copyright (c) 2010-2015 openpyxl

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


def test_ctor_header(HeaderFooter):
    header = HeaderFooter()
    assert header.hasHeader() is False
    assert header.hasFooter() is False


def test_set_header(HeaderFooter):
    header = HeaderFooter()
    header.setHeader('&L&"Lucida Grande,Standard"&K000000Left top')
    assert header.hasHeader() is True
    hf = header.left_header
    assert hf.text == "Left top"


def test_set_item(HeaderFooterItem):
    hf = HeaderFooterItem('L')
    hf.set('&"Lucida Grande,Standard"&K000000Left top')

    assert hf.text == "Left top"
    assert hf.font_name == "Lucida Grande,Standard"
    assert hf.font_color == "000000"


def test_split_into_parts():
    from .. header_footer import ITEM_REGEX
    m = ITEM_REGEX.match("&Ltest header")
    assert m.group('left') == "test header"
    m = ITEM_REGEX.match("""&L&"Lucida Grande,Standard"&K000000Left top&C&"Lucida Grande,Standard"&K000000Middle top&R&"Lucida Grande,Standard"&K000000Right top""")
    assert m.group('left') == '&"Lucida Grande,Standard"&K000000Left top'
    assert m.group('center') == '&"Lucida Grande,Standard"&K000000Middle top'
    assert m.group('right') == '&"Lucida Grande,Standard"&K000000Right top'


def test_multiline_string():
    from .. header_footer import ITEM_REGEX
    s = """&L141023 V1&CRoute - Malls\nSchedules R1201 v R1301&RClient-internal use only"""
    match = ITEM_REGEX.match(s)
    assert match.groupdict() == {
        'center': 'Route - Malls\nSchedules R1201 v R1301',
        'left': '141023 V1',
        'right': 'Client-internal use only'
    }


def test_font_size():
    from .. header_footer import SIZE_REGEX
    s = "&9"
    match = SIZE_REGEX.search(s)
    assert match.group('size') == "9"
