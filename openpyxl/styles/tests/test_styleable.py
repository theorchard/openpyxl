from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.utils.indexed_list import IndexedList
from ..import Style, Font, Border, PatternFill, Alignment, Protection


def test_descriptor():
    from ..styleable import StyleDescriptor
    from ..fonts import Font

    class Styled(object):

        font = StyleDescriptor('_fonts', '_font_id')

        def __init__(self):
            self.parent = DummyWorksheet()

    styled = Styled()
    styled.font = Font()
    assert styled._font_id == 0
    assert styled.font == Font()


class DummyWorkbook:

    _fonts = IndexedList()
    _fills = IndexedList()
    _borders = IndexedList()
    _protections = IndexedList()
    _alignments = IndexedList()
    _number_formats = IndexedList()


class DummyWorksheet:

    parent = DummyWorkbook()


@pytest.fixture
def StyleableObject():
    from .. styleable import StyleableObject
    return StyleableObject


def test_has_style(StyleableObject):
    so = StyleableObject(sheet=DummyWorksheet())
    assert so._number_format_id == 0
    assert not so.has_style
    so._number_format_id = 1
    assert so.has_style


def test_style(StyleableObject):
    so = StyleableObject(sheet=DummyWorksheet())
    so.font = Font(bold=True)
    so.fill = PatternFill()
    so.border = Border()
    so.protection = Protection()
    so.alignment = Alignment()
    so.number_format = "not general"
    assert so.style == Style(font=Font(bold=True), number_format="not general")


def test_assign_style(StyleableObject):
    so = StyleableObject(sheet=DummyWorksheet())
    style = Style(font=Font(underline="single"))
    so.style = style
    assert style.font == Font(underline="single")
