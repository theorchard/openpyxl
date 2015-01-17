from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.styles import Font, Color, PatternFill


@pytest.fixture
def ConditionalFormat():
    from ..conditional import ConditionalFormat
    return ConditionalFormat


def test_parse(ConditionalFormat, datadir):
    datadir.chdir()
    with open("dxf_style.xml") as content:
        src = content.read()
    xml = fromstring(src)
    formats = []
    for node in xml.findall("{%s}dxfs/{%s}dxf" % (SHEET_MAIN_NS, SHEET_MAIN_NS) ):
        formats.append(ConditionalFormat.create(node))
    assert len(formats) == 164
    cond = formats[1]
    assert cond.font == Font(underline="double", color=Color(auto=1), strikethrough=True, italic=True)
    assert cond.fill == PatternFill(end_color='FFFFC7CE')
