from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def StyleId():
    from .. style import StyleId
    return StyleId


def test_ctor(StyleId):
    style = StyleId()
    assert dict(style) == {'borderId': 0, 'fillId': 0, 'fontId': 0,
                           'numFmtId': 0, 'xfId': 0, 'alignmentId':0, 'protectionId':0}


def test_protection(StyleId):
    style = StyleId(protectionId=1)
    assert style.applyProtection is True


def test_alignment(StyleId):
    style = StyleId(alignmentId=1)
    assert style.applyAlignment is True


def test_serialise(StyleId):
    style = StyleId()
    xml = tostring(style.to_tree())
    expected = """
     <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" />
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_style_copy():
    from .. import Style
    st1 = Style()
    st2 = st1.copy()
    assert st1 == st2
    assert st1.font is not st2.font
