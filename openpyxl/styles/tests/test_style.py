from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest


@pytest.fixture
def StyleId():
    from .. style import StyleId
    return StyleId


def test_ctor(StyleId):
    style = StyleId()
    assert dict(style) == {'borderId': '0', 'fillId': '0', 'fontId': '0',
                           'numFmtId': '0', 'xfId': '0'}


def test_protection(StyleId):
    style = StyleId(protection=1)
    assert dict(style) == {'borderId': '0', 'fillId': '0', 'fontId': '0',
                           'numFmtId': '0', 'xfId': '0', 'applyProtection':'1'}


def test_alignment(StyleId):
    style = StyleId(alignment=1)
    assert dict(style) == {'borderId': '0', 'fillId': '0', 'fontId': '0',
                           'numFmtId': '0', 'xfId': '0', 'applyAlignment':'1'}
