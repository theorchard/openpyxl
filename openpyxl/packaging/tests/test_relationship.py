from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import pytest
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import tostring


@pytest.fixture
def Relationship():
    from ..relationship import Relationship
    return Relationship


def test_ctor(Relationship):
    rel = Relationship("drawing", "drawings.xml", "external", "4")

    assert dict(rel) == {'Id': '4', 'Target': 'drawings.xml', 'TargetMode':
                         'external', 'Type':
                         'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'}

    expected = """<Relationship Id="4" Target="drawings.xml" TargetMode="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" />
    """
    xml = tostring(rel.to_tree())

    diff = compare_xml(xml, expected)
    assert diff is None, diff
