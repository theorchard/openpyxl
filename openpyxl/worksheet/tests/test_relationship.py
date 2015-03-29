from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import pytest
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Relationship():
    from ..relationship import Relationship
    return Relationship


def test_ctor(Relationship):
    rel = Relationship("drawing", "drawings.xml", "external", "4")
    expected = """
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship id="4" target="drawings.xml" target_mode="external" type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" />
    </Relationships>
    """
    xml = repr(rel)
    diff = compare_xml(xml, expected)
    assert diff is None, diff
