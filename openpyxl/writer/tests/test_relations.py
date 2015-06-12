from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


# package
from openpyxl import Workbook
from openpyxl.xml.functions import tostring

# test imports
from openpyxl.tests.helper import compare_xml


def test_write_hyperlink_rels(datadir):
    from .. relations import write_rels
    wb = Workbook()
    ws = wb.active

    assert 0 == len(ws.relationships)
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com/"
    assert 1 == len(ws.relationships)
    ws.cell('A2').value = "test"
    ws.cell('A2').hyperlink = "http://test2.com/"
    assert 2 == len(ws.relationships)

    el = write_rels(ws, 1, 1, 1)
    xml = tostring(el)

    datadir.chdir()
    with open('sheet1_hyperlink.xml.rels') as expected:
        diff = compare_xml(xml, expected.read())
        assert diff is None, diff

import pytest

class Worksheet:

    _comment_count = 0
    vba_controls = None
    relationships = ()
    _charts = ()
    _images = ()


@pytest.fixture
def writer():
    from ..relations import write_rels
    return write_rels


class TestRels:

    def test_comments(self, writer):
        ws = Worksheet()
        ws._comment_count = 1
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
         <Relationship Id="comments" Target="../comments1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" />
          <Relationship Id="commentsvml" Target="../drawings/commentsDrawing1.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>
        </Relationships>
        """
        xml = tostring(writer(ws, None, 1, None))
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_vba(self, writer):
        ws = Worksheet()
        ws.vba_controls = "vba"
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="vba" Target="../drawings/vmlDrawing1.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>
        </Relationships>
            """
        xml = tostring(writer(ws, None, None, 1))
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_drawing(self, writer):
        ws = Worksheet()
        ws._charts = [None]
        expected = """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Target="../drawings/drawing1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
            </Relationships>
                """
        xml = tostring(writer(ws, 1, None, None))
        diff = compare_xml(xml, expected)
        assert diff is None, diff
