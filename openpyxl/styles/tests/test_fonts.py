from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring

from openpyxl.tests.helper import compare_xml
from ..colors import Color
from .. import colors


@pytest.fixture
def Font():
    from ..fonts import Font
    return Font


class TestFont:

    def test_ctor(self, Font):
        f = Font()
        assert f.name == 'Calibri'
        assert f.size == 11
        assert f.bold is False
        assert f.italic is False
        assert f.underline is None
        assert f.strikethrough is False
        assert f.color.value == '00000000'
        assert f.color.type == 'rgb'
        assert f.vertAlign is None
        assert f.charset is None


    def test_serialise(self, Font):
        ft = Font(name='Calibri', charset=204, vertAlign='superscript', underline='single')
        xml = tostring(ft.to_tree())
        expected = """
        <font>
          <name val="Calibri"></name>
          <charset val="204"></charset>
          <family val="2"></family>
          <color rgb="00000000"></color>
          <sz val="11"></sz>
          <u val="single"/>
          <vertAlign val="superscript" />
         </font>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_create(self, Font):
        src = """
        <font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <charset val="204"></charset>
          <family val="2"></family>
          <name val="Calibri"></name>
          <sz val="11"></sz>
          <u val="single"/>
          <vertAlign val="superscript"></vertAlign>
          <color rgb="FF3300FF"></color>
         </font>"""
        xml = fromstring(src)
        ft = Font.from_tree(xml)
        assert ft == Font(name='Calibri', charset=204, vertAlign='superscript', underline='single', color="FF3300FF")


