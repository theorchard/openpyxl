from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring, fromstring

from openpyxl.tests.helper import compare_xml
from ..colors import Color
from .. import colors


@pytest.fixture
def Side():
    from ..borders import Side
    return Side


@pytest.fixture
def Border():
    from ..borders import Border
    return Border


class TestBorder:

    def test_create(self, Border):
        src = """
        <border xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <left style="thin">
          <color rgb="FF006600"/>
        </left>
        <right style="thin">
          <color rgb="FF006600"/>
        </right>
        <top style="thin">
          <color rgb="FF006600"/>
        </top>
        <bottom style="thin">
          <color rgb="FF006600"/>
        </bottom>
        <diagonal/>
        </border>
        """
        xml = fromstring(src)
        bd = Border.from_tree(xml)
        assert bd.left.style == "thin"
        assert bd.right.color.value == "FF006600"
        assert bd.diagonal.style == None


    def test_serialise(self, Border, Side):
        medium_blue = Side(border_style='medium', color=Color(colors.BLUE))
        bd = Border(left=medium_blue,
                             right=medium_blue,
                             top=medium_blue,
                             bottom=medium_blue)
        xml = tostring(bd.to_tree())
        expected = """
        <border>
          <left style="medium">
            <color rgb="000000FF"></color>
          </left>
          <right style="medium">
            <color rgb="000000FF"></color>
          </right>
          <top style="medium">
            <color rgb="000000FF"></color>
           </top>
          <bottom style="medium">
             <color rgb="000000FF"></color>
           </bottom>
           <diagonal />
        </border>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
