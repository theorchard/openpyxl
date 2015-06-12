from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.styles.colors import BLACK, WHITE, Color
from openpyxl.xml.functions import tostring, fromstring

from openpyxl.tests.helper import compare_xml

@pytest.fixture
def GradientFill():
    from openpyxl.styles.fills import GradientFill
    return GradientFill


class TestGradientFill:

    def test_empty_ctor(self, GradientFill):
        gf = GradientFill()
        assert gf.type == 'linear'
        assert gf.degree == 0
        assert gf.left == 0
        assert gf.right == 0
        assert gf.top == 0
        assert gf.bottom == 0
        assert gf.stop == ()


    def test_ctor(self, GradientFill):
        gf = GradientFill(degree=90, left=1, right=2, top=3, bottom=4)
        assert gf.degree == 90
        assert gf.left == 1
        assert gf.right == 2
        assert gf.top == 3
        assert gf.bottom == 4


    def test_sequence(self, GradientFill):
        colors = [Color(BLACK), Color(WHITE)]
        gf = GradientFill(stop=colors)
        assert gf.stop == tuple(colors)


    def test_invalid_sequence(self, GradientFill):
        colors = [BLACK, WHITE]
        with pytest.raises(TypeError):
            GradientFill(stop=colors)


    def test_dict_interface(self, GradientFill):
        gf = GradientFill(degree=90, left=1, right=2, top=3, bottom=4)
        assert dict(gf) == {'bottom': "4", 'degree': "90", 'left':"1",
                            'right': "2", 'top': "3", 'type': 'linear'}


    def test_serialise(self, GradientFill):
        gf = GradientFill(degree=90, left=1, right=2, top=3, bottom=4, stop=[Color(BLACK), Color(WHITE)])
        xml = tostring(gf.to_tree())
        expected = """
        <fill>
        <gradientFill bottom="4" degree="90" left="1" right="2" top="3" type="linear">
           <stop position="0">
              <color rgb="00000000"></color>
            </stop>
            <stop position="1">
              <color rgb="00FFFFFF"></color>
            </stop>
        </gradientFill>
        </fill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_create(self, GradientFill):
        src = """
        <fill>
        <gradientFill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" degree="90">
        <stop position="0">
          <color theme="0"/>
        </stop>
        <stop position="1">
          <color theme="4"/>
        </stop>
        </gradientFill>
        </fill>
        """
        xml = fromstring(src)
        fill = GradientFill.from_tree(xml)
        assert fill.stop == (Color(theme=0), Color(theme=4))


@pytest.fixture
def PatternFill():
    from ..fills import PatternFill
    return PatternFill


class TestPatternFill:

    def test_ctor(self, PatternFill):
        pf = PatternFill()
        assert pf.patternType is None
        assert pf.fgColor == Color()
        assert pf.bgColor == Color()


    def test_dict_interface(self, PatternFill):
        pf = PatternFill(fill_type='solid')
        assert dict(pf) == {'patternType':'solid'}


    def test_serialise(self, PatternFill):
        pf = PatternFill('solid', 'FF0000', 'FFFF00')
        xml = tostring(pf.to_tree())
        expected = """
        <fill>
        <patternFill patternType="solid">
            <fgColor rgb="00FF0000"/>
            <bgColor rgb="00FFFF00"/>
        </patternFill>
        </fill>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.parametrize("src, args",
                             [
                                 ("""
                                 <fill>
                                 <patternFill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" patternType="solid">
                                   <fgColor theme="0" tint="-0.14999847407452621"/>
                                   <bgColor indexed="64"/>
                                 </patternFill>
                                 </fill>
                                 """,
                                dict(patternType='solid',
                                     start_color=Color(theme=0, tint=-0.14999847407452621),
                                     end_color=Color(indexed=64)
                                     )
                                ),
                                 ("""
                                 <fill>
                                 <patternFill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" patternType="solid">
                                   <fgColor theme="0"/>
                                   <bgColor indexed="64"/>
                                 </patternFill>
                                 </fill>
                                 """,
                                dict(patternType='solid',
                                     start_color=Color(theme=0),
                                     end_color=Color(indexed=64)
                                     )
                                ),
                                 ("""
                                 <fill>
                                 <patternFill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" patternType="solid">
                                   <fgColor indexed="62"/>
                                   <bgColor indexed="64"/>
                                 </patternFill>
                                 </fill>
                                 """,
                                dict(patternType='solid',
                                     start_color=Color(indexed=62),
                                     end_color=Color(indexed=64)
                                     )
                                ),
                             ]
                             )
    def test_create(self, PatternFill, src, args):
        xml = fromstring(src)
        assert PatternFill.from_tree(xml) == PatternFill(**args)


def test_create_empty_fill():
    from ..fills import Fill

    src = fromstring("<fill/>")
    assert Fill.from_tree(src) is None
