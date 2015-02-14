from __future__ import absolute_import
# copyright 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.styles import Color

from openpyxl.tests.helper import compare_xml


@pytest.fixture
def FormatObject():
    from ..rule import FormatObject
    return FormatObject


class TestFormatObject:

    def test_create(self, FormatObject):
        xml = fromstring("""<cfvo type="num" val="3"/>""")
        cfvo = FormatObject.from_tree(xml)
        assert cfvo.type == "num"
        assert cfvo.val == 3
        assert cfvo.gte is None


    def test_serialise(self, FormatObject):
        cfvo = FormatObject(type="percent", val=4)
        xml = tostring(cfvo.to_tree())
        expected = """<cfvo type="percent" val="4"/>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def ColorScale():
    from ..rule import ColorScale
    return ColorScale


class TestColorScale:

    def test_create(self, ColorScale):
        src = """
        <colorScale xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FFFF7128"/>
        <color rgb="FFFFEF9C"/>
        </colorScale>
        """
        xml = fromstring(src)
        cs = ColorScale.from_tree(xml)
        assert len(cs.cfvo) == 2
        assert len(cs.color) == 2


    def test_serialise(self, ColorScale, FormatObject):
        fo1 = FormatObject(type="min", val="0")
        fo2 = FormatObject(type="percent", val="50")
        fo3 = FormatObject(type="max", val="0")

        col1 = Color(rgb="FFFF0000")
        col2 = Color(rgb="FFFFFF00")
        col3 = Color(rgb="FF00B050")

        cs = ColorScale(cfvo=[fo1, fo2, fo3], color=[col1, col2, col3])
        xml = tostring(cs.to_tree())
        expected = """
        <colorScale>
        <cfvo type="min" val="0"/>
        <cfvo type="percent" val="50"/>
        <cfvo type="max" val="0"/>
        <color rgb="FFFF0000"/>
        <color rgb="FFFFFF00"/>
        <color rgb="FF00B050"/>
        </colorScale>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def DataBar():
    from ..rule import DataBar
    return DataBar


class TestDataBar:

    def test_create(self, DataBar):
        src = """
        <dataBar xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF638EC6"/>
        </dataBar>
        """
        xml = fromstring(src)
        db = DataBar.from_tree(xml)
        assert len(db.cfvo) == 2
        assert db.color.value == "FF638EC6"


    def test_serialise(self, DataBar, FormatObject):
        fo1 = FormatObject(type="min", val="0")
        fo2 = FormatObject(type="percent", val="50")
        db = DataBar(minLength=4, maxLength=10, cfvo=[fo1, fo2], color="FF2266", showValue=True)
        xml = tostring(db.to_tree())
        expected = """
        <dataBar maxLength="10" minLength="4" showValue="1">
          <cfvo type="min" val="0"></cfvo>
          <cfvo type="percent" val="50"></cfvo>
          <color rgb="00FF2266"></color>
         </dataBar>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def IconSet():
    from ..rule import IconSet
    return IconSet


class TestIconSet:

    def test_create(self, IconSet):
        src = """
        <iconSet iconSet="5Rating" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cfvo type="percent" val="0"/>
          <cfvo type="percentile" val="20"/>
          <cfvo type="percentile" val="40"/>
          <cfvo type="percentile" val="60"/>
          <cfvo type="percentile" val="80"/>
        </iconSet>
        """
        xml = fromstring(src)
        icon = IconSet.from_tree(xml)
        assert icon.iconSet == "5Rating"
        assert len(icon.cfvo) == 5


    def test_serialise(self, IconSet, FormatObject):
        fo1 = FormatObject(type="num", val="2")
        fo2 = FormatObject(type="num", val="4")
        fo3 = FormatObject(type="num", val="6")
        fo4 = FormatObject(type="percent", val="0")
        icon = IconSet(cfvo=[fo1, fo2, fo3, fo4], iconSet="4ArrowsGray", reverse=True, showValue=False)
        xml = tostring(icon.to_tree())
        expected = """
        <iconSet iconSet="4ArrowsGray" showValue="0" reverse="1">
          <cfvo type="num" val="2"/>
          <cfvo type="num" val="4"/>
          <cfvo type="num" val="6"/>
          <cfvo type="percent" val="0"/>
        </iconSet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff



@pytest.fixture
def Rule():
    from ..rule import Rule
    return Rule


class TestRule:

    def test_create(self, Rule, datadir):
        datadir.chdir()
        with open("worksheet.xml") as src:
            xml = fromstring(src.read())

        rules = []
        for el in xml.findall("{%s}conditionalFormatting/{%s}cfRule" % (SHEET_MAIN_NS, SHEET_MAIN_NS)):
            rules.append(Rule.from_tree(el))

        assert len(rules) == 30
        assert rules[17].formula == ('2', '7')
        assert rules[-1].formula == ("AD1>3",)


    def test_serialise(self, Rule):

        rule = Rule(type="cellIs", dxfId="26", priority="13", operator="between")
        rule.formula = ["2", "7"]

        xml = tostring(rule.to_tree())
        expected = """
        <cfRule type="cellIs" dxfId="26" priority="13" operator="between">
        <formula>2</formula>
        <formula>7</formula>
        </cfRule>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


