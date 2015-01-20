from __future__ import absolute_import
# copyright 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.tests.helper import compare_xml


@pytest.fixture
def FormatObject():
    from ..rule import FormatObject
    return FormatObject


class TestFormatObject:

    def test_create(self, FormatObject):
        xml = fromstring("""<cfvo type="num" val="3"/>""")
        cfvo = FormatObject.create(xml)
        assert cfvo.type == "num"
        assert cfvo.val == "3"
        assert cfvo.gte is None


    def test_serialise(self, FormatObject):
        cfvo = FormatObject(type="percent", val="4")
        xml = tostring(cfvo.serialise())
        expected = """<cfvo type="percent" val="4"/>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestColorScale:

    def test_create(self):
        pass


    def test_serialise(self):
        pass


class TestDataBar:

    def test_create(self):
        pass


    def test_serialise(self):
        pass


class TestDataBar:

    def test_create(self):
        pass


    def test_serialise(self):
        pass


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
            rules.append(Rule.create(el))

        assert len(rules) == 30
        assert rules[17].formula == ('2', '7')
        assert rules[-1].formula == ("AD1>3",)


    def test_serialise(self, Rule):

        rule = Rule(type="cellIs", dxfId="26", priority="13", operator="between")
        rule.formula = ["2", "7"]

        xml = tostring(rule.serialise())
        expected = """
        <cfRule type="cellIs" dxfId="26" priority="13" operator="between">
        <formula>2</formula>
        <formula>7</formula>
        </cfRule>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


