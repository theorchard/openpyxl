from __future__ import absolute_import
# copyright 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Rule():
    from ..rule import Rule
    return Rule


def test_create(Rule, datadir):
    datadir.chdir()
    with open("worksheet.xml") as src:
        xml = fromstring(src.read())

    rules = []
    for el in xml.findall("{%s}conditionalFormatting/{%s}cfRule" % (SHEET_MAIN_NS, SHEET_MAIN_NS)):
        rules.append(Rule.create(el))

    assert len(rules) == 30
    assert rules[17].formula == ('2', '7')
    assert rules[-1].formula == ("AD1>3",)


def test_serialise(Rule):

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
