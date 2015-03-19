from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def CatAx():
    from ..axis import CatAx
    return CatAx


class TestCatAx:

    def test_ctor(self, CatAx):
        axis = CatAx()
        xml = tostring(axis.to_tree())
        expected = """
        <catAx>
            <lblOffset val="100"></lblOffset>
        </catAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


@pytest.fixture
def ValAx():
    from ..axis import ValAx
    return ValAx


class TestValAx:

    def test_ctor(self, ValAx):
        axis = ValAx()
        xml = tostring(axis.to_tree())
        expected = """
        <valAx>
        </valAx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
