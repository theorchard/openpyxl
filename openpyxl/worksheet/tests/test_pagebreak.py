from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.functions import tostring


@pytest.fixture
def Break():
    from ..pagebreak import Break
    return Break


@pytest.fixture
def PageBreak():
    from ..pagebreak import PageBreak
    return PageBreak



class TestBreak:

    def test_ctor(self, Break):
        brk = Break()
        assert dict(brk) == {'id': '0', 'man': '1', 'max': '16383', 'min': '0'}
        xml = tostring(brk.to_tree())
        expected = """
        <brk id="0" man="1" max="16383" min="0"></brk>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestPageBreak:

    def test_no_brks(self, PageBreak):
        pb = PageBreak()
        assert dict(pb) == {'count': '0', 'manualBreakCount': '0'}

    def test_brk(self, PageBreak):
        pb = PageBreak()
        pb.append()
        assert dict(pb) == {'count': '1', 'manualBreakCount': '1'}
        xml = tostring(pb.to_tree())
        expected = """
        <rowBreaks count="1" manualBreakCount="1">
           <brk id="1" man="1" max="16383" min="0"></brk>
        </rowBreaks>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

