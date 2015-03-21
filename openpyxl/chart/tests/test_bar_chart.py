from __future__ import absolute_import

import pytest

from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def BarChart():
    from ..bar_chart import BarChart
    return BarChart


def test_ctor(BarChart):
    bc = BarChart()
    xml = tostring(bc.to_tree())
    expected = """
    <barChart>
      <barDir val="col" />
      <grouping val="clustered" />

      <axId val="60871424" />
      <axId val="60873344" />
    </barChart>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
