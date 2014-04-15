import pytest
import os

from openpyxl.xml.functions import Element, fromstring, safe_iterator
from openpyxl.xml.constants import CHART_NS

from openpyxl.writer.charts import (ChartWriter,
                                    PieChartWriter,
                                    LineChartWriter,
                                    BarChartWriter,
                                    ScatterChartWriter,
                                    BaseChartWriter
                                    )
from openpyxl.styles import Color

from .helper import get_xml, DATADIR, compare_xml
from .schema import chart_schema




@pytest.fixture
def line_chart(ws, Reference, Series, LineChart):
    ws.title = 'Line'
    for i in range(1, 5):
        ws.append([i])
    chart = LineChart()
    chart.add_serie(Series(Reference(ws, (0, 0), (4, 0))))
    return chart


class TestLineChartWriter(object):

    def test_write_chart(self, line_chart):
        """check if some characteristic tags of LineChart are there"""
        cw = LineChartWriter(line_chart)
        cw._write_chart()
        tagnames = ['{%s}lineChart' % CHART_NS,
                    '{%s}valAx' % CHART_NS,
                    '{%s}catAx' % CHART_NS]

        root = safe_iterator(cw.root)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

    @pytest.mark.lxml_required
    def test_serialised(self, line_chart):
        """Check the serialised file against sample"""
        cw = LineChartWriter(line_chart)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        expected_file = os.path.join(DATADIR, "writer", "expected", "LineChart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff


@pytest.fixture
def bar_chart_2(ws, BarChart, Reference, Series):
    ws.title = 'Numbers'
    for i in range(10):
        ws.append([i])
    chart = BarChart()
    chart.add_serie(Series(Reference(ws, (0, 0), (9, 0))))
    return chart


class TestBarChartWriter(object):

    def test_write_chart(self, bar_chart_2):
        """check if some characteristic tags of LineChart are there"""
        cw = BarChartWriter(bar_chart_2)
        cw._write_chart()
        tagnames = ['{%s}barChart' % CHART_NS,
                    '{%s}valAx' % CHART_NS,
                    '{%s}catAx' % CHART_NS]
        root = safe_iterator(cw.root)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

    @pytest.mark.lxml_required
    def test_serialised(self, bar_chart_2):
        """Check the serialised file against sample"""
        cw = BarChartWriter(bar_chart_2)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        expected_file = os.path.join(DATADIR, "writer", "expected", "BarChart.xml")
        with open(expected_file) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff
