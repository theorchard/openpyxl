import pytest


class TestBarChart:

    def test_ctor(self, BarChart):
        c = BarChart()
        assert c.TYPE == "barChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"

from openpyxl.xml.functions import safe_iterator, fromstring
from openpyxl.xml.constants import CHART_NS
from openpyxl.writer.charts import BarChartWriter

from openpyxl.tests.helper import compare_xml
from openpyxl.tests.schema import chart_schema


@pytest.fixture
def bar_chart(ws, BarChart, Reference, Series):
    ws.title = 'Numbers'
    for i in range(10):
        ws.append([i])
    chart = BarChart()
    chart.add_serie(Series(Reference(ws, (1, 1), (10, 1))))
    return chart


class TestBarChartWriter(object):

    def test_write_chart(self, bar_chart):
        """check if some characteristic tags of LineChart are there"""
        cw = BarChartWriter(bar_chart)
        cw._write_chart()
        tagnames = ['{%s}barChart' % CHART_NS,
                    '{%s}valAx' % CHART_NS,
                    '{%s}catAx' % CHART_NS]
        root = safe_iterator(cw.root)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

    @pytest.mark.lxml_required
    def test_serialised(self, bar_chart, datadir):
        """Check the serialised file against sample"""
        cw = BarChartWriter(bar_chart)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        datadir.chdir()
        with open("BarChart.xml") as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff
