import pytest


class TestLineChart:

    def test_ctor(self, LineChart):
        c = LineChart()
        assert c.TYPE == "lineChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"


from openpyxl.writer.charts import LineChartWriter
from openpyxl.xml.functions import safe_iterator, fromstring
from openpyxl.xml.constants import CHART_NS

from openpyxl.tests.schema import chart_schema
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def line_chart(ws, Reference, Series, LineChart):
    ws.title = 'Line'
    for i in range(1, 5):
        ws.append([i])
    chart = LineChart()
    chart.add_serie(Series(Reference(ws, (1, 1), (5, 1))))
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
    def test_serialised(self, line_chart, datadir):
        """Check the serialised file against sample"""
        cw = LineChartWriter(line_chart)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        datadir.chdir()
        with open("LineChart.xml") as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff
