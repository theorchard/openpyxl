import pytest


class TestPieChart:

    def test_ctor(self, PieChart):
        c = PieChart()
        assert c.TYPE, "pieChart"


@pytest.fixture
def pie_chart(ws, Reference, Series, PieChart):
    ws.title = 'Pie'
    for i in range(1, 5):
        ws.append([i])
    chart = PieChart()
    values = Reference(ws, (1, 1), (10, 1))
    series = Series(values, labels=values)
    chart.add_serie(series)
    return chart

import os
from openpyxl.writer.charts import PieChartWriter
from openpyxl.xml.constants import CHART_NS
from openpyxl.xml.functions import safe_iterator, fromstring
from openpyxl.tests.schema import chart_schema

from openpyxl.tests.helper import get_xml, compare_xml

class TestPieChartWriter(object):

    def test_write_chart(self, pie_chart):
        """check if some characteristic tags of PieChart are there"""
        cw = PieChartWriter(pie_chart)
        cw._write_chart()

        tagnames = ['{%s}pieChart' % CHART_NS,
                    '{%s}varyColors' % CHART_NS
                    ]
        root = safe_iterator(cw.root)
        chart_tags = [e.tag for e in root]
        for tag in tagnames:
            assert tag in chart_tags

        assert 'c:catAx' not in chart_tags

    @pytest.mark.lxml_required
    def test_serialised(self, pie_chart, datadir):
        """Check the serialised file against sample"""
        cw = PieChartWriter(pie_chart)
        xml = cw.write()
        tree = fromstring(xml)
        chart_schema.assertValid(tree)
        datadir.chdir()
        with open("PieChart.xml") as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff
