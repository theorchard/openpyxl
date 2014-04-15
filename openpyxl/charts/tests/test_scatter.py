import pytest


class TestScatterChart:

    def test_ctor(self, ScatterChart):
        c = ScatterChart()
        assert c.TYPE == "scatterChart"
        assert c.x_axis.type == "valAx"
        assert c.y_axis.type == "valAx"


@pytest.fixture()
def scatter_chart(ws, ScatterChart, Reference, Series):
    ws.title = 'Scatter'
    for i in range(10):
        ws.cell(row=i+1, column=1).value = i
        ws.cell(row=i+1, column=2).value = i
    chart = ScatterChart()
    chart.add_serie(Series(Reference(ws, (1, 1), (11, 1)),
                                      xvalues=Reference(ws, (1, 2), (11, 2))))
    return chart

import os
from openpyxl.xml.constants import CHART_NS
from openpyxl.writer.charts import ScatterChartWriter

from openpyxl.tests.helper import get_xml, compare_xml


class TestScatterChartWriter(object):

    def test_write_xaxis(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        scatter_chart.x_axis.title = 'test x axis title'
        cw._write_axis(root_xml, scatter_chart.x_axis, '{%s}valAx' % CHART_NS)

        expected = """<test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>test x axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_yaxis(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        scatter_chart.y_axis.title = 'test y axis title'
        cw._write_axis(root_xml, scatter_chart.y_axis, '{%s}valAx' % CHART_NS)

        expected = """<test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>test y axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_series(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_series(root_xml)

        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ser><c:idx val="0" /><c:order val="0" /><c:xVal><c:numRef><c:f>\'Scatter\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'Scatter\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_legend(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_legend(root_xml)
        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_print_settings(self, scatter_chart):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_print_settings()

        expected = """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></c:chartSpace>"""
        xml = get_xml(cw.root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_chart(self, scatter_chart):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_chart()
        xml = get_xml(cw.root)
        expected = """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.03375" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:scatterChart><c:scatterStyle val="lineMarker" /><c:ser><c:idx val="0" /><c:order val="0" /><c:xVal><c:numRef><c:f>\'Scatter\'!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\'Scatter\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser><c:axId val="60871424" /><c:axId val="60873344" /></c:scatterChart><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></c:chartSpace>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_serialised(self, scatter_chart, datadir):
        cw = ScatterChartWriter(scatter_chart)
        xml = cw.write()

        #fname = os.path.join(DATADIR, "writer", "expected", "ScatterChart.xml")
        datadir.chdir()
        fname = "ScatterChart.xml"
        with open(fname) as expected:
            diff = compare_xml(xml, expected.read())
            assert diff is None, diff

