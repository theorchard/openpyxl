import pytest

from openpyxl.xml.functions import tostring
from openpyxl.xml.constants import CHART_NS
from openpyxl.charts.writer import ScatterChartWriter

from openpyxl.tests.helper import compare_xml


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




class TestScatterChartWriter(object):

    def test_write_xaxis(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        scatter_chart.x_axis.title = 'test x axis title'
        cw._write_axis(root_xml, scatter_chart.x_axis, '{%s}valAx' % CHART_NS)

        expected = """
        <test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:valAx>
            <c:axId val="60871424"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:axPos val="b"/>
            <c:majorGridlines/>
            <c:numFmt formatCode="General" sourceLinked="0"/>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr/>
                    </a:pPr>
                    <a:r>
                      <a:rPr lang="en-GB"/>
                      <a:t>test x axis title</a:t>
                    </a:r>
                  </a:p>
                </c:rich>
              </c:tx>
              <c:layout/>
            </c:title>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="60873344"/>
            <c:crosses val="autoZero"/>
            <c:auto val="1"/>
            <c:lblAlgn val="ctr"/>
            <c:lblOffset val="100"/>
            <c:crossBetween val="midCat"/>
          </c:valAx>
        </test>
        """
        xml = tostring(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_yaxis(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        scatter_chart.y_axis.title = 'test y axis title'
        cw._write_axis(root_xml, scatter_chart.y_axis, '{%s}valAx' % CHART_NS)

        expected = """
        <test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:valAx>
            <c:axId val="60873344"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:axPos val="l"/>
            <c:majorGridlines/>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr/>
                    </a:pPr>
                    <a:r>
                      <a:rPr lang="en-GB"/>
                      <a:t>test y axis title</a:t>
                    </a:r>
                  </a:p>
                </c:rich>
              </c:tx>
              <c:layout/>
            </c:title>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="60871424"/>
            <c:crosses val="autoZero"/>
            <c:crossBetween val="midCat"/>
          </c:valAx>
        </test>
        """
        xml = tostring(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_series(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_series(root_xml)

        expected = """
        <test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:xVal>
              <c:numRef>
                <c:f>'Scatter'!$B$1:$B$11</c:f>
              </c:numRef>
            </c:xVal>
            <c:yVal>
              <c:numRef>
                <c:f>'Scatter'!$A$1:$A$11</c:f>
              </c:numRef>
            </c:yVal>
          </c:ser>
        </test>
        """
        xml = tostring(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_legend(self, scatter_chart, root_xml):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_legend(root_xml)
        expected = """
        <test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:legend>
            <c:legendPos val="r"/>
            <c:layout/>
          </c:legend>
        </test>
        """
        xml = tostring(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_print_settings(self, scatter_chart):
        cw = ScatterChartWriter(scatter_chart)
        cw._write_print_settings()

        expected = """
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:printSettings>
            <c:headerFooter/>
            <c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75"/>
            <c:pageSetup/>
          </c:printSettings>
        </c:chartSpace>
        """
        xml = tostring(cw.root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_serialised(self, scatter_chart, datadir):
        cw = ScatterChartWriter(scatter_chart)
        xml = cw.write()

        datadir.chdir()
        fname = "ScatterChart.xml"
        with open(fname) as expected:
            diff = compare_xml(xml, expected.read())
        assert diff is None, diff

