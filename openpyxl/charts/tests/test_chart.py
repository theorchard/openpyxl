import pytest


class TestChart:

    def test_ctor(self, Chart):
        from openpyxl.charts import Legend
        from openpyxl.drawing import Drawing
        c = Chart()
        assert c.TYPE == None
        assert c.GROUPING == "standard"
        assert isinstance(c.legend, Legend)
        assert c.show_legend
        assert c.lang == 'en-GB'
        assert c.title == ''
        assert c.print_margins == {'b':0.75, 'l':0.7, 'r':0.7, 't':0.75,
                                   'header':0.3, 'footer':0.3}
        assert isinstance(c.drawing, Drawing)
        assert c.width == 0.6
        assert c.height == 0.6
        assert c.margin_top == 0.31
        assert c.series == []
        assert c.shapes == []
        with pytest.raises(ValueError):
            assert c.margin_left == 0

    def test_mymax(self, Chart):
        c = Chart()
        assert c.mymax(range(10)) == 9
        from string import ascii_letters as letters
        assert c.mymax(list(letters)) == "z"
        assert c.mymax(range(-10, 1)) == 0
        assert c.mymax([""]*10) == ""

    def test_mymin(self, Chart):
        c = Chart()
        assert c.mymin(range(10)) == 0
        from string import ascii_letters as letters
        assert c.mymin(list(letters)) == "A"
        assert c.mymin(range(-10, 1)) == -10
        assert c.mymin([""]*10) == ""

    def test_margin_top(self, Chart):
        c = Chart()
        assert c.margin_top == 0.31

    def test_margin_left(self, series, Chart):
        c = Chart()
        c.append(series)
        assert c.margin_left == 0.03375

    def test_set_margin_top(self, Chart):
        c = Chart()
        c.margin_top = 1
        assert c.margin_top == 0.31

    def test_set_margin_left(self, series, Chart):
        c = Chart()
        c.append(series)
        c.margin_left = 0
        assert c.margin_left  == 0.03375


@pytest.fixture
def bar_chart(ten_row_sheet, BarChart, Series, Reference):
    from openpyxl.styles.colors import GREEN
    ws = ten_row_sheet
    chart = BarChart()
    chart.title = "TITLE"
    series = Series(Reference(ws, (1, 1), (11, 1)))
    series.color = GREEN
    chart.add_serie(series)
    return chart


from openpyxl.writer.charts import BarChartWriter, BaseChartWriter
from openpyxl.xml.constants import CHART_NS
from openpyxl.xml.functions import Element, fromstring, safe_iterator

from openpyxl.tests.helper import get_xml, compare_xml
from openpyxl.tests.schema import chart_schema


def test_write_serial(ten_row_sheet, LineChart, Series, Reference, root_xml):
    ws = ten_row_sheet
    chart = LineChart()
    for idx, l in enumerate("ABCDEF", 1):
        ws.cell(row=idx, column=1).value = l
    ref = Reference(ws, (1, 1), (10, 1))
    series = Series(ref)
    chart.add_serie(series)
    cw = BaseChartWriter(chart)
    cw._write_serial(cw.root, ref)
    xml = get_xml(cw.root)
    expected = """ <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:strRef><c:f>'data'!$A$1:$A$10</c:f><c:strCache><c:ptCount val="10"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt><c:pt idx="4"><c:v>E</c:v></c:pt><c:pt idx="5"><c:v>F</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:strCache></c:strRef></c:chartSpace>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


class TestChartWriter(object):

    def test_write_title(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_title(root_xml)
        expected = """<?xml version='1.0' ?><test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_xaxis(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_axis(root_xml, bar_chart.x_axis, '{%s}catAx' % CHART_NS)
        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_yaxis(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_axis(root_xml, bar_chart.y_axis, '{%s}valAx' % CHART_NS)
        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_series(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_series(root_xml)
        expected = """<?xml version='1.0' ?><test xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:val><c:numRef><c:f>\'data\'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_legend(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        cw._write_legend(root_xml)
        expected = """<?xml version='1.0' ?><test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_no_write_legend(self, bar_chart, root_xml):
        cw = BarChartWriter(bar_chart)
        bar_chart.show_legend = False
        cw._write_legend(root_xml)
        children = [e for e in root_xml]
        assert len(children) == 0

    def test_write_print_settings(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        cw._write_print_settings()
        tagnames = ['test',
                    '{%s}printSettings' % CHART_NS,
                    '{%s}headerFooter' % CHART_NS,
                    '{%s}pageMargins' % CHART_NS,
                    '{%s}pageSetup' % CHART_NS]
        for e in cw.root:
            assert e.tag in tagnames
            if e.tag == "{%s}pageMargins" % CHART_NS:
                assert e.keys() == list(bar_chart.print_margins.keys())
                for k, v in e.items():
                    assert float(v) == bar_chart.print_margins[k]
            else:
                assert e.text == None
                assert e.attrib == {}

    @pytest.mark.lxml_required
    def test_write_chart(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        cw._write_chart()
        assert chart_schema.validate(cw.root)

        expected = """<c:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="en-GB" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="0.03375" /><c:y val="0.31" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:val><c:numRef><c:f>'data'!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="60871424" /><c:axId val="60873344" /></c:barChart><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></c:chartSpace>"""

        xml = get_xml(cw.root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_rels(self, bar_chart):
        cw = BarChartWriter(bar_chart)
        xml = cw.write_rels(1)
        expected = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/drawing1.xml"/></Relationships>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_no_ascii(self, ten_row_sheet, Series, BarChart, Reference, root_xml):
        from openpyxl.writer.charts import ChartWriter
        ws = ten_row_sheet
        ws.append([b"D\xc3\xbcsseldorf"]*10)
        serie = Series(values=Reference(ws, (1,1), (1,10)),
                      title=(ws.cell(row=1, column=1).value)
                      )
        c = BarChart()
        c.add_serie(serie)
        cw = BaseChartWriter(c)
        cw._write_series(root_xml)
        xml = get_xml(root_xml)
        expected = """<test><c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="10"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>None</c:v></c:pt><c:pt idx="2"><c:v>None</c:v></c:pt><c:pt idx="3"><c:v>None</c:v></c:pt><c:pt idx="4"><c:v>None</c:v></c:pt><c:pt idx="5"><c:v>None</c:v></c:pt><c:pt idx="6"><c:v>None</c:v></c:pt><c:pt idx="7"><c:v>None</c:v></c:pt><c:pt idx="8"><c:v>None</c:v></c:pt><c:pt idx="9"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></test>"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_label_no_number_format(self, ten_column_sheet, Reference, Series, BarChart, root_xml):
        ws = ten_column_sheet
        for i in range(10):
            ws.append([i, i])
        labels = Reference(ws, (1,1), (1,10))
        values = Reference(ws, (1,1), (1,10))
        serie = Series(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        cw._write_serial(root_xml, c.series[0].labels)
        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>"""
        xml = get_xml(root_xml)
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_label_number_format(self, ten_column_sheet, Reference, Series, BarChart):
        ws = ten_column_sheet
        labels = Reference(ws, (1,1), (1,10))
        labels.number_format = 'd-mmm'
        values = Reference(ws, (1,1), (1,10))
        serie = Series(values=values, labels=labels)
        c = BarChart()
        c.add_serie(serie)
        cw = BarChartWriter(c)
        root = Element('test')
        cw._write_serial(root, c.series[0].labels)

        expected = """<test xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numRef><c:f>'data'!$A$1:$J$1</c:f><c:numCache><c:formatCode>d-mmm</c:formatCode><c:ptCount val="10" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt></c:numCache></c:numRef></test>"""

        xml = get_xml(root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff
