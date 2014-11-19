# Copyright (c) 2010-2014 openpyxl

import pytest

from io import BytesIO
import datetime

from openpyxl.compat import unicode
from openpyxl.formatting import ConditionalFormatting
from openpyxl.formatting.rules import FormulaRule

from openpyxl.xml.functions import SubElement
from openpyxl.styles import (
    Alignment,
    numbers,
    Color,
    Font,
    GradientFill,
    PatternFill,
    Border,
    Side,
    Protection,
    Style,
    colors,
    fills,
    borders,
)
from openpyxl.reader.excel import load_workbook

from openpyxl.writer.styles import StyleWriter
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.workbook import Workbook

from openpyxl.xml.functions import Element, tostring
from openpyxl.tests.helper import compare_xml


class DummyElement:

    def __init__(self):
        self.attrib = {}


class DummyWorkbook:

    style_properties = []
    _fonts = set()
    _borders = set()


def test_write_gradient_fill():
    fill = GradientFill(degree=90, stop=[Color(theme=0), Color(theme=4)])
    writer = StyleWriter(DummyWorkbook())
    writer._write_gradient_fill(writer._root, fill)
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <gradientFill degree="90" type="linear">
        <stop position="0">
          <color theme="0"/>
        </stop>
        <stop position="1">
          <color theme="4"/>
        </stop>
      </gradientFill>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_pattern_fill():
    fill = PatternFill(fill_type='solid',
                       start_color=Color(colors.DARKYELLOW))
    writer = StyleWriter(DummyWorkbook())
    writer._write_pattern_fill(writer._root, fill)
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <patternFill patternType="solid">
         <fgColor rgb="00808000" />
      </patternFill>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_borders():
    wb = DummyWorkbook()
    wb._borders.add(Border())
    writer = StyleWriter(DummyWorkbook())
    writer._write_border()
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <borders count="1">
      <border>
        <left/>
        <right/>
        <top/>
        <bottom/>
        <diagonal/>
      </border>
      </borders>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_font():
    wb = DummyWorkbook()
    from openpyxl.styles import Font
    ft = Font(name='Calibri', charset=204, vertAlign='superscript', underline=Font.UNDERLINE_SINGLE)
    wb._fonts.add(ft)
    writer = StyleWriter(wb)
    writer._write_font()
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font>
          <vertAlign val="superscript"></vertAlign>
          <sz val="11.0"></sz>
          <color rgb="00000000"></color>
          <name val="Calibri"></name>
          <family val="2"></family>
          <u></u>
          <charset val="204"></charset>
         </font>
    </fonts>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.xfail
def test_write_number_formats():
    wb = DummyWorkbook()
    wb._number_formats = ['YYYY']
    writer = StyleWriter(wb)
    writer._write_number_format()
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
           <numFmt formatCode="YYYY" numFmtId="0"></numFmt>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


class TestStyleWriter(object):

    def setup(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.create_sheet()

    def test_no_style(self):
        w = StyleWriter(self.workbook)
        assert len(w.styles) == 1  # there is always the empty (defaul) style

    def test_nb_style(self):
        for i in range(1, 6):
            cell = self.worksheet.cell(row=1, column=i)
            cell.font = Font(size=i)
            _ = cell.style_id
        w = StyleWriter(self.workbook)
        assert len(w.styles) == 6  # 5 + the default

        cell = self.worksheet.cell('A10')
        cell.border=Border(top=Side(border_style=borders.BORDER_THIN))
        _ = cell.style_id
        w = StyleWriter(self.workbook)
        assert len(w.styles) == 7


    def test_default_xfs(self):
        w = StyleWriter(self.workbook)
        fonts = nft = borders = fills = DummyElement()
        w._write_cell_xfs()
        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="1">
          <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
        </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.xfail
    def test_xfs_number_format(self):
        for o in range(1, 4):
            for i in range(1, 4):
                # Two of these are custom, 0.0% and 0.000%. 0.00% is a built in format ID
                self.worksheet.cell(row=o, column=i).number_format = '0.' + '0' * i + '%'
                # hack
        w = StyleWriter(self.workbook)
        fonts = borders = fills = DummyElement()
        nft = SubElement(w._root, 'numFmts')
        w._write_cell_xfs(nft, fonts, fills, borders)

        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="2">
                <numFmt formatCode="0.0%" numFmtId="165"/>
                <numFmt formatCode="0.000%" numFmtId="166"/>
            </numFmts>
            <cellXfs count="4">
                <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
                <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"/>
                <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="10" xfId="0"/>
                <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="166" xfId="0"/>
            </cellXfs>
        </styleSheet>
        """

        xml = tostring(w._root)
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_fonts(self):
        cell = self.worksheet.cell('A1')
        cell.font = Font(size=12, bold=True)
        _ = cell.style_id # update workbook styles
        w = StyleWriter(self.workbook)

        w._write_cell_xfs()
        xml = tostring(w._root)

        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cellXfs count="2">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf applyFont="1" borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_fills(self):
        cell = self.worksheet.cell('A1')
        cell.fill = fill=PatternFill(fill_type='solid',
                                     start_color=Color(colors.DARKYELLOW))
        _ = cell.style_id # update workbook styles
        w = StyleWriter(self.workbook)
        w._write_cell_xfs()

        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cellXfs count="2">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf applyFill="1" borderId="0" fillId="2" fontId="0" numFmtId="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_borders(self):
        cell = self.worksheet.cell('A1')
        cell.border=Border(top=Side(border_style=borders.BORDER_THIN,
                                    color=Color(colors.DARKYELLOW)))
        _ = cell.style_id # update workbook styles

        w = StyleWriter(self.workbook)
        w._write_cell_xfs()

        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="2">
          <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
          <xf applyBorder="1" borderId="1" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
        </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.parametrize("value, expected",
                             [
                                 (Color('FFFFFF'), {'rgb': '00FFFFFF'}),
                                 (Color(indexed=7), {'indexed': '7'}),
                                 (Color(theme=7, tint=0.8), {'theme':'7', 'tint':'0.8'}),
                                 (Color(auto=True), {'auto':'1'}),
                             ])
    def test_write_color(self, value, expected):
        w = StyleWriter(self.workbook)
        root = Element("root")
        w._write_color(root, value)
        assert root.find('color') is not None
        assert root.find('color').attrib == expected


    def test_alignment(self):
        w = StyleWriter(self.workbook)
        al = Alignment(horizontal='center', vertical='center')
        w._write_alignment(w._root, al)
        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <alignment horizontal="center" vertical="center"/>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_alignment_default(self):
        w = StyleWriter(self.workbook)
        al = Alignment()
        w._write_alignment(w._root, al)
        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <alignment/>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write_dxf(self):
        redFill = PatternFill(start_color=Color('FFEE1111'),
                       end_color=Color('FFEE1111'),
                       fill_type=fills.FILL_SOLID)
        whiteFont = Font(color=Color("FFFFFFFF"),
                         bold=True, italic=True, underline='single',
                         strikethrough=True)
        medium_blue = Side(border_style='medium', color=Color(colors.BLUE))
        blueBorder = Border(left=medium_blue,
                             right=medium_blue,
                             top=medium_blue,
                             bottom=medium_blue)
        cf = ConditionalFormatting()
        cf.add('A1:A2', FormulaRule(formula="[A1=1]", font=whiteFont, border=blueBorder, fill=redFill))
        cf._save_styles(self.workbook)
        assert len(self.workbook.style_properties['dxf_list']) == 1
        assert 'font' in self.workbook.style_properties['dxf_list'][0]
        assert 'border' in self.workbook.style_properties['dxf_list'][0]
        assert 'fill' in self.workbook.style_properties['dxf_list'][0]

        w = StyleWriter(self.workbook)
        w._write_dxfs()
        xml = tostring(w._root)

        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dxfs count="1">
            <dxf>
              <font>
                <color rgb="FFFFFFFF" />
                <b val="1" />
                <i val="1" />
                <u val="single" />
                <strike />
              </font>
              <fill>
                <patternFill patternType="solid">
                  <fgColor rgb="FFEE1111" />
                  <bgColor rgb="FFEE1111" />
                </patternFill>
              </fill>
              <border>
                <left style="medium">
                    <color rgb="000000FF"></color>
                </left>
                <right style="medium">
                    <color rgb="000000FF"></color>
                </right>
                <top style="medium">
                    <color rgb="000000FF"></color>
                </top>
                <bottom style="medium">
                    <color rgb="000000FF"></color>
                </bottom>
            </border>
            </dxf>
          </dxfs>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_protection(self):
        prot = Protection(locked=True,
                          hidden=True)
        self.worksheet.cell('A1').style = Style(protection=prot)
        w = StyleWriter(self.workbook)
        w._write_protection(w._root, prot)
        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <protection hidden="1" locked="1"/>
        </styleSheet>
                """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

        nft = fonts = borders = fills = Element('empty')
        w._write_cell_xfs(nft, fonts, fills, borders)
        xml = tostring(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <protection hidden="1" locked="1"/>
          <cellXfs count="2">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf applyProtection="1" borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0">
              <protection hidden="1" locked="1"/>
            </xf>
          </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestCreateStyle(object):

    @classmethod
    def setup_class(cls):
        now = datetime.datetime.now()
        cls.workbook = Workbook(guess_types=True)
        cls.worksheet = cls.workbook.create_sheet()
        cls.worksheet.cell(coordinate='A1').value = '12.34%'  # 2
        cls.worksheet.cell(coordinate='B4').value = now  # 3
        cls.worksheet.cell(coordinate='B5').value = now
        cls.worksheet.cell(coordinate='C14').value = 'This is a test'  # 1
        cls.worksheet.cell(coordinate='D9').value = '31.31415'  # 3
        st = Style(number_format=numbers.FORMAT_NUMBER_00,
                   protection=Protection(locked=True))  # 4
        cls.worksheet.cell(coordinate='D9').style = st
        st2 = Style(protection=Protection(hidden=True))  # 5
        cls.worksheet.cell(coordinate='E1').style = st2
        cls.writer = StyleWriter(cls.workbook)

    @pytest.mark.xfail
    def test_create_style_table(self):
        assert len(self.writer.styles) == 5

    @pytest.mark.xfail
    def test_write_style_table(self, datadir):
        datadir.chdir()
        with open('simple-styles.xml') as reference_file:
            xml = self.writer.write_table()
            diff = compare_xml(xml, reference_file.read())
            assert diff is None, diff


def test_empty_workbook():
    wb = Workbook()
    writer = StyleWriter(wb)
    expected = """
    <styleSheet>
      <numFmts count="0"/>
      <fonts count="1">
        <font>
          <sz val="11.0"/>
          <color rgb="00000000"/>
          <name val="Calibri"/>
          <family val="2"/>
        </font>
      </fonts>
      <fills count="2">
       <fill>
          <patternFill patternType="none" />
       </fill>
       <fill>
          <patternFill patternType="gray125"/>
        </fill>
      </fills>
      <borders count="1">
        <border>
          <left/>
          <right/>
          <top/>
          <bottom/>
          <diagonal/>
        </border>
      </borders>
      <cellStyleXfs count="1">
        <xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>
      </cellStyleXfs>
      <cellXfs count="1">
        <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
      </cellXfs>
      <cellStyles count="1">
        <cellStyle builtinId="0" name="Normal" xfId="0"/>
      </cellStyles>
      <dxfs count="0"/>
      <tableStyles count="0" defaultPivotStyle="PivotStyleLight16" defaultTableStyle="TableStyleMedium9"/>
    </styleSheet>
    """
    xml = writer.write_table()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.xfail
def test_complex_styles(datadir):
    """Hold on to your hats"""
    from openpyxl import load_workbook
    datadir.join("..", "..", "..", "reader", "tests", "data").chdir()
    wb = load_workbook("complex-styles.xlsx")

    datadir.chdir()
    with open("complex-styles.xml") as reference:
        writer = StyleWriter(wb)
        xml = writer.write_table()
        expected = reference.read()
        diff = compare_xml(xml, expected)
        assert diff is None, diff
