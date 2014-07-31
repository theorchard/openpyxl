# Copyright (c) 2010-2014 openpyxl

import pytest

from io import BytesIO
import os.path
import datetime
from functools import partial

from openpyxl.formatting import ConditionalFormatting
from openpyxl.formatting.rules import FormulaRule

from openpyxl.styles import (
    Alignment,
    NumberFormat,
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

from openpyxl.xml.functions import Element
from openpyxl.tests.helper import get_xml, compare_xml


class DummyElement:

    def __init__(self):
        self.attrib = {}


class DummyWorkbook:

    style_properties = []


def test_write_gradient_fill():
    fill = GradientFill(degree=90, stop=[Color(theme=0), Color(theme=4)])
    writer = StyleWriter(DummyWorkbook())
    writer._write_gradient_fill(writer._root, fill)
    xml = get_xml(writer._root)
    expected = """<?xml version="1.0" ?>
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
    xml = get_xml(writer._root)
    expected = """<?xml version="1.0" ?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <patternFill patternType="solid">
     <fgColor rgb="00808000" />
  </patternFill>
</styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_borders():
    borders = Border()
    writer = StyleWriter(DummyWorkbook())
    writer._write_border(writer._root, borders)
    xml = get_xml(writer._root)
    expected = """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <border>
    <left/>
    <right/>
    <top/>
    <bottom/>
    <diagonal/>
  </border>
</styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_font():
    wb = DummyWorkbook()
    from openpyxl.styles import Font
    ft = Font(name='Calibri', charset=204, vertAlign='superscript', underline=Font.UNDERLINE_SINGLE)
    writer = StyleWriter(wb)
    writer._write_font(writer._root, ft)
    xml = get_xml(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <font>
          <vertAlign val="superscript"></vertAlign>
          <sz val="11.0"></sz>
          <color rgb="00000000"></color>
          <name val="Calibri"></name>
          <family val="2"></family>
          <u></u>
          <charset val="204"></charset>
         </font>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_number_formats():
    wb = DummyWorkbook()
    from openpyxl.xml.functions import Element
    from openpyxl.styles import NumberFormat, Style
    wb.shared_styles = [
        Style(),
        Style(number_format=NumberFormat('YYYY'))
    ]
    writer = StyleWriter(wb)
    writer._write_number_format(writer._root, 0, "YYYY")
    xml = get_xml(writer._root)
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
            self.worksheet.cell(row=1, column=i).style = Style(font=Font(size=i))
        w = StyleWriter(self.workbook)
        assert len(w.styles) == 6  # 5 + the default

        self.worksheet.cell('A10').style = Style(border=Border(top=Side(border_style=borders.BORDER_THIN)))
        w = StyleWriter(self.workbook)
        assert len(w.styles) == 7


    def test_default_xfs(self):
        w = StyleWriter(self.workbook)
        fonts = nft = borders = fills = DummyElement()
        w._write_cell_xfs(nft, fonts, fills, borders)
        xml = get_xml(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="1">
          <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
        </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_fonts(self):
        st = Style(font=Font(size=12, bold=True))
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)

        nft = borders = fills = DummyElement()
        fonts = Element("fonts")
        w._write_cell_xfs(nft, fonts, fills, borders)
        xml = get_xml(w._root)
        assert """applyFont="1" """ in xml
        assert """fontId="1" """ in xml

        expected = """
        <fonts count="2">
        <font>
            <sz val="12.0" />
            <color rgb="00000000"></color>
            <name val="Calibri" />
            <family val="2" />
            <b></b>
        </font>
        </fonts>
        """
        xml = get_xml(fonts)
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_fills(self):
        st = Style(fill=PatternFill(
            fill_type='solid',
            start_color=Color(colors.DARKYELLOW))
                   )
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)
        nft = borders = fonts = DummyElement()
        fills = Element("fills")
        w._write_cell_xfs(nft, fonts, fills, borders)

        xml = get_xml(w._root)
        expected = """ <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <cellXfs count="2">
          <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
          <xf applyFill="1" borderId="0" fillId="2" fontId="0" numFmtId="0" xfId="0"/>
        </cellXfs>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

        expected = """<fills count="3">
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="00808000"></fgColor>
               </patternFill>
            </fill>
          </fills>
        """
        xml = get_xml(fills)
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xfs_borders(self):
        st = Style(border=Border(top=Side(border_style=borders.BORDER_THIN,
                                              color=Color(colors.DARKYELLOW))))
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)
        fonts = nft = fills = DummyElement()
        border_node = Element("borders")
        w._write_cell_xfs(nft, fonts, fills, border_node)

        xml = get_xml(w._root)
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

        xml = get_xml(border_node)
        expected = """
          <borders count="2">
            <border>
              <left />
              <right />
              <top style="thin">
                <color rgb="00808000"></color>
              </top>
              <bottom />
              <diagonal />
            </border>
          </borders>
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
        xml = get_xml(w._root)
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
        xml = get_xml(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <alignment/>
        </styleSheet>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_rewrite_styles(self):
        """Test to verify Bugfix # 46"""
        self.worksheet['A1'].value = 'Value'
        self.worksheet['B2'].value = '14%'
        saved_wb = save_virtual_workbook(self.workbook)
        second_wb = load_workbook(BytesIO(saved_wb))
        assert isinstance(second_wb, Workbook)
        ws = second_wb.get_sheet_by_name('Sheet1')
        assert ws.cell('A1').value == 'Value'
        ws['A2'].value = 'Bar!'
        saved_wb = save_virtual_workbook(second_wb)
        third_wb = load_workbook(BytesIO(saved_wb))
        assert third_wb

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
        cf.setDxfStyles(self.workbook)
        assert len(self.workbook.style_properties['dxf_list']) == 1
        assert 'font' in self.workbook.style_properties['dxf_list'][0]
        assert 'border' in self.workbook.style_properties['dxf_list'][0]
        assert 'fill' in self.workbook.style_properties['dxf_list'][0]

        w = StyleWriter(self.workbook)
        w._write_dxfs()
        xml = get_xml(w._root)

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
        xml = get_xml(w._root)
        expected = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <protection hidden="1" locked="1"/>
        </styleSheet>
                """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

        nft = fonts = borders = fills = Element('empty')
        w._write_cell_xfs(nft, fonts, fills, borders)
        xml = get_xml(w._root)
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
        st = Style(number_format=NumberFormat(NumberFormat.FORMAT_NUMBER_00),
                   protection=Protection(locked=True))  # 4
        cls.worksheet.cell(coordinate='D9').style = st
        st2 = Style(protection=Protection(hidden=True))  # 5
        cls.worksheet.cell(coordinate='E1').style = st2
        cls.writer = StyleWriter(cls.workbook)

    def test_create_style_table(self):
        assert len(self.writer.styles) == 5

    def test_write_style_table(self, datadir):
        datadir.chdir()
        with open('simple-styles.xml') as reference_file:
            xml = self.writer.write_table()
            diff = compare_xml(xml, reference_file.read())
            assert diff is None, diff


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
