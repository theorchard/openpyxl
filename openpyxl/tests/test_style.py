# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

# Python stdlib imports
from io import BytesIO
import os.path
import datetime
from functools import partial

import pytest

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.reader.style import read_style_table
from openpyxl.workbook import Workbook
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.writer.styles import StyleWriter
from openpyxl.styles import (
    NumberFormat,
    Color,
    Font,
    PatternFill,
    Border,
    Side,
    Protection,
    Style,
    colors,
    fills,
    borders,
)
from openpyxl.formatting import ConditionalFormatting
from openpyxl.formatting.rules import FormulaRule
from openpyxl.xml.functions import Element, SubElement, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS

# test imports
from openpyxl.tests.helper import DATADIR, get_xml, compare_xml
from openpyxl.styles.alignment import Alignment


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
                   protection=Protection(locked=Protection.PROTECTION_UNPROTECTED))  # 4
        cls.worksheet.cell(coordinate='D9').style = st
        st2 = Style(protection=Protection(hidden=Protection.PROTECTION_UNPROTECTED))  # 5
        cls.worksheet.cell(coordinate='E1').style = st2
        cls.writer = StyleWriter(cls.workbook)

    def test_create_style_table(self):
        assert len(self.writer.styles) == 5

    @pytest.mark.xfail
    def test_write_style_table(self):
        reference_file = os.path.join(DATADIR, 'writer', 'expected', 'simple-styles.xml')

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

    def test_style_unicity(self):
        for i in range(1, 6):
            self.worksheet.cell(row=1, column=i).style = Style(font=Font(bold=True))
        w = StyleWriter(self.workbook)
        assert len(w.styles) == 2

    def test_fonts(self):
        st = Style(font=Font(size=12, bold=True))
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)
        w._write_fonts()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts>
            <font>
              <sz val="11" />
              <color theme="1" />
              <name val="Calibri" />
              <family val="2" />
              <scheme val="minor" />
            </font>
          </fonts>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_fonts_with_underline(self):
        st = Style(font=Font(size=12, bold=True,
                             underline=Font.UNDERLINE_SINGLE))
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)
        w._write_fonts()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts>
            <font>
              <sz val="11" />
              <color theme="1" />
              <name val="Calibri" />
              <family val="2" />
              <scheme val="minor" />
            </font>
          </fonts>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_fills(self):
        st = Style(fill=PatternFill(fill_type='solid',
                             start_color=Color(colors.DARKYELLOW)))
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)
        w._write_fills()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="2">
            <fill>
              <patternFill patternType="none" />
            </fill>
            <fill>
              <patternFill patternType="gray125" />
            </fill>
          </fills>
        </styleSheet>
        """)
        assert diff is None, diff

    def test_borders(self):
        st = Style(border=Border(top=Side(border_style=borders.BORDER_THIN,
                                              color=Color(colors.DARKYELLOW))))
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)
        w._write_borders()
        xml = get_xml(w._root)
        diff = compare_xml(xml, """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <borders count="2">
            <border>
              <left />
              <right />
              <top />
              <bottom />
              <diagonal />
            </border>
            <border>
              <left />
              <right />
              <top style="thin">
                <color rgb="0000FF00" />
              </top>
              <bottom />
              <diagonal />
            </border>
          </borders>
        </styleSheet>
        """)
        assert diff is None, diff

    @pytest.mark.parametrize("value, expected",
                             [
                                 (Color('FFFFFF'), {'rgb': 'FFFFFF'}),
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

    def test_write_cell_xfs(self):
        self.worksheet.cell('A1').style = Style(font=Font(size=12))
        w = StyleWriter(self.workbook)
        ft = w._write_fonts()
        nft = w._write_number_formats()
        fills = Element('fills')
        w._write_cell_xfs(nft, ft, fills, {})
        xml = get_xml(w._root)
        assert 'applyFont="1"' in xml
        assert 'applyFill="1"' not in xml
        assert 'applyBorder="1"' not in xml
        assert 'applyAlignment="1"' not in xml

    def test_alignment(self):
        st = Style(alignment=Alignment(horizontal='center', vertical='center'))
        self.worksheet.cell('A1').style = st
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        fonts = fills = Element("empty")
        w._write_cell_xfs(nft, fonts, fills, {})
        xml = get_xml(w._root)
        assert 'applyAlignment="1"' in xml
        assert 'horizontal="center"' in xml
        assert 'vertical="center"' in xml

    def test_alignment_rotation(self):
        self.worksheet.cell('A1').style = Style(alignment=Alignment(vertical='center', text_rotation=90))
        self.worksheet.cell('A2').style = Style(alignment=Alignment(vertical='center', text_rotation=135))
        self.worksheet.cell('A3').style = Style(alignment=Alignment(text_rotation=-34))
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        fonts = fills = Element("empty")
        w._write_cell_xfs(nft,fonts, fills, {})
        xml = get_xml(w._root)
        assert 'textRotation="90"' in xml
        assert 'textRotation="135"' in xml
        assert 'textRotation="124"' in xml

    def test_alignment_indent(self):
        self.worksheet.cell('A1').style = Style(alignment=Alignment(indent=1))
        self.worksheet.cell('A2').style = Style(alignment=Alignment(indent=4))
        self.worksheet.cell('A3').style = Style(alignment=Alignment(indent=0))
        self.worksheet.cell('A3').style = Style(alignment=Alignment(indent=-1))
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        fonts = fills = Element("empty")
        w._write_cell_xfs(nft, fonts, fills, {})
        xml = get_xml(w._root)
        assert 'indent="1"' in xml
        assert 'indent="4"' in xml
        #Indents not greater than zero are ignored when writing
        assert 'indent="0"' not in xml
        assert 'indent="-1"' not in xml

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
        self.worksheet.cell('A1').style = Style(protection=Protection(locked=Protection.PROTECTION_UNPROTECTED,
                                                                      hidden=Protection.PROTECTION_UNPROTECTED))
        w = StyleWriter(self.workbook)
        nft = w._write_number_formats()
        fonts = fills = Element("empty")
        w._write_cell_xfs(nft, fonts, fills, {})
        xml = get_xml(w._root)
        assert 'protection' in xml
        assert 'locked="0"' in xml
        assert 'hidden="0"' in xml
