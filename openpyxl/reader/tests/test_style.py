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
import datetime
from functools import partial
from io import BytesIO

import pytest

# package imports
from openpyxl.reader.excel import load_workbook
from openpyxl.reader.style import read_style_table
from openpyxl.writer.excel import save_virtual_workbook

from openpyxl.styles import (
    numbers,
    Color,
    Font,
    PatternFill,
    GradientFill,
    Border,
    Side,
    Protection,
    Style
)
from openpyxl.styles import colors, fills, borders

# test imports
from openpyxl.tests.helper import get_xml, compare_xml
from openpyxl.styles.alignment import Alignment


@pytest.fixture
def StyleReader():
    from ..style import SharedStylesParser
    return SharedStylesParser


def test_read_pattern_fill(StyleReader, datadir):
    datadir.chdir()
    expected = [
        PatternFill(),
        PatternFill(fill_type='gray125'),
        PatternFill(fill_type='solid',
             start_color=Color(theme=0, tint=-0.14999847407452621),
             end_color=Color(indexed=64)
             ),
        PatternFill(fill_type='solid',
             start_color=Color(theme=0),
             end_color=Color(indexed=64)
             ),
        PatternFill(fill_type='solid',
             start_color=Color(indexed=62),
             end_color=Color(indexed=64)
             )
    ]
    with open("bug311-styles.xml") as src:
        reader = StyleReader(src.read())
        for val, exp in zip(reader.parse_fills(), expected):
            assert val == exp


def test_read_gradient_fill(StyleReader, datadir):
    datadir.chdir()
    expected = [
        GradientFill(degree=90, stop=(Color(theme=0), Color(theme=4)))
    ]
    with open("bug284-styles.xml") as src:
        reader = StyleReader(src.read())
        assert list(reader.parse_fills()) == expected


def test_read_cell_style(datadir):
    datadir.chdir()
    with open("empty-workbook-styles.xml") as content:
        style_properties = read_style_table(content.read())
        assert len(style_properties) == 3


def test_read_simple_style_mappings(datadir):
    datadir.chdir()
    with open("simple-styles.xml") as content:
        style_properties = read_style_table(content.read())[0]
        assert len(style_properties) == 4
        assert numbers.BUILTIN_FORMATS[9] == style_properties[1].number_format
        assert 'yyyy-mm-dd' == style_properties[2].number_format


def test_read_complex_style_mappings(datadir):
    datadir.chdir()
    with open("complex-styles.xml") as content:
        style_properties = read_style_table(content.read())[0]
        assert len(style_properties) == 29
        assert style_properties[-1].font.bold is False


def test_read_complex_style(datadir):
    datadir.chdir()
    wb = load_workbook("complex-styles.xlsx")
    ws = wb.get_active_sheet()
    assert ws.column_dimensions['A'].width == 31.1640625

    assert ws.get_style('I').font == Font(sz=12.0, color='FF3300FF')
    assert ws.get_style('I').fill == PatternFill(patternType='solid', fgColor='FF006600', bgColor=Color(indexed=64))

    assert ws['A2'].font == Font(sz=10, name='Arial', color=Color(theme=1))
    assert ws['A3'].font == Font(sz=12, name='Arial', bold=True, color=Color(theme=1))
    assert ws['A4'].font == Font(sz=14, name='Arial', italic=True, color=Color(theme=1))

    assert ws['A5'].font.color.value == 'FF3300FF'
    assert ws['A6'].font.color.value == 9
    assert ws['A7'].fill.start_color.value == 'FFFFFF66'
    assert ws['A8'].fill.start_color.value == 8
    assert ws['A9'].alignment.horizontal == 'left'
    assert ws['A10'].alignment.horizontal == 'right'
    assert ws['A11'].alignment.horizontal == 'center'
    assert ws['A12'].alignment.vertical == 'top'
    assert ws['A13'].alignment.vertical == 'center'
    assert ws['A14'].alignment.vertical == 'bottom'
    assert ws['A15'].number_format == '0.00'
    assert ws['A16'].number_format == 'mm-dd-yy'
    assert ws['A17'].number_format == '0.00%'

    assert 'A18:B18' in ws._merged_cells
    assert ws.cell('B18').merged

    assert ws['A19'].border == Border(
        left=Side(style='thin', color='FF006600'),
        top=Side(style='thin', color='FF006600'),
        right=Side(style='thin', color='FF006600'),
        bottom=Side(style='thin', color='FF006600'),
    )

    assert ws['A21'].border == Border(
        left=Side(style='double', color=Color(theme=7)),
        top=Side(style='double', color=Color(theme=7)),
        right=Side(style='double', color=Color(theme=7)),
        bottom=Side(style='double', color=Color(theme=7)),
    )

    assert ws['A23'].fill == PatternFill(patternType='solid', start_color='FFCCCCFF', end_color=(Color(indexed=64)))
    assert ws['A23'].border.top == Side(style='mediumDashed', color=Color(theme=6))

    assert 'A23:B24' in ws._merged_cells

    assert ws['A25'].alignment == Alignment(wrapText=True)
    assert ws['A26'].alignment == Alignment(shrinkToFit=True)


def test_change_existing_styles(datadir):
    wb = load_workbook("complex-styles.xlsx")
    ws = wb.get_active_sheet()

    ws.column_dimensions['A'].width = 20
    i_style = ws.get_style('I')
    ws.set_style('I', i_style.copy(fill=PatternFill(fill_type='solid',
                                             start_color=Color('FF442200')),
                                   font=Font(color=Color('FF002244'))))

    assert ws.get_style('I').fill.start_color.value == 'FF442200'
    assert ws.get_style('I').font.color.value == 'FF002244'

    ws.cell('A2').style = ws.cell('A2').style.copy(font=Font(name='Times New Roman',
                                                             size=12,
                                                             bold=True,
                                                             italic=True))
    assert ws['A2'].font == Font(name='Times New Roman', size=12, bold=True,
                                 italic=True)

    ws.cell('A3').style = ws.cell('A3').style.copy(font=Font(name='Times New Roman',
                                                             size=14,
                                                             bold=False,
                                                             italic=True))
    assert ws['A3'].font == Font(name='Times New Roman', size=14,
                                    bold=False, italic=True)


    ws.cell('A4').style = ws.cell('A4').style.copy(font=Font(name='Times New Roman',
                                                             size=16,
                                                             bold=True,
                                                             italic=False))
    assert ws['A4'].font == Font(name='Times New Roman', size=16, bold=True,
                                 italic=False)

    ws.cell('A5').style = ws.cell('A5').style.copy(font=Font(color=Color('FF66FF66')))
    assert ws['A5'].font == Font(color='FF66FF66')


    ws.cell('A6').style = ws.cell('A6').style.copy(font=Font(color=Color(theme='1')))
    assert ws['A6'].font == Font(color=Color(theme='1'))

    ws.cell('A7').style = ws.cell('A7').style.copy(fill=PatternFill(fill_type='solid',
                                                             start_color=Color('FF330066')))
    assert ws['A7'].fill == PatternFill(fill_type='solid',
                                        start_color=Color('FF330066'))

    ws.cell('A8').style = ws.cell('A8').style.copy(fill=PatternFill(fill_type='solid',
                                                             start_color=Color(theme='2')))
    assert ws['A8'].fill == PatternFill(fill_type='solid',
                                        start_color=Color(theme='2'))

    ws.cell('A9').style = ws.cell('A9').style.copy(alignment=Alignment(horizontal='center'))
    assert ws['A9'].alignment == Alignment(horizontal='center')

    ws.cell('A10').style = ws.cell('A10').style.copy(alignment=Alignment(horizontal='left'))
    assert ws['A10'].alignment == Alignment(horizontal='left')

    ws.cell('A11').style = ws.cell('A11').style.copy(alignment=Alignment(horizontal='right'))
    assert ws['A11'].alignment == Alignment(horizontal='right')

    ws.cell('A12').style = ws.cell('A12').style.copy(alignment=Alignment(vertical='bottom'))
    assert ws['A12'].alignment == Alignment(vertical='bottom')

    ws.cell('A13').style = ws.cell('A13').style.copy(alignment=Alignment(vertical='top'))
    assert ws['A13'].alignment == Alignment(vertical='top')

    ws.cell('A14').style = ws.cell('A14').style.copy(alignment=Alignment(vertical='center'))
    assert ws['A14'].alignment == Alignment(vertical='center')

    ws.cell('A15').style = ws.cell('A15').style.copy(number_format='0.00%')
    assert ws['A15'].number_format == '0.00%'

    ws.cell('A16').style = ws.cell('A16').style.copy(number_format='0.00')
    assert ws['A16'].number_format == '0.00'

    ws.cell('A17').style = ws.cell('A17').style.copy(number_format='mm-dd-yy')
    assert ws['A17'].number_format == 'mm-dd-yy'

    ws.unmerge_cells('A18:B18')

    ws.cell('A19').style = ws.cell('A19').style.copy(border=Border(top=Side(border_style=borders.BORDER_THIN,
                                                                                color=Color('FF006600')),
                                                                     bottom=Side(border_style=borders.BORDER_THIN,
                                                                                   color=Color('FF006600')),
                                                                     left=Side(border_style=borders.BORDER_THIN,
                                                                                 color=Color('FF006600')),
                                                                     right=Side(border_style=borders.BORDER_THIN,
                                                                                  color=Color('FF006600'))))
    assert ws['A19'].border == Border(
        top=Side(border_style=borders.BORDER_THIN, color='FF006600'),
        bottom=Side(border_style=borders.BORDER_THIN, color='FF006600'),
        left=Side(border_style=borders.BORDER_THIN, color='FF006600'),
        right=Side(border_style=borders.BORDER_THIN, color='FF006600'))

    ws.cell('A21').style = ws.cell('A21').style.copy(border=Border(top=Side(border_style=borders.BORDER_THIN,
                                                                                color=Color(theme=7)),
                                                                     bottom=Side(border_style=borders.BORDER_THIN,
                                                                                   color=Color(theme=7)),
                                                                     left=Side(border_style=borders.BORDER_THIN,
                                                                                 color=Color(theme=7)),
                                                                     right=Side(border_style=borders.BORDER_THIN,
                                                                                  color=Color(theme=7))))
    assert ws['A21'].border == Border(
        top=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)),
        bottom=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)),
        left=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)),
        right=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)))

    ws.cell('A23').style = ws.cell('A23').style.copy(border=Border(top=Side(border_style=borders.BORDER_THIN,
                                                                                color=Color(theme=6))),
                                                     fill=PatternFill(fill_type='solid',
                                                               start_color=Color('FFCCCCFF')))
    assert ws['A23'].border == Border(
        top=Side(style=borders.BORDER_THIN, color=Color(theme=6))
    )

    ws.unmerge_cells('A23:B24')

    ws.cell('A25').style = ws.cell('A25').style.copy(alignment=Alignment(wrap_text=False))
    assert ws['A25'].alignment == Alignment(wrap_text=False)

    ws.cell('A26').style = ws.cell('A26').style.copy(alignment=Alignment(shrink_to_fit=False))
    assert ws['A26'].alignment == Alignment(shrink_to_fit=False)

    assert ws.column_dimensions['A'].width == 20.0
