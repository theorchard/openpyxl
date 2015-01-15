from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from io import BytesIO
from zipfile import ZipFile

# package imports
from openpyxl.compat import safe_string, OrderedDict
from openpyxl.reader.excel import load_workbook

from openpyxl.styles import (
    borders,
    numbers,
    Color,
    Font,
    PatternFill,
    GradientFill,
    Border,
    Side,
    Alignment
)
from openpyxl.xml.functions import Element


@pytest.fixture
def StyleReader():
    from ..style import SharedStylesParser
    return SharedStylesParser


@pytest.mark.parametrize("value, expected",
                         [
                         ('f', False),
                         ('0', False),
                         ('false', False),
                         ('1', True),
                         ('t', True),
                         ('true', True),
                         ('anyvalue', True),
                         ])
def test_bool_attrib(value, expected):
    from .. style import bool_attrib
    el = Element("root", value=value)
    assert bool_attrib(el, "value") is expected


def test_unprotected_cell(StyleReader, datadir):
    datadir.chdir()
    with open ("worksheet_unprotected_style.xml") as src:
        reader = StyleReader(src.read())
    from openpyxl.styles import Font
    reader.font_list = [Font(), Font(), Font(), Font(), Font()]
    reader.parse_cell_styles()
    assert len(reader.shared_styles) == 3
    # default is cells are locked
    style = reader.shared_styles[0]
    assert style.protection.locked is True

    style = reader.shared_styles[2]
    assert style.protection.locked is False


def test_read_cell_style(datadir, StyleReader):
    datadir.chdir()
    with open("empty-workbook-styles.xml") as content:
        reader = StyleReader(content.read())
    reader.parse()
    style_properties = reader.shared_styles
    assert len(style_properties) == 2


def test_read_xf_no_number_format(datadir, StyleReader):
    datadir.chdir()
    with open("no_number_format.xml") as src:
        reader = StyleReader(src.read())

    from openpyxl.styles import Font
    reader.font_list = [Font(), Font()]
    reader.parse_cell_styles()

    styles = reader.shared_styles
    assert len(styles) == 3
    assert styles[0].number_format == 'General'
    assert styles[1].number_format == 'General'
    assert styles[2].number_format == 'mm-dd-yy'



def test_read_simple_style_mappings(datadir, StyleReader):
    datadir.chdir()
    with open("simple-styles.xml") as content:
        reader = StyleReader(content.read())
    reader.parse()
    style_properties = reader.shared_styles
    assert len(style_properties) == 4
    assert numbers.BUILTIN_FORMATS[9] == style_properties[1].number_format
    assert 'yyyy-mm-dd' == style_properties[2].number_format


def test_read_complex_style_mappings(datadir, StyleReader):
    datadir.chdir()
    with open("complex-styles.xml") as content:
        reader = StyleReader(content.read())
    reader.parse()
    style_properties = reader.shared_styles
    assert len(style_properties) == 29
    assert style_properties[-1].font.bold is False


def test_read_complex_fonts(datadir, StyleReader):
    from openpyxl.styles import Font
    datadir.chdir()
    with open("complex-styles.xml") as content:
        reader = StyleReader(content.read())
    fonts = list(reader.parse_fonts())
    assert len(fonts) == 8
    assert fonts[7] == Font(size=12, color=Color(theme=9), name="Calibri", scheme="minor")


def test_read_complex_fills(datadir, StyleReader):
    datadir.chdir()
    with open("complex-styles.xml") as content:
        reader = StyleReader(content.read())
    fills = list(reader.parse_fills())
    assert len(fills) == 6


def test_read_complex_borders(datadir, StyleReader):
    datadir.chdir()
    with open("complex-styles.xml") as content:
        reader = StyleReader(content.read())
    borders = list(reader.parse_borders())
    assert len(borders) == 7


def test_read_complex_style(datadir):
    datadir.chdir()
    wb = load_workbook("complex-styles.xlsx")
    ws = wb.get_active_sheet()
    assert ws.column_dimensions['A'].width == 31.1640625

    assert ws.column_dimensions['I'].style.font == Font(sz=12.0, color='FF3300FF', scheme='minor')
    assert ws.column_dimensions['I'].style.fill == PatternFill(patternType='solid', fgColor='FF006600', bgColor=Color(indexed=64))

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
    assert ws['A15'].number_format == '0.00'
    assert ws['A16'].number_format == 'mm-dd-yy'
    assert ws['A17'].number_format == '0.00%'

    assert 'A18:B18' in ws._merged_cells

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
    i_style = ws.column_dimensions['I'].style
    ws.column_dimensions['I'].style = i_style.copy(fill=PatternFill(fill_type='solid',
                                             start_color=Color('FF442200')),
                                   font=Font(color=Color('FF002244')))

    assert ws.column_dimensions['I'].style.fill.start_color.value == 'FF442200'
    assert ws.column_dimensions['I'].style.font.color.value == 'FF002244'

    ws['A2'].font = Font(name='Times New Roman',
                         size=12,
                         bold=True,
                         italic=True,
                         )
    assert ws['A2'].font == Font(name='Times New Roman', size=12, bold=True,
                                 italic=True)

    ws['A3'].font = Font(name='Times New Roman',
                         size=14,
                         bold=False,
                         italic=True,
                         )
    assert ws['A3'].font == Font(name='Times New Roman', size=14,
                                 bold=False, italic=True,)


    ws['A4'].font = Font(name='Times New Roman',
                         size=16,
                         bold=True,
                         italic=False,
                         )
    assert ws['A4'].font == Font(name='Times New Roman', size=16, bold=True,
                                 italic=False,)

    ws['A5'].font = Font(color=Color('FF66FF66'))
    assert ws['A5'].font == Font(color='FF66FF66')


    ws['A6'].font =Font(color=Color(theme='1'))
    assert ws['A6'].font == Font(color=Color(theme='1'))

    ws['A7'].fill = PatternFill(fill_type='solid',
                                start_color=Color('FF330066'))
    assert ws['A7'].fill == PatternFill(fill_type='solid',
                                        start_color=Color('FF330066'))

    ws['A8'].fill = PatternFill(fill_type='solid',
                                start_color=Color(theme='2'))
    assert ws['A8'].fill == PatternFill(fill_type='solid',
                                        start_color=Color(theme='2'))

    ws['A9'].alignment = Alignment(horizontal='center')
    assert ws['A9'].alignment == Alignment(horizontal='center')

    ws['A10'].alignment = Alignment(horizontal='left')
    assert ws['A10'].alignment == Alignment(horizontal='left')

    ws['A11'].alignment = Alignment(horizontal='right')
    assert ws['A11'].alignment == Alignment(horizontal='right')

    ws['A12'].alignment = Alignment(vertical='bottom')
    assert ws['A12'].alignment == Alignment(vertical='bottom')

    ws['A13'].alignment= Alignment(vertical='top')
    assert ws['A13'].alignment == Alignment(vertical='top')

    ws['A14'].alignment = Alignment(vertical='center')
    assert ws['A14'].alignment == Alignment(vertical='center')

    ws['A15'].number_format='0.00%'
    assert ws['A15'].number_format == '0.00%'

    ws['A16'].number_format ='0.00'
    assert ws['A16'].number_format == '0.00'

    ws['A17'].number_format='mm-dd-yy'
    assert ws['A17'].number_format == 'mm-dd-yy'

    ws.unmerge_cells('A18:B18')

    ws['A19'].border = Border(top=Side(border_style=borders.BORDER_THIN,
                                       color=Color('FF006600')),
                              bottom=Side(border_style=borders.BORDER_THIN,
                                          color=Color('FF006600')),
                              left=Side(border_style=borders.BORDER_THIN,
                                        color=Color('FF006600')),
                              right=Side(border_style=borders.BORDER_THIN,
                                         color=Color('FF006600')))
    assert ws['A19'].border == Border(
        top=Side(border_style=borders.BORDER_THIN, color='FF006600'),
        bottom=Side(border_style=borders.BORDER_THIN, color='FF006600'),
        left=Side(border_style=borders.BORDER_THIN, color='FF006600'),
        right=Side(border_style=borders.BORDER_THIN, color='FF006600'))

    ws['A21'].border = Border(top=Side(border_style=borders.BORDER_THIN,
                                       color=Color(theme=7)),
                              bottom=Side(border_style=borders.BORDER_THIN,
                                          color=Color(theme=7)),
                              left=Side(border_style=borders.BORDER_THIN,
                                        color=Color(theme=7)),
                              right=Side(border_style=borders.BORDER_THIN,
                                         color=Color(theme=7)))
    assert ws['A21'].border == Border(
        top=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)),
        bottom=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)),
        left=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)),
        right=Side(border_style=borders.BORDER_THIN, color=Color(theme=7)))


    ws['A23'].fill = PatternFill(fill_type='solid',
                               start_color=Color('FFCCCCFF'))
    assert ws['A23'].fill == PatternFill(fill_type='solid',
                               start_color=Color('FFCCCCFF'))

    ws.unmerge_cells('A23:B24')

    ws['A25'].alignment = Alignment(wrap_text=False)
    assert ws['A25'].alignment == Alignment(wrap_text=False)

    ws['A26'].alignment = Alignment(shrink_to_fit=False)
    assert ws['A26'].alignment == Alignment(shrink_to_fit=False)

    assert ws.column_dimensions['A'].width == 20.0


def test_none_values(datadir, StyleReader):
    datadir.chdir()
    with open("none_value_styles.xml") as src:
        reader = StyleReader(src.read())
    fonts = tuple(reader.parse_fonts())
    assert fonts[0].scheme is None
    assert fonts[0].vertAlign is None
    assert fonts[1].u is None


def test_alignment(datadir, StyleReader):
    datadir.chdir()
    with open("alignment_styles.xml") as src:
        reader = StyleReader(src.read())
    reader.parse_cell_styles()
    st1 = reader.shared_styles[2]
    assert st1.alignment.textRotation == 255


def test_style_names(datadir, StyleReader):
    datadir.chdir()
    with open("complex-styles.xml") as src:
        reader = StyleReader(src.read())

    styles = list(reader._parse_style_names())
    assert styles == [
        ('Followed Hyperlink', 2),
        ('Followed Hyperlink', 4),
        ('Followed Hyperlink', 6),
        ('Followed Hyperlink', 8),
        ('Followed Hyperlink', 10),
        ('Hyperlink', 1),
        ('Hyperlink', 3),
        ('Hyperlink', 5),
        ('Hyperlink', 7),
        ('Hyperlink', 9),
        ('Normal', 0),
    ]


def test_named_styles(datadir, StyleReader):
    from openpyxl.styles.named_styles import NamedStyle
    from openpyxl.styles.fonts import DEFAULT_FONT
    from openpyxl.styles.fills import DEFAULT_EMPTY_FILL

    datadir.chdir()
    with open("complex-styles.xml") as src:
        reader = StyleReader(src.read())

    reader.border_list = list(reader.parse_borders())
    reader.fill_list = list(reader.parse_fills())
    reader.font_list = list(reader.parse_fonts())
    reader.parse_cell_styles()
    reader.parse_named_styles()
    assert len(reader.named_styles) == 11
    first_style = reader.named_styles[0]
    assert first_style.name == "Followed Hyperlink"
    assert first_style.font == Font(size=12, color=Color(theme=11), underline="single", scheme="minor")
    assert first_style.fill == DEFAULT_EMPTY_FILL
    assert first_style.border == Border()


def test_no_styles():
    from .. style import read_style_table
    archive = ZipFile(BytesIO(), "a")
    assert read_style_table(archive) is None
