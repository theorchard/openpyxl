from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# Python stdlib imports
from io import BytesIO

# compatibility imports
from openpyxl.compat import iteritems, OrderedDict

# package imports
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, Rule
from openpyxl.reader.excel import load_workbook
from openpyxl.reader.style import SharedStylesParser
from openpyxl.xml.constants import ARC_STYLE
from openpyxl.xml.functions import tostring
from openpyxl.writer.worksheet import write_conditional_formatting
from openpyxl.writer.styles import StyleWriter
from openpyxl.styles import Color, PatternFill, Font, Border, Side
from openpyxl.styles import borders, fills, colors
from openpyxl.formatting import ConditionalFormatting

# test imports
import pytest
from zipfile import ZIP_DEFLATED, ZipFile
from openpyxl.tests.helper import compare_xml
from openpyxl.utils.indexed_list import IndexedList


@pytest.fixture
def rules():

    class DummyRule:

        def __init__(self, priority):
            self.priority = priority

    return OrderedDict(
        [
            ('H1:H10', [DummyRule(23)]),
            ('Q1:Q10', [DummyRule(14)]),
            ('G1:G10', [DummyRule(24)]),
            ('F1:F10', [DummyRule(25)]),
            ('O1:O10', [DummyRule(16)]),
            ('K1:K10', []),
            ('T1:T10', [DummyRule(11)]),
            ('X1:X10', [DummyRule(7)]),
            ('R1:R10', [DummyRule(13)]),
            ('C1:C10', [DummyRule(28)]),
            ('J1:J10', [DummyRule(21)]),
            ('E1:E10', [DummyRule(26)]),
            ('I1:I10', [DummyRule(22)]),
            ('Z1:Z10', [DummyRule(5)]),
            ('V1:V10', [DummyRule(9)]),
            ('AC1:AC10', [DummyRule(2)]),
            ('L1:L10', []),
            ('N1:N10', [DummyRule(17)]),
            ('AA1:AA10', [DummyRule(4)]),
            ('M1:M10', []),
            ('Y1:Y10', [DummyRule(6)]),
            ('B1:B10', [DummyRule(29)]),
            ('P1:P10', [DummyRule(15)]),
            ('W1:W10', [DummyRule(8)]),
            ('AB1:AB10', [DummyRule(3)]),
            ('A1:A1048576', [DummyRule(29)]),
            ('S1:S10', [DummyRule(12)]),
            ('D1:D10', [DummyRule(27)])
        ]
    )


def test_unpack_rules(rules):
    from .. import unpack_rules
    assert list(unpack_rules(rules)) == [
        ('H1:H10', 0, 23),
        ('Q1:Q10', 0, 14),
        ('G1:G10', 0, 24),
        ('F1:F10', 0, 25),
        ('O1:O10', 0, 16),
        ('T1:T10', 0, 11),
        ('X1:X10', 0, 7),
        ('R1:R10', 0, 13),
        ('C1:C10', 0, 28),
        ('J1:J10', 0, 21),
        ('E1:E10', 0, 26),
        ('I1:I10', 0, 22),
        ('Z1:Z10', 0, 5),
        ('V1:V10', 0, 9),
        ('AC1:AC10', 0, 2),
        ('N1:N10', 0, 17),
        ('AA1:AA10', 0, 4),
        ('Y1:Y10', 0, 6),
        ('B1:B10', 0, 29),
        ('P1:P10', 0, 15),
        ('W1:W10', 0, 8),
        ('AB1:AB10', 0, 3),
        ('A1:A1048576',0 ,29),
        ('S1:S10', 0, 12),
        ('D1:D10', 0, 27),
    ]


def test_fix_priorities(rules):
    from .. import unpack_rules, ConditionalFormatting
    cf = ConditionalFormatting()
    cf.cf_rules = rules
    cf._fix_priorities()
    assert cf.max_priority == 25
    assert list(unpack_rules(cf.cf_rules)) == [
        ('H1:H10', 0, 18),
        ('Q1:Q10', 0, 12),
        ('G1:G10', 0, 19),
        ('F1:F10', 0, 20),
        ('O1:O10', 0, 14),
        ('T1:T10', 0, 9),
        ('X1:X10', 0, 6),
        ('R1:R10', 0, 11),
        ('C1:C10', 0, 23),
        ('J1:J10', 0, 16),
        ('E1:E10', 0, 21),
        ('I1:I10', 0, 17),
        ('Z1:Z10', 0, 4),
        ('V1:V10', 0, 8),
        ('AC1:AC10', 0, 1),
        ('N1:N10', 0, 15),
        ('AA1:AA10', 0, 3),
        ('Y1:Y10', 0, 5),
        ('B1:B10', 0, 24),
        ('P1:P10', 0, 13),
        ('W1:W10', 0, 7),
        ('AB1:AB10', 0, 2),
        ('A1:A1048576', 0, 25),
        ('S1:S10', 0, 10),
        ('D1:D10', 0, 22),
    ]


class DummyWorkbook():

    def __init__(self):
        self.differential_styles = []
        self.shared_styles = IndexedList()
        self.worksheets = []

class DummyWorksheet():

    def __init__(self):
        self.conditional_formatting = ConditionalFormatting()
        self.parent = DummyWorkbook()


class TestConditionalFormatting(object):


    def setup(self):
        self.ws = DummyWorksheet()

    def test_conditional_formatting_customRule(self):

        worksheet = self.ws
        worksheet.conditional_formatting.add('C1:C10',
                                             Rule(**{'type': 'expression', 'formula': ['ISBLANK(C1)'],
                                                     'stopIfTrue': '1',}
                                                  )
                                             )
        cfs = write_conditional_formatting(worksheet)
        xml = b""
        for cf in cfs:
            xml += tostring(cf)

        diff = compare_xml(xml, """
        <conditionalFormatting sqref="C1:C10">
          <cfRule type="expression" stopIfTrue="1" priority="1">
            <formula>ISBLANK(C1)</formula>
          </cfRule>
        </conditionalFormatting>
        """)
        assert diff is None, diff

    def test_conditional_formatting_setDxfStyle(self):
        ws = self.ws
        cf = ConditionalFormatting()
        ws.conditional_formatting = cf

        fill = PatternFill(start_color=Color('FFEE1111'),
                    end_color=Color('FFEE1111'),
                    patternType=fills.FILL_SOLID)
        font = Font(name='Arial', size=12, bold=True,
                    underline=Font.UNDERLINE_SINGLE)
        border = Border(top=Side(border_style=borders.BORDER_THIN,
                                 color=Color(colors.DARKYELLOW)),
                        bottom=Side(border_style=borders.BORDER_THIN,
                                    color=Color(colors.BLACK)))
        cf.add('C1:C10', FormulaRule(formula=['ISBLANK(C1)'], font=font, border=border, fill=fill))
        cf.add('D1:D10', FormulaRule(formula=['ISBLANK(D1)'], fill=fill))
        from openpyxl.writer.worksheet import write_conditional_formatting
        for _ in write_conditional_formatting(ws):
            pass # exhaust generator

        wb = ws.parent
        assert len(wb.differential_styles) == 2
        ft1, ft2 = wb.differential_styles
        assert ft1.font == font
        assert ft1.border == border
        assert ft1.fill == fill
        assert ft2.fill == fill


    def test_conditional_font(self):
        """Test to verify font style written correctly."""

        ws = self.ws
        cf = ConditionalFormatting()
        ws.conditional_formatting = cf

        # Create cf rule
        redFill = PatternFill(start_color=Color('FFEE1111'),
                       end_color=Color('FFEE1111'),
                       patternType=fills.FILL_SOLID)
        whiteFont = Font(color=Color("FFFFFFFF"))
        ws.conditional_formatting.add('A1:A3',
                                      CellIsRule(operator='equal', formula=['"Fail"'], stopIfTrue=False,
                                                 font=whiteFont, fill=redFill))

        from openpyxl.writer.worksheet import write_conditional_formatting

        # First, verify conditional formatting xml
        cfs = write_conditional_formatting(ws)
        xml = b""
        for cf in cfs:
            xml += tostring(cf)

        diff = compare_xml(xml, """
        <conditionalFormatting sqref="A1:A3">
          <cfRule dxfId="0" operator="equal" priority="1" type="cellIs" stopIfTrue="0">
            <formula>"Fail"</formula>
          </cfRule>
        </conditionalFormatting>
        """)
        assert diff is None, diff


def test_conditional_formatting_read(datadir):
    datadir.chdir()
    reference_file = 'conditional-formatting.xlsx'
    wb = load_workbook(reference_file)
    ws = wb.active
    rules = ws.conditional_formatting.cf_rules
    assert len(rules) == 30

    # First test the conditional formatting rules read
    rule = rules['A1:A1048576'][0]
    assert dict(rule) == {'priority':'30', 'type': 'colorScale', }

    rule = rules['B1:B10'][0]
    assert dict(rule) == {'priority': '29', 'type': 'colorScale'}

    rule = rules['C1:C10'][0]
    assert dict(rule) == {'priority': '28', 'type': 'colorScale'}

    rule = rules['D1:D10'][0]
    assert dict(rule) == {'priority': '27', 'type': 'colorScale', }

    rule = rules['E1:E10'][0]
    assert dict(rule) == {'priority': '26', 'type': 'colorScale', }

    rule = rules['F1:F10'][0]
    assert dict(rule) == {'priority': '25', 'type': 'colorScale', }

    rule = rules['G1:G10'][0]
    assert dict(rule) == {'priority': '24', 'type': 'colorScale', }

    rule = rules['H1:H10'][0]
    assert dict(rule) == {'priority': '23', 'type': 'colorScale', }

    rule = rules['I1:I10'][0]
    assert dict(rule) == {'priority': '22', 'type': 'colorScale', }

    rule = rules['J1:J10'][0]
    assert dict(rule) == {'priority': '21', 'type': 'colorScale', }

    rule = rules['K1:K10'][0]
    assert dict(rule) ==  {'priority': '20', 'type': 'dataBar'}

    rule = rules['L1:L10'][0]
    assert dict(rule) ==  {'priority': '19', 'type': 'dataBar'}

    rule = rules['M1:M10'][0]
    assert dict(rule) ==  {'priority': '18', 'type': 'dataBar'}

    rule = rules['N1:N10'][0]
    assert dict(rule) == {'priority': '17', 'type': 'iconSet'}

    rule = rules['O1:O10'][0]
    assert dict(rule) == {'priority': '16', 'type': 'iconSet'}

    rule = rules['P1:P10'][0]
    assert dict(rule) == {'priority': '15', 'type': 'iconSet'}

    # need to check dxf
    rule = rules['Q1:Q10'][0]
    assert dict(rule) == {'text': '3', 'priority': '14', 'dxfId': '27',
                          'operator': 'containsText', 'type': 'containsText'}

    rule = rules['R1:R10'][0]
    assert dict(rule) == {'operator': 'between', 'dxfId': '26', 'type':
                          'cellIs', 'priority': '13'}

    rule = rules['S1:S10'][0]
    assert dict(rule) == {'priority': '12', 'dxfId': '25', 'percent': '1',
                          'type': 'top10', 'rank': '10'}

    rule = rules['T1:T10'][0]
    assert dict(rule) == {'priority': '11', 'dxfId': '24', 'type': 'top10',
                          'rank': '4', 'bottom': '1'}

    rule = rules['U1:U10'][0]
    assert dict(rule) == {'priority': '10', 'dxfId': '23', 'type':
                          'aboveAverage'}

    rule = rules['V1:V10'][0]
    assert dict(rule) == {'aboveAverage': '0', 'dxfId': '22', 'type':
                          'aboveAverage', 'priority': '9'}

    rule = rules['W1:W10'][0]
    assert dict(rule) == {'priority': '8', 'dxfId': '21', 'type':
                          'aboveAverage', 'equalAverage': '1'}

    rule = rules['X1:X10'][0]
    assert dict(rule) == {'aboveAverage': '0', 'dxfId': '20', 'priority': '7',
                           'type': 'aboveAverage', 'equalAverage': '1'}

    rule = rules['Y1:Y10'][0]
    assert dict(rule) == {'priority': '6', 'dxfId': '19', 'type':
                          'aboveAverage', 'stdDev': '1'}

    rule = rules['Z1:Z10'][0]
    assert dict(rule)== {'aboveAverage': '0', 'dxfId': '18', 'type':
                         'aboveAverage', 'stdDev': '1', 'priority': '5'}

    rule = rules['AA1:AA10'][0]
    assert dict(rule) == {'priority': '4', 'dxfId': '17', 'type':
                          'aboveAverage', 'stdDev': '2'}

    rule = rules['AB1:AB10'][0]
    assert dict(rule) == {'priority': '3', 'dxfId': '16', 'type':
                          'duplicateValues'}

    rule = rules['AC1:AC10'][0]
    assert dict(rule) == {'priority': '2', 'dxfId': '15', 'type':
                          'uniqueValues'}

    rule = rules['AD1:AD10'][0]
    assert dict(rule) == {'priority': '1', 'dxfId': '14', 'type': 'expression',}
