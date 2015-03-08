from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# Python stdlib imports
from io import BytesIO

# compatibility imports
from openpyxl.compat import iteritems, OrderedDict

# package imports
from openpyxl import Workbook
from openpyxl.formatting.rules import ColorScaleRule, CellIsRule, FormulaRule
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
    return OrderedDict(
        [
            ('H1:H10', [{'priority': 23, }]),
            ('Q1:Q10', [{'priority': 14, }]),
            ('G1:G10', [{'priority': 24, }]),
            ('F1:F10', [{'priority': 25, }]),
            ('O1:O10', [{'priority': 16, }]),
            ('K1:K10', []),
            ('T1:T10', [{'priority': 11, }]),
            ('X1:X10', [{'priority': 7, }]),
            ('R1:R10', [{'priority': 13}]),
            ('C1:C10', [{'priority': 28, }]),
            ('J1:J10', [{'priority': 21, }]),
            ('E1:E10', [{'priority': 26, }]),
            ('I1:I10', [{'priority': 22, }]),
            ('Z1:Z10', [{'priority': 5}]),
            ('V1:V10', [{'priority': 9}]),
            ('AC1:AC10', [{'priority': 2, }]),
            ('L1:L10', []),
            ('N1:N10', [{'priority': 17, }]),
            ('AA1:AA10', [{'priority': 4, }]),
            ('M1:M10', []),
            ('Y1:Y10', [{'priority': 6, }]),
            ('B1:B10', [{'priority': 29, }]),
            ('P1:P10', [{'priority': 15, }]),
            ('W1:W10', [{'priority': 8, }]),
            ('AB1:AB10', [{'priority': 3, }]),
            ('A1:A1048576', [{'priority': 29, }]),
            ('S1:S10', [{'priority': 12, }]),
            ('D1:D10', [{'priority': 27, }])
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


def test_update(rules):
    from .. import unpack_rules, ConditionalFormatting
    cf = ConditionalFormatting()
    cf.update(rules)
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


class TestRule:

    def test_ctor(self, FormatRule):
        r = FormatRule()
        assert r == {}

    @pytest.mark.parametrize("key, value",
                             [('aboveAverage', 1),
                              ('bottom', 0),
                              ('dxfId', True),
                              ('equalAverage', False),
                              ('operator', ""),
                              ('percent', 0),
                              ('priority', 1),
                              ('rank', 4),
                              ('stdDev', 2),
                              ('stopIfTrue', False),
                              ('text', "Once upon a time"),
                             ])
    def test_setitem(self, FormatRule, key, value):
        r1 = FormatRule()
        r2 = FormatRule()
        r1[key] = value
        setattr(r2, key, value)
        assert r1 == r2

    def test_getitem(self, FormatRule):
        r = FormatRule()
        r.aboveAverage = 1
        assert r.aboveAverage == r['aboveAverage']

    def test_invalid_key(self, FormatRule):
        r = FormatRule()
        with pytest.raises(KeyError):
            r['randomkey'] = 1
        with pytest.raises(KeyError):
            r['randomkey']

    def test_update_from_dict(self, FormatRule):
        r = FormatRule()
        d = {'aboveAverage':1}
        r.update(d)

    def test_len(self, FormatRule):
        r = FormatRule()
        assert len(r) == 0
        r.aboveAverage = 1
        assert len(r) == 1

    def test_keys(self, FormatRule):
        r = FormatRule()
        assert r.keys() == []
        r['operator'] = True
        assert r.keys() == ['operator']

    def test_values(self, FormatRule):
        r = FormatRule()
        assert r.values() == []
        r['rank'] = 1
        assert r.values() == [1]

    def test_items(self, FormatRule):
        r = FormatRule()
        assert r.items() == []
        r['stopIfTrue'] = False
        assert r.items() == [('stopIfTrue', False)]


class DummyWorkbook():

    def __init__(self):
        self._differential_styles = []
        self.shared_styles = IndexedList()
        self.worksheets = []


class TestConditionalFormatting(object):


    def setup(self):
        self.workbook = DummyWorkbook()

    def test_conditional_formatting_customRule(self):
        class DummyWorksheet():
            conditional_formatting = ConditionalFormatting()

        worksheet = DummyWorksheet()
        worksheet.conditional_formatting.add('C1:C10', {'type': 'expression', 'formula': ['ISBLANK(C1)'],
                                                        'stopIfTrue': '1', 'dxf': {}})
        worksheet.conditional_formatting._save_styles(self.workbook)
        cfs = write_conditional_formatting(worksheet)
        xml = b""
        for cf in cfs:
            xml += tostring(cf)

        diff = compare_xml(xml, """
        <conditionalFormatting sqref="C1:C10">
          <cfRule dxfId="0" type="expression" stopIfTrue="1" priority="1">
            <formula>ISBLANK(C1)</formula>
          </cfRule>
        </conditionalFormatting>
        """)
        assert diff is None, diff

    def test_conditional_formatting_setDxfStyle(self):
        cf = ConditionalFormatting()
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
        cf._save_styles(self.workbook)
        assert len(self.workbook._differential_styles) == 2
        ft1, ft2 = self.workbook._differential_styles
        assert ft1.font == font
        assert ft1.border == border
        assert ft1.fill == fill
        assert ft2.fill == fill

    def test_conditional_formatting_update(self):
        class WS():
            conditional_formatting = ConditionalFormatting()
        worksheet = WS()
        rules = {'A1:A4': [{'type': 'colorScale', 'priority': 13,
                            'colorScale': {'cfvo': [{'type': 'min'}, {'type': 'max'}], 'color':
                                           [Color('FFFF7128'), Color('FFFFEF9C')]}}]}
        worksheet.conditional_formatting.update(rules)

        cfs = write_conditional_formatting(worksheet)
        xml = b""
        for cf in cfs:
            xml += tostring(cf)

        diff = compare_xml(xml, """
        <conditionalFormatting sqref="A1:A4">
          <cfRule type="colorScale" priority="1">
            <colorScale>
              <cfvo type="min" />
              <cfvo type="max" />
              <color rgb="FFFF7128" />
              <color rgb="FFFFEF9C" />
            </colorScale>
          </cfRule>
        </conditionalFormatting>
        """)
        assert diff is None, diff

    def test_conditional_font(self):
        """Test to verify font style written correctly."""
        class WS():
            conditional_formatting = ConditionalFormatting()
        worksheet = WS()

        # Create cf rule
        redFill = PatternFill(start_color=Color('FFEE1111'),
                       end_color=Color('FFEE1111'),
                       patternType=fills.FILL_SOLID)
        whiteFont = Font(color=Color("FFFFFFFF"))
        worksheet.conditional_formatting.add('A1:A3',
                                             CellIsRule(operator='equal', formula=['"Fail"'], stopIfTrue=False,
                                                        font=whiteFont, fill=redFill))
        worksheet.conditional_formatting._save_styles(self.workbook)

        # First, verify conditional formatting xml
        cfs = write_conditional_formatting(worksheet)
        xml = b""
        for cf in cfs:
            xml += tostring(cf)

        diff = compare_xml(xml, """
        <conditionalFormatting sqref="A1:A3">
          <cfRule dxfId="0" operator="equal" priority="1" type="cellIs">
            <formula>"Fail"</formula>
          </cfRule>
        </conditionalFormatting>
        """)
        assert diff is None, diff


class TestColorScaleRule(object):

    def test_two_colors(self):
        cfRule = ColorScaleRule(start_type='min', start_value=None, start_color='FFAA0000',
                                end_type='max', end_value=None, end_color='FF00AA00')
        assert dict(cfRule) == {'colorScale':
                                {'cfvo': [{'type': 'min'}, {'type': 'max'}],
                                'color': ['FFAA0000', 'FF00AA00']},
                                'type': 'colorScale'
                                }


    def test_three_colors(self):
        cfRule = ColorScaleRule(start_type='percentile', start_value=10, start_color='FFAA0000',
                                mid_type='percentile', mid_value=50, mid_color='FF0000AA',
                                end_type='percentile', end_value=90, end_color='FF00AA00')
        assert dict(cfRule) == {'colorScale':
                                {'cfvo': [
                                    {'type': 'percentile', 'val':'10'},
                                    {'type': 'percentile', 'val':'50'},
                                    {'type': 'percentile', 'val':'90'}],
                                'color': ['FFAA0000', 'FF0000AA', 'FF00AA00']},
                                'type': 'colorScale'
                                }


@pytest.mark.parametrize("value, expansion",
                         [
                             ('<=', 'lessThanOrEqual'),
                             ('>', 'greaterThan'),
                             ('!=', 'notEqual'),
                             ('=', 'equal'),
                             ('>=', 'greaterThanOrEqual'),
                             ('==', 'equal'),
                             ('<', 'lessThan'),
                         ]
                         )
def test_operator_expansion(value, expansion):
    cf1 = CellIsRule()
    cf2 = CellIsRule()
    cf1.operator = value
    cf2.operator = expansion
    assert cf1.operator == expansion
    assert cf2.operator == expansion


def test_formula_rule():

    cf1 = FormulaRule(formula=['ISBLANK(C1)'], stopIfTrue=True)
    assert cf1.rule == {'dxf':
                        {'border': None, 'fill': None, 'font': None},
                        'formula': ['ISBLANK(C1)'],
                        'stopIfTrue': '1',
                        'type': 'expression'
                        }


def test_conditional_formatting_read(datadir):
    datadir.chdir()
    reference_file = 'conditional-formatting.xlsx'
    wb = load_workbook(reference_file)
    ws = wb.get_active_sheet()

    # First test the conditional formatting rules read
    assert ws.conditional_formatting.cf_rules['A1:A1048576'] == [
        {'priority': 27,
         'type': 'colorScale',
         'colorScale': {'color': [Color('FFFF7128'), Color('FFFFEF9C')],
                        'cfvo': [{'type': 'min'}, {'type': 'max'}]}
         }
    ]
    assert ws.conditional_formatting.cf_rules['B1:B10'] == [
        {'priority': 26,
         'type': 'colorScale',
         'colorScale': {'color': [Color(theme=6), Color(theme=4)],
                        'cfvo': [{'type': 'num', 'val': '3'},
                                 {'type': 'num', 'val': '7'}]}
         }
    ]
    assert ws.conditional_formatting.cf_rules['C1:C10'] == [
        {'priority': 25,
         'type': 'colorScale',
         'colorScale': {'color': [Color('FFFF7128'), Color('FFFFEF9C')],
                        'cfvo': [{'type': 'percent', 'val': '10'},
                                 {'type': 'percent', 'val': '90'}]}
         }]
    assert ws.conditional_formatting.cf_rules['D1:D10'] == [
        {'priority': 24,
         'type': 'colorScale',
         'colorScale': {'color': [Color(theme=6), Color(theme=5)],
                        'cfvo': [{'type': 'formula', 'val': '2'},
                                 {'type': 'formula', 'val': '4'}]}
         }
    ]
    assert ws.conditional_formatting.cf_rules['E1:E10'] == [
        {'priority': 23,
         'type': 'colorScale',
         'colorScale': {'color': [Color('FFFF7128'), Color('FFFFEF9C')],
                        'cfvo': [{'type': 'percentile', 'val': '10'},
                                 {'type': 'percentile', 'val': '90'}]}
         }
    ]
    assert ws.conditional_formatting.cf_rules['F1:F10'] == [
        {'priority': 22,
         'type': 'colorScale',
         'colorScale': {'color': [Color('FFFF7128'), Color('FFFFEB84'), Color('FF63BE7B')],
                        'cfvo': [{'type': 'min'},
                                 {'type': 'percentile', 'val': '50'},
                                 {'type': 'max'}]}
         }
    ]
    assert ws.conditional_formatting.cf_rules['G1:G10'] == [
        {'priority': 21,
         'type': 'colorScale',
         'colorScale': {'color': [Color(theme=4), Color('FFFFEB84'), Color(theme=5)],
                        'cfvo': [{'type': 'num', 'val': '0'},
                                 {'type': 'percentile', 'val': '50'},
                                 {'type': 'num', 'val': '10'}]}
         }]
    assert ws.conditional_formatting.cf_rules['H1:H10'] == [
        {'priority': 20,
         'type': 'colorScale',
         'colorScale': {'color': [Color('FFFF7128'), Color('FFFFEB84'), Color('FF63BE7B')],
                        'cfvo': [{'type': 'percent', 'val': '0'},
                                 {'type': 'percent', 'val': '50'},
                                 {'type': 'percent', 'val': '100'}]}
         }]
    assert ws.conditional_formatting.cf_rules['I1:I10'] == [
        {'priority': 19,
         'type': 'colorScale',
         'colorScale': {'color': [Color('FF0000FF'), Color('FFFF6600'), Color('FF008000')],
                        'cfvo': [{'type': 'formula', 'val': '2'},
                                 {'type': 'formula', 'val': '7'},
                                 {'type': 'formula', 'val': '9'}]}
         }]
    assert ws.conditional_formatting.cf_rules['J1:J10'] == [
        {'priority': 18,
         'type': 'colorScale',
         'colorScale': {'color': [Color('FFFF7128'), Color('FFFFEB84'), Color('FF63BE7B')],
                        'cfvo': [{'type': 'percentile', 'val': '10'},
                                 {'type': 'percentile', 'val': '50'},
                                 {'type': 'percentile', 'val': '90'}]}
         }]
    assert ws.conditional_formatting.cf_rules['K1:K10'] == []  # K - M are dataBar conditional formatting, which are not
    assert ws.conditional_formatting.cf_rules['L1:L10'] == []  # handled at the moment, and should not load, but also
    assert ws.conditional_formatting.cf_rules['M1:M10'] == []  # should not interfere with the loading / saving of the file.
    assert ws.conditional_formatting.cf_rules['N1:N10'] == [
        {'priority': 17,
         'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'},
                              {'type': 'percent', 'val': '33'},
                              {'type': 'percent', 'val': '67'}]},
         'type': 'iconSet'}
    ]
    assert ws.conditional_formatting.cf_rules['O1:O10'] == [
        {'priority': 16,
         'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'},
                              {'type': 'num', 'val': '2'},
                              {'type': 'num', 'val': '4'},
                              {'type': 'num', 'val': '6'}],
                     'showValue': '0',
                     'iconSet': '4ArrowsGray', 'reverse': '1'},
         'type': 'iconSet'}
    ]
    assert ws.conditional_formatting.cf_rules['P1:P10'] == [
        {'priority': 15,
         'iconSet': {'cfvo': [{'type': 'percent', 'val': '0'},
                              {'type': 'percentile', 'val': '20'},
                              {'type': 'percentile', 'val': '40'},
                              {'type': 'percentile', 'val': '60'},
                              {'type': 'percentile', 'val': '80'}],
                     'iconSet': '5Rating'},
         'type': 'iconSet'}
    ]
    assert ws.conditional_formatting.cf_rules['Q1:Q10'] == [
        {'text': '3',
         'priority': 14,
         'dxfId': '27',
         'operator': 'containsText',
         'formula': ['NOT(ISERROR(SEARCH("3",Q1)))'],
         'type': 'containsText'}
                                                            ]
    assert ws.conditional_formatting.cf_rules['R1:R10'] == [
        {'operator': 'between',
         'dxfId': '26',
         'type': 'cellIs',
         'formula': ['2', '7'],
         'priority': 13}
    ]
    assert ws.conditional_formatting.cf_rules['S1:S10'] == [
        {'priority': 12,
         'dxfId': '25',
         'percent': '1',
         'type': 'top10',
         'rank': '10'}
    ]
    assert ws.conditional_formatting.cf_rules['T1:T10'] == [
        {'priority': 11,
         'dxfId': '24',
         'type': 'top10',
         'rank': '4',
         'bottom': '1'}
    ]
    assert ws.conditional_formatting.cf_rules['U1:U10'] == [
        {'priority': 10,
         'dxfId': '23',
         'type': 'aboveAverage'}
    ]
    assert ws.conditional_formatting.cf_rules['V1:V10'] == [
        {'aboveAverage': '0',
         'dxfId': '22',
         'type': 'aboveAverage',
         'priority': 9}
    ]
    assert ws.conditional_formatting.cf_rules['W1:W10'] == [
        {'priority': 8,
         'dxfId': '21',
         'type': 'aboveAverage',
         'equalAverage': '1'}
    ]
    assert ws.conditional_formatting.cf_rules['X1:X10'] == [
        {'aboveAverage': '0',
         'dxfId': '20',
         'priority': 7,
         'type': 'aboveAverage',
         'equalAverage': '1'}
    ]
    assert ws.conditional_formatting.cf_rules['Y1:Y10'] == [
        {'priority': 6,
         'dxfId': '19',
         'type': 'aboveAverage',
         'stdDev': '1'}
    ]
    assert ws.conditional_formatting.cf_rules['Z1:Z10'] == [
        {'aboveAverage': '0',
         'dxfId': '18',
         'type': 'aboveAverage',
         'stdDev': '1', 'priority': 5}
    ]
    assert ws.conditional_formatting.cf_rules['AA1:AA10'] == [
        {'priority': 4,
         'dxfId': '17',
         'type': 'aboveAverage',
         'stdDev': '2'}
    ]
    assert ws.conditional_formatting.cf_rules['AB1:AB10'] == [
        {'priority': 3,
         'dxfId': '16',
         'type': 'duplicateValues'}
    ]
    assert ws.conditional_formatting.cf_rules['AC1:AC10'] == [
        {'priority': 2,
         'dxfId': '15',
         'type': 'uniqueValues'}
    ]
    assert ws.conditional_formatting.cf_rules['AD1:AD10'] == [
        {'priority': 1,
         'dxfId': '14',
         'type': 'expression',
         'formula': ['AD1>3']}
    ]
