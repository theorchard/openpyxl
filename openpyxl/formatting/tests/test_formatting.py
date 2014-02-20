from collections import OrderedDict
import pytest

from openpyxl.compat import iteritems

rules = OrderedDict([
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
    ('A1:A1048576', [{'priority': 30, }]),
    ('S1:S10', [{'priority': 12, }]),
    ('D1:D10', [{'priority': 27, }])
]
                    )


def test_unpack_rules():
    from openpyxl.formatting import unpack_rules
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
        ('A1:A1048576',0 ,30),
        ('S1:S10', 0, 12),
        ('D1:D10', 0, 27),
    ]


def test_update():
    from openpyxl.formatting import unpack_rules
    from openpyxl.formatting import ConditionalFormatting
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


def test_fix_priorities():
    from openpyxl.formatting import unpack_rules
    from openpyxl.formatting import ConditionalFormatting
    cf = ConditionalFormatting()
    cf.cf_rules = rules
    cf._fix_priorities()
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

