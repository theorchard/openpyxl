# Copyright (c) 2010-2014 openpyxl

import pytest


@pytest.fixture
def NumberFormat():
    """NumberFormat Class"""
    from openpyxl.styles import NumberFormat
    return NumberFormat


def test_format_comparisions(NumberFormat):
    format1 = NumberFormat('m/d/yyyy')
    format2 = NumberFormat('m/d/yyyy')
    format3 = NumberFormat('mm/dd/yyyy')
    assert format1 == format2
    assert format1 == 'm/d/yyyy' and format1 != 'mm/dd/yyyy'
    assert format3 != 'm/d/yyyy' and format3 == 'mm/dd/yyyy'
    assert format1 != format3


def test_builtin_format(NumberFormat):
    fmt = NumberFormat(format_code='0.00')
    assert fmt.builtin_format_code(2) == fmt.format_code
