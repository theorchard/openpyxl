from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.styles import numbers


def test_builtin_format():
    fmt = '0.00'
    assert numbers.builtin_format_code(2) == fmt


def test_number_descriptor():
    from openpyxl.descriptors import Strict
    from ..numbers import NumberFormatDescriptor

    class Dummy(Strict):

        value = NumberFormatDescriptor()

        def __init__(self, value=None):
            self.value = value

    dummy = Dummy()
    assert dummy.value == "General"


@pytest.mark.parametrize("format, result",
                         [
                             ("DD/MM/YY", True),
                             ("H:MM:SS;@", True),
                             ("H:MM:SS;@", True),
                             (u'#,##0\\ [$\u20bd-46D]', False)
                         ]
                         )
def test_is_date_format(format, result):
    from ..numbers import is_date_format
    assert is_date_format(format) is result


@pytest.mark.parametrize("fmt, result",
                         [
                             ("[h]:mm:ss", True),
                             ("[hh]:mm:ss", True),
                             (u'#,##0\\ [$\u20bd-46D]', True),
                             ('"$"#,##0_);[Red]("$"#,##0)', True)
                         ]
                         )
def test_datetime_regex(fmt, result):
    from ..numbers import BAD_DATE_RE
    match = BAD_DATE_RE.search(fmt.lower()) is not None
    assert match is result
