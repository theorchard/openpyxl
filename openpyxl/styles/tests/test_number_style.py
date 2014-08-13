# Copyright (c) 2010-2014 openpyxl

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
