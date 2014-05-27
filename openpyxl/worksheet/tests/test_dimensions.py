from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest


def test_invalid_dimension_ctor():
    from .. dimensions import Dimension
    with pytest.raises(TypeError):
        d = Dimension()


def test_dimension_interface():
    from .. dimensions import Dimension
    d = Dimension(1, True, 1, False)
    assert dict(d) == {'hidden': '1', 'outlineLevel': '1'}


@pytest.mark.parametrize("key, value, expected",
                         [
                             ('ht', 1, {'ht':'1', 'customHeight':'1'}),
                             ('style', 10, {'s':'10', 'customFormat':'1'}),
                         ]
                         )
def test_row_dimension(key, value, expected):
    from .. dimensions import RowDimension
    rd = RowDimension()
    setattr(rd, key, value)
    assert dict(rd) == expected


@pytest.mark.parametrize("key, value, expected",
                         [
                             ('width', 1, {'width':'1', 'customWidth':'1'}),
                             ('bestFit', True, {'bestFit':'1'}),
                         ]
                         )
def test_col_dimensions(key, value, expected):
    from .. dimensions import ColumnDimension
    cd = ColumnDimension()
    setattr(cd, key, value)
    assert dict(cd) == expected
