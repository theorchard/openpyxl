from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest


def test_dimension():
    from .. dimensions import Dimension
    with pytest.raises(TypeError):
        d = Dimension()


def test_row_dimension():
    from .. dimensions import RowDimension
    rd = RowDimension()
    assert dict(rd) == {}


def test_col_dimensions():
    from .. dimensions import ColumnDimension
    cd = ColumnDimension()
    assert dict(cd) == {'min':'1', 'max':'1'}
    assert cd.min == 1
    assert cd.max == 1
