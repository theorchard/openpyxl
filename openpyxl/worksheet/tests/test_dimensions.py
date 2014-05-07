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
    assert dict(rd) == {'collapsed': '0', 'index': '0', 'outline_level': '0',
                        'hidden': '0'}


def test_col_dimensions():
    from .. dimensions import ColumnDimension
    cd = ColumnDimension()
    assert dict(cd) == {'auto_size': '0', 'collapsed': '0', 'hidden': '0',
                      'index': 'A', 'outline_level': '0'}
