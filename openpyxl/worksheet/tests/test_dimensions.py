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
    assert dict(cd) == {}


def test_group_columns_simple():
    from ..worksheet import Worksheet
    from ..dimensions import ColumnDimension
    class DummyWorkbook:
        def get_sheet_names(self):
            return []
    ws = Worksheet(parent_workbook=DummyWorkbook())
    dims = ws.column_dimensions
    dims.group('A', 'C', 1)
    assert len(dims) == 1
    group = dims.values()[0]
    assert group.outline_level == 1
    assert group.min == 1
    assert group.max == 3


def test_group_columns_collapse():
    from ..worksheet import Worksheet
    from ..dimensions import ColumnDimension
    class DummyWorkbook:
        def get_sheet_names(self):
            return []
    ws = Worksheet(parent_workbook=DummyWorkbook())
    dims = ws.column_dimensions
    dims.group('A', 'C', 1, hidden=True)
    group = dims.values()[0]
    assert group.hidden
