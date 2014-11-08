# Copyright (c) 2010-2014 openpyxl

from .. import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
    absolute_coordinate
)

from openpyxl.exceptions import (
    CellCoordinatesException,
    )


def test_coordinates():
    column, row = coordinate_from_string('ZF46')
    assert "ZF" == column
    assert 46 == row


def test_invalid_coordinate():
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string('AAA')

def test_zero_row():
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string('AQ0')

def test_absolute():
    assert '$ZF$51' == absolute_coordinate('ZF51')

def test_absolute_multiple():

    assert '$ZF$51:$ZF$53' == absolute_coordinate('ZF51:ZF$53')
