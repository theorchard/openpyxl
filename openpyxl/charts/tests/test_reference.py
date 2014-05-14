# Copyright (c) 2010-2014 openpyxl

import pytest

@pytest.fixture
def column_of_letters(sheet, Reference):
    for idx, l in enumerate("ABCDEFGHIJ", 1):
        sheet.cell(row=idx, column=2).value = l
    return Reference(sheet, (1, 2), (10, 2))


class TestReference:

    def test_single_cell_ctor(self, cell):
        assert cell.pos1 == (1, 1)
        assert cell.pos2 is None

    def test_range_ctor(self, cell_range):
        assert cell_range.pos1 == (1, 1)
        assert cell_range.pos2 == (10, 1)

    def test_single_cell_ref(self, cell):
        assert cell.values == [0]
        assert str(cell) == "'reference'!$A$1"

    def test_cell_range_ref(self, cell_range):
        assert cell_range.values == [0, 1, 2, 3, 4, 5, 6, 7, 8 , 9]
        assert str(cell_range) == "'reference'!$A$1:$A$10"

    def test_data_type(self, cell):
        with pytest.raises(ValueError):
            cell.data_type = 'f'
            cell.data_type = None

    def test_type_inference(self, cell, cell_range, column_of_letters,
                            missing_values):
        assert cell.values == [0]
        assert cell.data_type == 'n'

        assert cell_range.values == [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
        assert cell_range.data_type == 'n'

        assert column_of_letters.values == list("ABCDEFGHIJ")
        assert column_of_letters.data_type == "s"

        assert missing_values.values == ['', '', 1, 2, 3, 4, 5, 6, 7, 8]
        missing_values.values
        assert missing_values.data_type == 'n'

    def test_number_format(self, cell):
        with pytest.raises(ValueError):
            cell.number_format = 'YYYY'
        cell.number_format = 'd-mmm'
        assert cell.number_format == 'd-mmm'

