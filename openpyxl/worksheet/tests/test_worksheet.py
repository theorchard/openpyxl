# Copyright (c) 2010-2015 openpyxl

# test imports
import pytest

# compatibility imports
from openpyxl.compat import zip

# package imports
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Relationship, flatten
from openpyxl.cell import Cell, coordinate_from_string
from openpyxl.comments import Comment
from openpyxl.utils.exceptions import (
    CellCoordinatesException,
    SheetTitleException,
    InsufficientCoordinatesException,
    NamedRangeException
    )

from openpyxl.styles.colors import Color
from ..properties import WorksheetProperties


@pytest.fixture
def Worksheet():
    from ..worksheet import Worksheet
    return Worksheet


@pytest.mark.parametrize('range_string, coords',
                         [
                             ('C1:C4', (3, 1, 3, 4)),
                             ('C1', (3, 1, 3, 1)),
                         ])
def test_bounds(range_string, coords):
    from .. worksheet import range_boundaries
    assert range_boundaries(range_string) == coords


def test_cells_from_range():
    from .. worksheet import cells_from_range
    cells = cells_from_range("A1:D4")
    cells = [list(row) for row in cells]
    assert cells == [
       ['A1', 'B1', 'C1', 'D1'],
       ['A2', 'B2', 'C2', 'D2'],
       ['A3', 'B3', 'C3', 'D3'],
       ['A4', 'B4', 'C4', 'D4'],
                           ]


class TestWorksheet:

    def test_new_worksheet(self, Worksheet):
        wb = Workbook()
        ws = Worksheet(wb)
        assert ws._parent == wb

    def test_new_sheet_name(self, Worksheet):
        ws = Worksheet(Workbook(), title='')
        assert repr(ws) == '<Worksheet "Sheet2">'

    def test_get_cell(self, Worksheet):
        ws = Worksheet(Workbook())
        cell = ws.cell(row=1, column=1)
        assert cell.coordinate == 'A1'

    def test_set_bad_title(self, Worksheet):
        with pytest.raises(SheetTitleException):
            Worksheet(Workbook(), 'X' * 50)

    def test_escapes_regex_chars_in_title(self, Worksheet):
        wb = Workbook()
        ws1 = wb.create_sheet(title='Regex Test (')
        ws2 = wb.create_sheet(title='Regex Test (')
        assert ws1.title != ws2.title

    def test_increment_title(self, Worksheet):
        wb = Workbook()
        ws1 = wb.create_sheet(title="Test")
        assert ws1.title == "Test"
        ws2 = wb.create_sheet(title="Test")
        assert ws2.title == "Test1"

    @pytest.mark.parametrize("value", ["[", "]", "*", ":", "?", "/", "\\"])
    def test_set_bad_title_character(self, Worksheet, value):
        with pytest.raises(SheetTitleException):
            Worksheet(Workbook(), value)


    def test_unique_sheet_title(self, Worksheet):
        ws = Workbook().create_sheet(title="AGE")
        assert ws._unique_sheet_name("GE") == "GE"


    def test_worksheet_dimension(self, Worksheet):
        ws = Worksheet(Workbook())
        assert 'A1:A1' == ws.calculate_dimension()
        ws.cell('B12').value = 'AAA'
        assert 'A12:B12' == ws.calculate_dimension()


    def test_squared_range(self, Worksheet):
        ws = Worksheet(Workbook())
        expected = [
            ('A1', 'B1', 'C1'),
            ('A2', 'B2', 'C2'),
            ('A3', 'B3', 'C3'),
            ('A4', 'B4', 'C4'),
        ]
        rows = ws.get_squared_range(1, 1, 3, 4)
        for row, coord in zip(rows, expected):
            assert tuple(c.coordinate for c in row) == coord


    def test_iter_rows(self, Worksheet):
        ws = Worksheet(Workbook())
        expected = [
            ('A1', 'B1', 'C1'),
            ('A2', 'B2', 'C2'),
            ('A3', 'B3', 'C3'),
            ('A4', 'B4', 'C4'),
        ]

        rows = ws.iter_rows('A1:C4')
        for row, coord in zip(rows, expected):
            assert tuple(c.coordinate for c in row) == coord


    def test_iter_rows_offset(self, Worksheet):
        ws = Worksheet(Workbook())
        rows = ws.iter_rows('A1:C4', 1, 3)
        expected = [
            ('D2', 'E2', 'F2'),
            ('D3', 'E3', 'F3'),
            ('D4', 'E4', 'F4'),
            ('D5', 'E5', 'F5'),
        ]

        for row, coord in zip(rows, expected):
            assert tuple(c.coordinate for c in row) == coord


    def test_worksheet(self, Worksheet, recwarn):
        ws = Worksheet(Workbook())
        rows = ws.range("A1:D4")
        w = recwarn.pop()
        assert issubclass(w.category, UserWarning)


    def test_get_named_range(self, Worksheet):
        wb = Workbook()
        ws = Worksheet(wb)
        wb.create_named_range('test_range', ws, 'C5')
        xlrange = tuple(ws.get_named_range('test_range'))
        cell = xlrange[0]
        assert isinstance(cell, Cell)
        assert cell.row == 5


    def test_get_bad_named_range(self, Worksheet):
        ws = Worksheet(Workbook())
        with pytest.raises(NamedRangeException):
            ws.get_named_range('bad_range')


    def test_get_named_range_wrong_sheet(self, Worksheet):
        wb = Workbook()
        ws1 = Worksheet(wb)
        ws2 = Worksheet(wb)
        wb.create_named_range('wrong_sheet_range', ws1, 'C5')
        with pytest.raises(NamedRangeException):
            ws2.get_named_range('wrong_sheet_range')


    def test_cell_alternate_coordinates(self, Worksheet):
        ws = Worksheet(Workbook())
        cell = ws.cell(row=8, column=4)
        assert 'D8' == cell.coordinate

    def test_cell_insufficient_coordinates(self, Worksheet):
        ws = Worksheet(Workbook())
        with pytest.raises(InsufficientCoordinatesException):
            ws.cell(row=8)

    def test_cell_range_name(self, Worksheet):
        wb = Workbook()
        ws = Worksheet(wb)
        wb.create_named_range('test_range_single', ws, 'B12')
        c_range_name = ws.get_named_range('test_range_single')
        c_range_coord = tuple(tuple(ws.iter_rows('B12'))[0])
        c_cell = ws.cell('B12')
        assert c_range_coord == (c_cell,)
        assert c_range_name == (c_cell,)


    def test_garbage_collect(self, Worksheet):
        ws = Worksheet(Workbook())
        ws.cell('A1').value = ''
        ws.cell('B2').value = '0'
        ws.cell('C4').value = 0
        ws.cell('D1').comment = Comment('Comment', 'Comment')
        ws._garbage_collect()
        assert set(ws.get_cell_collection()), set([ws.cell('B2'), ws.cell('C4') == ws.cell('D1')])


    def test_hyperlink_value(self, Worksheet):
        ws = Worksheet(Workbook())
        ws.cell('A1').hyperlink = "http://test.com"
        assert "http://test.com" == ws.cell('A1').value
        ws.cell('A1').value = "test"
        assert "test" == ws.cell('A1').value


    def test_hyperlink_relationships(self, Worksheet):
        ws = Worksheet(Workbook())
        assert len(ws.relationships) == 0

        ws.cell('A1').hyperlink = "http://test.com"
        assert len(ws.relationships) == 1
        assert "rId1" == ws.cell('A1').hyperlink_rel_id
        assert "rId1" == ws.relationships[0].id
        assert "http://test.com" == ws.relationships[0].target
        assert "External" == ws.relationships[0].target_mode

        ws.cell('A2').hyperlink = "http://test2.com"
        assert len(ws.relationships) == 2
        assert "rId2" == ws.cell('A2').hyperlink_rel_id
        assert "rId2" == ws.relationships[1].id
        assert "http://test2.com" == ws.relationships[1].target
        assert "External" == ws.relationships[1].target_mode

    def test_bad_relationship_type(self, Worksheet):
        with pytest.raises(ValueError):
            Relationship('bad_type')


    def test_append(self, Worksheet):
        ws = Worksheet(Workbook())
        ws.append(['value'])
        assert ws['A1'].value == "value"
        assert ws.row_dimensions[1].parent is ws
        assert ws.column_dimensions['A'].parent is ws


    def test_append_list(self, Worksheet):
        ws = Worksheet(Workbook())

        ws.append(['This is A1', 'This is B1'])

        assert 'This is A1' == ws.cell('A1').value
        assert 'This is B1' == ws.cell('B1').value

    def test_append_dict_letter(self, Worksheet):
        ws = Worksheet(Workbook())

        ws.append({'A' : 'This is A1', 'C' : 'This is C1'})

        assert 'This is A1' == ws.cell('A1').value
        assert 'This is C1' == ws.cell('C1').value

    def test_append_dict_index(self, Worksheet):
        ws = Worksheet(Workbook())

        ws.append({1 : 'This is A1', 3 : 'This is C1'})

        assert 'This is A1' == ws.cell('A1').value
        assert 'This is C1' == ws.cell('C1').value

    def test_bad_append(self, Worksheet):
        ws = Worksheet(Workbook())
        assert ws.max_row == 0
        with pytest.raises(TypeError):
            ws.append("test")
        assert ws.max_row == 0


    def test_append_range(self, Worksheet):
        ws = Worksheet(Workbook())
        ws.append(range(30))
        assert ws['AD1'].value == 29


    def test_append_iterator(self, Worksheet):
        def itty():
            for i in range(30):
                yield i

        ws = Worksheet(Workbook())
        gen = itty()
        ws.append(gen)
        assert ws['AD1'].value == 29


    def test_append_2d_list(self, Worksheet):

        ws = Worksheet(Workbook())

        ws.append(['This is A1', 'This is B1'])
        ws.append(['This is A2', 'This is B2'])

        vals = ws.iter_rows('A1:B2')
        expected = (
            ('This is A1', 'This is B1'),
            ('This is A2', 'This is B2'),
        )
        for e, v in zip(expected, flatten(vals)):
            assert e == tuple(v)


    def test_append_cell(self, Worksheet):
        from openpyxl.cell import Cell

        cell = Cell(None, 'A', 1, 25)

        ws = Worksheet(Workbook())
        ws.append([])

        ws.append([cell])

        assert ws['A2'].value == 25


    @pytest.mark.parametrize("row, column, coordinate",
                             [
                                 (1, 0, 'A1'),
                                 (9, 2, 'C9'),
                             ])
    def test_iter_rows(self, Worksheet, row, column, coordinate):
        from itertools import islice
        ws = Worksheet(Workbook())
        ws.cell('A1').value = 'first'
        ws.cell('C9').value = 'last'
        assert ws.calculate_dimension() == 'A1:C9'
        rows = ws.iter_rows()
        first_row = tuple(next(islice(rows, row - 1, row)))
        assert first_row[column].coordinate == coordinate


    def test_rows(self, Worksheet):

        ws = Worksheet(Workbook())

        ws.cell('A1').value = 'first'
        ws.cell('C9').value = 'last'

        rows = ws.rows

        assert len(rows) == 9
        first_row = rows[0]
        last_row = rows[-1]

        assert first_row[0].value == 'first' and first_row[0].coordinate == 'A1'
        assert last_row[-1].value == 'last'


    def test_no_cols(self, Worksheet):
        ws = Worksheet(Workbook())
        assert ws.columns == ((),)


    def test_cols(self, Worksheet):
        ws = Worksheet(Workbook())

        ws.cell('A1').value = 'first'
        ws.cell('C9').value = 'last'
        expected = [
            ('A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9'),
            ('B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9'),
            ('C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9'),

        ]

        cols = ws.columns
        for col, coord in zip(cols, expected):
            assert tuple(c.coordinate for c in col) == coord

        assert len(cols) == 3

        assert cols[0][0].value == 'first'
        assert cols[-1][-1].value == 'last'

    def test_auto_filter(self, Worksheet):
        ws = Worksheet(Workbook())
        ws.auto_filter.ref = ws.iter_rows('a1:f1')
        assert ws.auto_filter.ref == 'A1:F1'

        ws.auto_filter.ref = ''
        assert ws.auto_filter.ref is None

        ws.auto_filter.ref = 'c1:g9'
        assert ws.auto_filter.ref == 'C1:G9'

    def test_getitem(self, Worksheet):
        ws = Worksheet(Workbook())
        c = ws['A1']
        assert isinstance(c, Cell)
        assert c.coordinate == "A1"
        assert ws['A1'].value is None

    def test_setitem(self, Worksheet):
        ws = Worksheet(Workbook())
        ws['A12'] = 5
        assert ws['A12'].value == 5

    def test_getslice(self, Worksheet):
        ws = Worksheet(Workbook())
        cell_range = ws['A1':'B2']
        assert tuple(cell_range) == (
            (ws['A1'], ws['B1']),
            (ws['A2'], ws['B2'])
        )


    def test_freeze(self, Worksheet):
        ws = Worksheet(Workbook())
        ws.freeze_panes = ws.cell('b2')
        assert ws.freeze_panes == 'B2'

        ws.freeze_panes = ''
        assert ws.freeze_panes is None

        ws.freeze_panes = 'c5'
        assert ws.freeze_panes == 'C5'

        ws.freeze_panes = ws.cell('A1')
        assert ws.freeze_panes is None


    def test_merged_cells_lookup(self, Worksheet):
        ws = Worksheet(Workbook())
        ws._merged_cells.append("A1:N50")
        merged = ws.merged_cells
        assert 'A1' in merged
        assert 'N50' in merged
        assert 'A51' not in merged
        assert 'O1' not in merged


    def test_merged_cell_ranges(self, Worksheet):
        ws = Worksheet(Workbook())
        assert ws.merged_cell_ranges == []


    def test_merge_range_string(self, Worksheet):
        ws = Worksheet(Workbook())
        ws['A1'] = 1
        ws['D4'] = 16
        ws.merge_cells(range_string="A1:D4")
        assert ws._merged_cells == ["A1:D4"]
        assert 'D4' not in ws._cells


    def test_merge_coordinate(self, Worksheet):
        ws = Worksheet(Workbook())
        ws.merge_cells(start_row=1, start_column=1, end_row=4, end_column=4)
        assert ws._merged_cells == ["A1:D4"]


    def test_unmerge_range_string(self, Worksheet):
        ws = Worksheet(Workbook())
        ws._merged_cells = ["A1:D4"]
        ws.unmerge_cells("A1:D4")


    def test_unmerge_coordinate(self, Worksheet):
        ws = Worksheet(Workbook())
        ws._merged_cells = ["A1:D4"]
        ws.unmerge_cells(start_row=1, start_column=1, end_row=4, end_column=4)
        
    
    def test_print_titles(self):
        wb = Workbook()
        ws = wb.active
        scope = wb._active_sheet_index
        ws.add_print_title(1, rows_or_cols='rows')
        print_titles = wb.get_named_range('_xlnm.Print_Titles')
        assert print_titles.name == '_xlnm.Print_Titles'
        assert str(print_titles.destinations) == """[(<Worksheet "Sheet">, '$1:$1')]"""
        assert print_titles.scope == scope


class TestPositioning(object):
    def test_point(self):
        wb = Workbook()
        ws = wb.get_active_sheet()
        assert ws.point_pos(top=40, left=150), ('C' == 3)

    @pytest.mark.parametrize("value", ('A1', 'D52', 'X11'))
    def test_roundtrip(self, value):
        wb = Workbook()
        ws = wb.get_active_sheet()
        assert ws.point_pos(*ws.cell(value).anchor) == coordinate_from_string(value)


def test_freeze_panes_horiz(Worksheet):
    ws = Worksheet(Workbook())
    ws.freeze_panes = 'A4'

    view = ws.sheet_view
    assert len(view.selection) == 1
    assert dict(view.selection[0]) == {'activeCell': 'A1', 'pane': 'bottomLeft', 'sqref': 'A1'}
    assert dict(view.pane) == {'activePane': 'bottomLeft', 'state': 'frozen',
                               'topLeftCell': 'A4', 'ySplit': '3'}


def test_freeze_panes_vert(Worksheet):
    ws = Worksheet(Workbook())
    ws.freeze_panes = 'D1'

    view = ws.sheet_view
    assert len(view.selection) == 1
    assert dict(view.selection[0]) ==  {'activeCell': 'A1', 'pane': 'topRight', 'sqref': 'A1'}
    assert dict(view.pane) == {'activePane': 'topRight', 'state': 'frozen',
                               'topLeftCell': 'D1', 'xSplit': '3'}


def test_freeze_panes_both(Worksheet):
    ws = Worksheet(Workbook())
    ws.freeze_panes = 'D4'

    view = ws.sheet_view
    assert len(view.selection) == 3
    assert dict(view.selection[0]) == {'pane': 'topRight'}
    assert dict(view.selection[1]) == {'pane': 'bottomLeft',}
    assert dict(view.selection[2]) == {'activeCell': 'A1', 'pane': 'bottomRight', 'sqref': 'A1'}
    assert dict(view.pane) == {'activePane': 'bottomRight', 'state': 'frozen',
                               'topLeftCell': 'D4', 'xSplit': '3', "ySplit":"3"}
