from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

""" Iterators-based worksheet reader
*Still very raw*
"""
# stdlib
import operator
from itertools import groupby

# compatibility
from openpyxl.compat import range
from openpyxl.xml.functions import iterparse

# package
from openpyxl.worksheet import Worksheet
from openpyxl.utils import (
    ABSOLUTE_RE,
    coordinate_from_string,
    column_index_from_string,
    get_column_letter,
)
from openpyxl.cell import Cell
from openpyxl.cell.read_only import ReadOnlyCell, EMPTY_CELL
from openpyxl.xml.functions import safe_iterator
from openpyxl.xml.constants import SHEET_MAIN_NS


def read_dimension(source):
    min_row = min_col =  max_row = max_col = None
    DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS
    DATA_TAG = '{%s}sheetData' % SHEET_MAIN_NS
    it = iterparse(source, tag=[DIMENSION_TAG, DATA_TAG])
    for _event, element in it:
        if element.tag == DIMENSION_TAG:
            dim = element.get("ref")
            m = ABSOLUTE_RE.match(dim.upper())
            min_col, min_row, sep, max_col, max_row = m.groups()
            min_row = int(min_row)
            if max_col is None or max_row is None:
                max_col = min_col
                max_row = min_row
            else:
                max_row = int(max_row)
            return min_col, min_row, max_col, max_row

        elif element.tag == DATA_TAG:
            # Dimensions missing
            break
        element.clear()


ROW_TAG = '{%s}row' % SHEET_MAIN_NS
CELL_TAG = '{%s}c' % SHEET_MAIN_NS
VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS


class IterableWorksheet(Worksheet):

    _xml = None
    min_col = 'A'
    min_row = 1
    max_col = max_row = None

    def __init__(self, parent_workbook, title, worksheet_path,
                 xml_source, shared_strings, style_table):
        Worksheet.__init__(self, parent_workbook, title)
        self.worksheet_path = worksheet_path
        self.shared_strings = shared_strings
        self.base_date = parent_workbook.excel_base_date
        self.xml_source = xml_source
        dimensions = read_dimension(self.xml_source)
        if dimensions is not None:
            self.min_col, self.min_row, self.max_col, self.max_row = dimensions


    @property
    def xml_source(self):
        """Parse xml source on demand, default to Excel archive"""
        if self._xml is None:
            return self.parent._archive.open(self.worksheet_path)
        return self._xml


    @xml_source.setter
    def xml_source(self, value):
        self._xml = value


    def get_squared_range(self, min_col, min_row, max_col, max_row):
        """
        The source worksheet file may have columns or rows missing.
        Missing cells will be created.
        """
        if max_col is not None:
            expected_columns = [get_column_letter(ci) for ci in range(min_col, max_col + 1)]
        else:
            expected_columns = []
        row_counter = min_row

        # get cells row by row
        for row, cells in groupby(self.get_cells(min_row, min_col,
                                                 max_row, max_col),
                                  operator.attrgetter('row')):
            full_row = []
            if row_counter < row:
                # Rows requested before those in the worksheet
                for gap_row in range(row_counter, row):
                    yield tuple(EMPTY_CELL for column in expected_columns)
                    row_counter = row

            if expected_columns:
                retrieved_columns = dict([(c.column, c) for c in cells])
                for column in expected_columns:
                    if column in retrieved_columns:
                        cell = retrieved_columns[column]
                        full_row.append(cell)
                    else:
                        # create missing cell
                        full_row.append(EMPTY_CELL)
            else:
                full_row = cells

            row_counter = row + 1
            yield tuple(full_row)

    def get_cells(self, min_row, min_col, max_row, max_col):
        p = iterparse(self.xml_source, tag=[ROW_TAG], remove_blank_text=True)
        col_counter = 0
        for _event, element in p:
            if element.tag == ROW_TAG:
                row = int(element.get("r"))
                if max_row is not None and row > max_row:
                    break
                if min_row <= row:
                    for cell in safe_iterator(element, CELL_TAG):
                        col_counter += 1
                        coord = cell.get('r')
                        column_str, row = coordinate_from_string(coord)
                        column = column_index_from_string(column_str)

                        if max_col is not None and column > max_col:
                            break
                        if min_col <= column:
                            while column > col_counter:
                                # pad row with missing cells
                                yield ReadOnlyCell(self, row, col_counter, None)
                                col_counter += 1
                            data_type = cell.get('t', 'n')
                            style_id = int(cell.get('s', 0))
                            formula = cell.findtext(FORMULA_TAG)
                            value = cell.find(VALUE_TAG)
                            if value is not None:
                                value = value.text
                            if formula is not None:
                                if not self.parent.data_only:
                                    data_type = Cell.TYPE_FORMULA
                                    value = "=%s" % formula

                            yield ReadOnlyCell(self, row, column_str,
                                               value, data_type, self.parent._cell_styles[style_id])
            if element.tag in (CELL_TAG, VALUE_TAG, FORMULA_TAG):
                # sub-elements of rows should be skipped
                continue
            element.clear()

    def _get_cell(self, coordinate):
        """.iter_rows always returns a generator of rows each of which
        contains a generator of cells. This can be empty in which case
        return None"""
        result = list(self.iter_rows(coordinate))
        if result:
            return result[0][0]

    @property
    def rows(self):
        return self.iter_rows()

    def calculate_dimension(self):
        if not all([self.max_col, self.max_row]):
            raise ValueError("Worksheet is unsized, cannot calculate dimensions")
        return '%s%s:%s%s' % (self.min_col, self.min_row, self.max_col, self.max_row)

    def get_highest_column(self):
        if self.max_col is not None:
            return column_index_from_string(self.max_col)

    def get_highest_row(self):
        return self.max_row

    def get_style(self, coordinate):
        raise NotImplementedError("use `cell.style` instead")
