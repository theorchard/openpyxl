from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from io import BytesIO
from lxml.etree import xmlfile, Element, SubElement, tounicode

from openpyxl.compat import safe_string
from openpyxl.cell import get_column_letter, Cell

from . dump_worksheet import DumpWorksheet, DESCRIPTORS_CACHE_SIZE, WriteOnlyCell
from . lxml_worksheet import (
    write_format,
    write_sheetviews,
    write_cols,
)

from openpyxl.xml.constants import SHEET_MAIN_NS


class LXMLWorksheet(DumpWorksheet):

    __saved = False
    writer = None

    def write_header(self):
        NSMAP = {None : SHEET_MAIN_NS}

        with xmlfile(self.filename) as xf:
            with xf.element("worksheet", nsmap=NSMAP):
                pr = Element('sheetPr')
                SubElement(pr, 'outlinePr',
                           {'summaryBelow':
                            '%d' %  (self.show_summary_below),
                            'summaryRight': '%d' % (self.show_summary_right)})
                if self.page_setup.fitToPage:
                    SubElement(pr, 'pageSetUpPr', {'fitToPage': '1'})
                xf.write(pr)

                dim = Element('dimension', {'ref': 'A1:%s' % (self.get_dimensions())})
                xf.write(dim)

                xf.write(write_sheetviews(self))
                xf.write(write_format(self))
                cols = write_cols(self)
                if cols:
                    xf.write(cols)

    def _write_row(self):
        with xmlfile(self._fileobj_content_name) as xf:
            with xf.element("sheetData"):
                try:
                    while True:
                        r = (yield)
                        xf.write(r)
                except GeneratorExit:
                    pass

    def close_content(self):
        pass

    def _get_content_generator(self):
        pass

    def append(self, row):
        """
        :param row: iterable containing values to append
        :type row: iterable
        """
        cell = WriteOnlyCell(self) # singleton

        self._max_row += 1
        span = len(row)
        self._max_col = max(self._max_col, span)
        row_idx = self._max_row
        if self.writer is None:
            self.writer = self._write_row()
            next(self.writer)

        attrs = {'r': '%d' % self._max_row,
                 'spans': '1:%d' % self._max_col}
        el = Element("row", attrs)

        for col_idx, value in enumerate(row, 1):
            if value is None:
                continue
            dirty_cell = False
            column = get_column_letter(col_idx)

            if isinstance(value, Cell):
                cell = value
                dirty_cell = True # cell may have other properties than a value
            else:
                cell.value = value

            cell.coordinate = '%s%d' % (column, row_idx)
            if cell.comment is not None:
                comment = cell.comment
                comment._parent = CommentParentCell(cell)
                self._comments.append(comment)

            tree = write_cell(self, cell)
            el.append(tree)
            if dirty_cell:
                cell = WriteOnlyCell(self)
        self.writer.send(el)


def write_cell(worksheet, cell):
    string_table = worksheet.parent.shared_strings
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    if cell.has_style:
        attributes['s'] = '%d' % cell._style

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell.internal_value

    el = Element("c", attributes)
    if value in ('', None):
        return el

    if cell.data_type == 'f':
        shared_formula = worksheet.formula_attributes.get(coordinate, {})
        if shared_formula is not None:
            if (shared_formula.get('t') == 'shared'
                and 'ref' not in shared_formula):
                value = None
        formula = SubElement(el, 'f', shared_formula)
        if value is not None:
            formula.text= value[1:]
            value = None

    if cell.data_type == 's':
        value = string_table.add(value)
    cell_content = SubElement(el, 'v')
    if value is not None:
        cell_content.text = safe_string(value)
    return el
