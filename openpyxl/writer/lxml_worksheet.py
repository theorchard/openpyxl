from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# Experimental writer of worksheet data using lxml incremental API

from lxml.etree import xmlfile, Element, SubElement

from openpyxl.compat import iterkeys, itervalues, safe_string
from .worksheet import row_sort


def write_worksheet_data(doc, worksheet, string_table, style_table):
    """Write worksheet data to xml."""

    # Ensure a blank cell exists if it has a style
    for styleCoord in iterkeys(worksheet._styles):
        if isinstance(styleCoord, str) and COORD_RE.search(styleCoord):
            worksheet.cell(styleCoord)

    # create rows of cells
    cells_by_row = {}
    for cell in itervalues(worksheet._cells):
        cells_by_row.setdefault(cell.row, []).append(cell)

    with xmlfile(doc) as xf:
        with xf.element("sheetData"):
            for row_idx in sorted(cells_by_row):
                # row meta data
                row_dimension = worksheet.row_dimensions[row_idx]
                row_dimension.style = worksheet._styles.get(row_idx)
                attrs = {'r': '%d' % row_idx,
                         'spans': '1:%d' % worksheet.max_column}
                attrs.update(dict(row_dimension))

                with xf.element("row", attrs):

                    row_cells = cells_by_row[row_idx]
                    for cell in sorted(row_cells, key=row_sort):
                        write_cell(xf, worksheet, cell, string_table)


def write_cell(xf, worksheet, cell, string_table):
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    cell_style = worksheet._styles.get(coordinate)
    if cell_style is not None:
        attributes['s'] = '%d' % cell_style

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell.internal_value

    if value in ('', None):
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            shared_formula = worksheet.formula_attributes.get(coordinate, {})
            if shared_formula is not None:
                if (shared_formula.get('t') == 'shared'
                    and 'ref' not in shared_formula):
                    value = None
            with xf.element('f', shared_formula):
                if value is not None:
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            value = string_table.index(value)
        with xf.element("v") as v:
            if value is not None:
                xf.write(safe_string(value))
