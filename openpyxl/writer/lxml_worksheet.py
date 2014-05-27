from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# Experimental writer of worksheet data using lxml incremental API

from lxml.etree import xmlfile, Element

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

    if cell.data_type != cell.TYPE_FORMULA:
        attributes['t'] = cell.data_type

    value = cell.internal_value
    if value in ('', None):
        xf.element("c", attributes)
    else:
        with xf.element('c', attributes):
            if cell.data_type == cell.TYPE_STRING:
                el = Element('v')
                el.text = '%d' % string_table.index(value)
                xf.write(el)
                el = None
            elif cell.data_type == cell.TYPE_FORMULA:
                shared_formula = worksheet.formula_attributes.get(coordinate)
                attr = {}
                if shared_formula is not None:
                    attr = shared_formula
                    if ('t' in attr
                        and attr['t'] == 'shared'
                        and 'ref' not in attr):
                        # Don't write body for shared formula
                        value = None
                el = Element('f', attr)
                if value is not None:
                    el.text = value[1:]
                    xf.write(el)
                    el = None
                xf.write(Element("v"))
            else:
                el = Element('v')
                el.text = safe_string(value)
                xf.write(el)
                el = None
