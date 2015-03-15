from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from io import BytesIO
from operator import itemgetter

from openpyxl.compat import (
    itervalues,
    safe_string,
    iteritems
)
from openpyxl.cell import (
    column_index_from_string,
    coordinate_from_string,
)
from openpyxl.xml.constants import (
    REL_NS,
    SHEET_MAIN_NS
)


from .etree_worksheet import get_rows_to_write
from openpyxl.xml.functions import xmlfile, Element, SubElement


### LXML optimisation using xf.element to reduce instance creation

def write_rows(xf, worksheet):
    """Write worksheet data to xml."""

    all_rows = get_rows_to_write(worksheet)

    dims = worksheet.row_dimensions
    max_column = worksheet.max_column

    with xf.element("sheetData"):
        for row_idx, row in sorted(all_rows):

            attrs = {'r': '%d' % row_idx, 'spans': '1:%d' % max_column}
            if row_idx in dims:
                row_dimension = dims[row_idx]
                attrs.update(dict(row_dimension))
            with xf.element("row", attrs):

                for col, cell in sorted(row, key=itemgetter(0)):
                    if cell.value is None and not cell.has_style:
                        continue
                    write_cell(xf, worksheet, cell)


def write_cell(xf, worksheet, cell):
    string_table = worksheet.parent.shared_strings
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    if cell.has_style:
        attributes['s'] = '%d' % cell.style_id

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell.internal_value

    if value in ('', None):
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            shared_formula = worksheet.formula_attributes.get(coordinate, {})
            if (shared_formula.get('t') == 'shared'
                and 'ref' not in shared_formula):
                value = None
            with xf.element('f', shared_formula):
                if value is not None:
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            value = string_table.add(value)
        with xf.element("v"):
            if value is not None:
                xf.write(safe_string(value))
