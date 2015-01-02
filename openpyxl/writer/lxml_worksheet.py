from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# Experimental writer of worksheet data using lxml incremental API

from io import BytesIO
from lxml.etree import xmlfile, Element, SubElement, fromstring

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

from openpyxl.formatting import ConditionalFormatting
from openpyxl.worksheet.datavalidation import writer

from .worksheet import (
    write_datavalidation,
    write_properties,
    write_sheetviews,
    write_format,
    write_cols,
    write_autofilter,
    write_mergecells,
    write_conditional_formatting,
    write_header_footer,
    write_hyperlinks,
    write_pagebreaks,
)

from .etree_worksheet import get_rows_to_write, row_sort

### LXML optimisation

def write_rows(xf, worksheet):
    """Write worksheet data to xml."""

    cells_by_row = get_rows_to_write(worksheet)

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
            if shared_formula is not None:
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
