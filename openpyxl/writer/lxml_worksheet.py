from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# Experimental writer of worksheet data using lxml incremental API

from lxml.etree import xmlfile, Element, SubElement

from openpyxl.compat import (
    iterkeys,
    itervalues,
    safe_string,
    iteritems
)
from openpyxl.cell import (
    column_index_from_string,
    coordinate_from_string
)

from .worksheet import row_sort


def write_cols(xf, worksheet, style_table=None):
    """Write worksheet columns to xml.

    style_table is ignored but required
    for compatibility with the dumped worksheet <cols> may never be empty -
    spec says must contain at least one child
    """
    cols = []
    for label, dimension in iteritems(worksheet.column_dimensions):
        dimension.style = worksheet._styles.get(label)
        col_def = dict(dimension)
        if col_def == {}:
            continue
        idx = column_index_from_string(label)
        cols.append((idx, col_def))

    if not cols:
        return

    with xf.element('cols'):
        for idx, col_def in sorted(cols):
            v = "%d" % idx
            cmin = col_def.get('min') or v
            cmax = col_def.get('max') or v
            col_def.update({'min': cmin, 'max': cmax})
            c = Element('col', col_def)
            xf.write(c)


def write_worksheet_data(xf, worksheet, string_table, style_table=None):
    """Write worksheet data to xml."""

    # Ensure a blank cell exists if it has a style
    for styleCoord in iterkeys(worksheet._styles):
        if isinstance(styleCoord, str) and COORD_RE.search(styleCoord):
            worksheet.cell(styleCoord)

    # create rows of cells
    cells_by_row = {}
    for cell in itervalues(worksheet._cells):
        cells_by_row.setdefault(cell.row, []).append(cell)

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


def write_autofilter(xf, worksheet):
    auto_filter = worksheet.auto_filter
    el = Element('autoFilter', {'ref': auto_filter.ref})
    if (auto_filter.filter_columns
        or auto_filter.sort_conditions):
        for col_id, filter_column in sorted(auto_filter.filter_columns.items()):
            fc = SubElement(el, 'filterColumn', {'colId': str(col_id)})
            attrs = {}
            if filter_column.blank:
                attrs = {'blank': '1'}
            flt = SubElement(fc, 'filters', attrs)
            for val in filter_column.vals:
                SubElement(flt, 'filter', {'val': val})
        if auto_filter.sort_conditions:
            srt = SubElement(el,  'sortState', {'ref': auto_filter.ref})
            for sort_condition in auto_filter.sort_conditions:
                sort_attr = {'ref': sort_condition.ref}
                if sort_condition.descending:
                    sort_attr['descending'] = '1'
                SubElement(srt, 'sortCondtion', sort_attr)
    xf.write(el)


def write_sheetviews(xf, worksheet):
    views = Element('sheetViews')
    view = SubElement(views, 'sheetView', {'workbookViewId': '0'})
    selectionAttrs = {}
    topLeftCell = worksheet.freeze_panes
    if topLeftCell:
        colName, row = coordinate_from_string(topLeftCell)
        column = column_index_from_string(colName)
        pane = 'topRight'
        paneAttrs = {}
        if column > 1:
            paneAttrs['xSplit'] = str(column - 1)
        if row > 1:
            paneAttrs['ySplit'] = str(row - 1)
            pane = 'bottomLeft'
            if column > 1:
                pane = 'bottomRight'
        paneAttrs.update(dict(topLeftCell=topLeftCell,
                              activePane=pane,
                              state='frozen'))
        SubElement(view, 'pane', paneAttrs)
        selectionAttrs['pane'] = pane
        if row > 1 and column > 1:
            SubElement(view, 'selection', {'pane': 'topRight'})
            SubElement(view, 'selection', {'pane': 'bottomLeft'})

    selectionAttrs.update({'activeCell': worksheet.active_cell,
                           'sqref': worksheet.selected_cell})

    SubElement(view, 'selection', selectionAttrs)
    xf.write(views)
