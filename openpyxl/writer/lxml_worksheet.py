from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

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
    row_sort,
    get_rows_to_write,
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


def write_worksheet(worksheet, shared_strings):
    """Write a worksheet to an xml file."""

    out = BytesIO()
    NSMAP = {None : SHEET_MAIN_NS}

    with xmlfile(out) as xf:
        with xf.element('worksheet', nsmap=NSMAP):

            props = write_properties(worksheet)
            xf.write(props)

            dim = Element('dimension', {'ref': '%s' % worksheet.calculate_dimension()})
            xf.write(dim)

            xf.write(write_sheetviews(worksheet))
            xf.write(write_format(worksheet))
            cols = write_cols(worksheet)
            if cols is not None:
                xf.write(cols)
            write_rows(xf, worksheet)

            if worksheet.protection.sheet:
                prot = Element('sheetProtection', dict(worksheet.protection))
                xf.write(prot)

            af = write_autofilter(worksheet)
            if af is not None:
                xf.write(af)

            merge = write_mergecells(worksheet)
            if merge is not None:
                xf.write(merge)

            cfs = write_conditional_formatting(worksheet)
            for cf in cfs:
                xf.write(cf)

            dv = write_datavalidation(worksheet)
            if dv is not None:
                xf.write(dv)

            hyper = write_hyperlinks(worksheet)
            if hyper is not None:
                xf.write(hyper)


            options = worksheet.page_setup.options
            if len(dict(options)) > 0:
                new_element = options.write_xml_element()
                xf.write(new_element)
                del new_element

            margins = Element('pageMargins', dict(worksheet.page_margins))
            xf.write(margins)
            del margins

            setup = worksheet.page_setup.setup
            if len(dict(setup)) > 0:
                new_element = setup.write_xml_element()
                xf.write(new_element)
                del new_element

            hf = write_header_footer(worksheet)
            if hf is not None:
                xf.write(hf)

            if worksheet._charts or worksheet._images:
                drawing = Element('drawing', {'{%s}id' % REL_NS: 'rId1'})
                xf.write(drawing)
                del drawing

            # If vba is being preserved then add a legacyDrawing element so
            # that any controls can be drawn.
            if worksheet.vba_controls is not None:
                xml = Element("{%s}legacyDrawing" % SHEET_MAIN_NS,
                              {"{%s}id" % REL_NS : worksheet.vba_controls})
                xf.write(xml)

            pb = write_pagebreaks(worksheet)
            if pb is not None:
                xf.write(pb)

            # add a legacyDrawing so that excel can draw comments
            if worksheet._comment_count > 0:
                comments = Element('legacyDrawing', {'{%s}id' % REL_NS: 'commentsvml'})
                xf.write(comments)

    xml = out.getvalue()
    out.close()
    return xml


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
        attributes['s'] = '%d' % cell._style_id

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

