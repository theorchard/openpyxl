from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


"""Write worksheets to xml representations in an optimized way"""

import datetime
import os
from tempfile import NamedTemporaryFile

from openpyxl.compat import OrderedDict, unicode
from openpyxl.comments.comments import Comment
from openpyxl.cell import get_column_letter, Cell, TIME_TYPES
from openpyxl.styles import Style, DEFAULTS
from openpyxl.worksheet import Worksheet
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import (
    XMLGenerator,
    start_tag,
    end_tag,
    tag,
    Element,
    SubElement,
    ConditionalElement,
    tostring
)
from openpyxl.date_time import (
    to_excel,
    timedelta_to_days,
    time_to_days
)
from openpyxl.xml.constants import MAX_COLUMN, MAX_ROW, PACKAGE_XL
from openpyxl.compat import safe_string
from openpyxl.exceptions import WorkbookAlreadySaved
from openpyxl.writer.excel import ExcelWriter
from openpyxl.writer.strings import write_string_table
from openpyxl.writer.styles import StyleWriter
from openpyxl.writer.comments import CommentWriter
from .relations import write_rels
from .worksheet import write_worksheet_cols, write_worksheet_format

from openpyxl.xml.constants import (ARC_SHARED_STRINGS, PACKAGE_WORKSHEETS)


DESCRIPTORS_CACHE_SIZE = 50
BOUNDING_BOX_PLACEHOLDER = 'A1:%s%d' % (get_column_letter(MAX_COLUMN), MAX_ROW)


class CommentParentCell(object):
    __slots__ = ('coordinate', 'row', 'column')

    def __init__(self, cell):
        self.coordinate = cell.coordinate
        self.row = cell.row
        self.column = cell.column


def create_temporary_file(suffix=''):
    fobj = NamedTemporaryFile(mode='w+', suffix=suffix,
                              prefix='openpyxl.', delete=False)
    filename = fobj.name
    return filename


def WriteOnlyCell(ws=None, value=None):
    return Cell(worksheet=ws, column='A', row=1, value=value)


class DumpWorksheet(Worksheet):
    """
    .. warning::

        You shouldn't initialize this yourself, use :class:`openpyxl.workbook.Workbook` constructor instead,
        with `optimized_write = True`.
    """
    def __init__(self, parent_workbook, title):
        Worksheet.__init__(self, parent_workbook, title)

        self._max_col = 0
        self._max_row = 0
        self._parent = parent_workbook

        self._fileobj_header_name = create_temporary_file(suffix='.header')
        self._fileobj_content_name = create_temporary_file(suffix='.content')
        self._fileobj_name = create_temporary_file()

        self._strings = self._parent.shared_strings
        self._styles = self._parent.shared_styles
        self._comments = []

    def get_temporary_file(self, filename):
        if filename in self._descriptors_cache:
            fobj = self._descriptors_cache[filename]
            # re-insert the value so it does not get evicted
            # from cache soon
            del self._descriptors_cache[filename]
            self._descriptors_cache[filename] = fobj
            return fobj
        else:
            if filename is None:
                raise WorkbookAlreadySaved('this workbook has already been saved '
                        'and cannot be modified or saved anymore.')

            fobj = open(filename, 'r+')
            self._descriptors_cache[filename] = fobj
            if len(self._descriptors_cache) > DESCRIPTORS_CACHE_SIZE:
                filename, fileobj = self._descriptors_cache.popitem(last=False)
                fileobj.close()
            return fobj

    @property
    def _descriptors_cache(self):
        try:
            return self._parent._local_data.cache
        except AttributeError:
            self._parent._local_data.cache = OrderedDict()
            return self._parent._local_data.cache

    @property
    def filename(self):
        return self._fileobj_name

    @property
    def _temp_files(self):
        return (self._fileobj_content_name,
                self._fileobj_header_name,
                self._fileobj_name)

    def _unset_temp_files(self):
        self._fileobj_header_name = None
        self._fileobj_content_name = None
        self._fileobj_name = None

    def write_header(self):

        fobj = self.get_temporary_file(filename=self._fileobj_header_name)
        doc = XMLGenerator(fobj)

        start_tag(doc, 'worksheet',
                {
                'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
        start_tag(doc, 'sheetPr')
        tag(doc, 'outlinePr',
                {'summaryBelow': '1',
                'summaryRight': '1'})
        end_tag(doc, 'sheetPr')
        tag(doc, 'dimension', {'ref': 'A1:%s' % (self.get_dimensions())})
        start_tag(doc, 'sheetViews')
        start_tag(doc, 'sheetView', {'workbookViewId': '0'})
        tag(doc, 'selection', {'activeCell': 'A1',
                'sqref': 'A1'})
        end_tag(doc, 'sheetView')
        end_tag(doc, 'sheetViews')
        write_worksheet_format(doc, self)
        write_worksheet_cols(doc, self)

        return doc

    def close(self):
        self._close_content()
        self._fileobj = self.get_temporary_file(filename=self._fileobj_name)
        self._write_fileobj(self._fileobj_header_name)
        self._write_fileobj(self._fileobj_content_name)
        self._fileobj.close()

    def _write_fileobj(self, fobj_name):
        fobj = self.get_temporary_file(filename=fobj_name)
        fobj.flush()
        fobj.seek(0)

        while True:
            chunk = fobj.read(4096)
            if not chunk:
                break
            self._fileobj.write(chunk)

        fobj.close()
        self._fileobj.flush()

    def _close_content(self):
        doc = self._get_content_generator()
        end_tag(doc, 'sheetData')
        if self._comments:
            tag(doc, 'legacyDrawing', {'r:id': 'commentsvml'})
        end_tag(doc, 'worksheet')

    def get_dimensions(self):
        if not self._max_col or not self._max_row:
            return 'A1'
        else:
            return '%s%d' % (get_column_letter(self._max_col), (self._max_row))

    def _get_content_generator(self):
        """ XXX: this is ugly, but it allows to resume writing the file
        even after the handle is closed"""

        # when I'll recreate the XMLGenerator, it will start writing at the
        # begining of the file, erasing previously entered rows, so we have
        # to move to the end of the file before adding new tags
        handle = self.get_temporary_file(filename=self._fileobj_content_name)
        handle.seek(0, 2)

        return XMLGenerator(out=handle)

    def append(self, row):
        """
        :param row: iterable containing values to append
        :type row: iterable
        """
        doc = self._get_content_generator()
        cell = WriteOnlyCell(self) # singleton

        self._max_row += 1
        span = len(row)
        self._max_col = max(self._max_col, span)
        row_idx = self._max_row
        attrs = {'r': '%d' % row_idx,
                 'spans': '1:%d' % span}
        start_tag(doc, 'row', attrs)

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

            self.write_cell(doc, cell)
            if dirty_cell:
                cell = WriteOnlyCell(self)
        end_tag(doc, 'row')

    def write_cell(self, doc, cell):
        coordinate = cell.coordinate
        attributes = {'r': coordinate}
        if cell._style!= 0:
            attributes['s'] = '%d' % cell._style

        if cell.data_type != cell.TYPE_FORMULA:
            attributes['t'] = cell.data_type

        value = cell.internal_value
        if value in ('', None):
            tag(doc, 'c', attributes)
        else:
            start_tag(doc, 'c', attributes)
            if cell.data_type == cell.TYPE_STRING:
                tag(doc, 'v', body='%s' % self._strings.add(value))
            elif cell.data_type == cell.TYPE_FORMULA:
                shared_formula = self.formula_attributes.get(coordinate)
                if shared_formula is not None:
                    attr = shared_formula
                    if 't' in attr and attr['t'] == 'shared' and 'ref' not in attr:
                        # Don't write body for shared formula
                        tag(doc, 'f', attr=attr)
                    else:
                        tag(doc, 'f', attr=attr, body=value[1:])
                else:
                    tag(doc, 'f', body=value[1:])
                tag(doc, 'v')
            elif cell.data_type in (cell.TYPE_NUMERIC, cell.TYPE_BOOL):
                tag(doc, 'v', body=safe_string(value))
            else:
                tag(doc, 'v', body=value)
            end_tag(doc, 'c')


def removed_method(*args, **kw):
    raise NotImplementedError

setattr(DumpWorksheet, '__getitem__', removed_method)
setattr(DumpWorksheet, '__setitem__', removed_method)
setattr(DumpWorksheet, 'cell', removed_method)
setattr(DumpWorksheet, 'range', removed_method)
setattr(DumpWorksheet, 'merge_cells', removed_method)


def save_dump(workbook, filename):
    if workbook.worksheets == []:
        workbook.create_sheet()
    writer = ExcelDumpWriter(workbook)
    writer.save(filename)
    return True


class DumpCommentWriter(CommentWriter):
    def extract_comments(self):
        for comment in self.sheet._comments:
            if comment is not None:
                self.authors.add(comment.author)
                self.comments.append(comment)


class ExcelDumpWriter(ExcelWriter):

    def __init__(self, workbook):
        self.workbook = workbook
        self.style_writer = StyleWriter(workbook)

    def _write_string_table(self, archive):
        shared_strings = self.workbook.shared_strings
        archive.writestr(ARC_SHARED_STRINGS,
                write_string_table(shared_strings))

    def _write_worksheets(self, archive, style_writer):
        drawing_id = 1
        comments_id = 1

        for i, sheet in enumerate(self.workbook.worksheets):
            header_doc = sheet.write_header()

            start_tag(header_doc, 'sheetData')

            sheet.close()
            archive.write(sheet.filename, PACKAGE_WORKSHEETS + '/sheet%d.xml' % (i + 1))
            for filename in sheet._temp_files:
                del sheet._descriptors_cache[filename]
                os.remove(filename)
            sheet._unset_temp_files()

            # write comments
            if sheet._comments:
                rels = write_rels(sheet, drawing_id, comments_id)
                archive.writestr(
                    PACKAGE_WORKSHEETS + '/_rels/sheet%d.xml.rels' % (i + 1),
                    tostring(rels)
                        )

                cw = DumpCommentWriter(sheet)
                archive.writestr(PACKAGE_XL + '/comments%d.xml' % comments_id,
                    cw.write_comments())
                archive.writestr(PACKAGE_XL + '/drawings/commentsDrawing%d.vml' % comments_id,
                    cw.write_comments_vml())
                comments_id += 1
