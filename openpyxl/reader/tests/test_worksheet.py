from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from io import BytesIO
from zipfile import ZipFile

from lxml.etree import iterparse, fromstring

from openpyxl.utils.exceptions import InvalidFileException
from openpyxl import load_workbook
from openpyxl.compat import unicode
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.cell import Cell
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.styles import Style


def test_get_xml_iter():
    #1 file object
    #2 stream (file-like)
    #3 string
    #4 zipfile
    from openpyxl.reader.worksheet import _get_xml_iter
    from tempfile import TemporaryFile

    FUT = _get_xml_iter
    s = b""
    stream = FUT(s)
    assert isinstance(stream, BytesIO), type(stream)

    u = unicode(s)
    stream = FUT(u)
    assert isinstance(stream, BytesIO), type(stream)

    f = TemporaryFile(mode='rb+', prefix='openpyxl.', suffix='.unpack.temp')
    stream = FUT(f)
    assert stream == f
    f.close()

    t = TemporaryFile()
    z = ZipFile(t, mode="w")
    z.writestr("test", "whatever")
    stream = FUT(z.open("test"))
    assert hasattr(stream, "read")
    # z.close()
    try:
        z.close()
    except IOError:
        # you can't just close zipfiles in Windows
        if z.fp is not None:
            z.fp.close() # python 2.6
        else:
            z.close() # python 2.7


@pytest.fixture
def Worksheet(Workbook):
    from openpyxl.styles import numbers
    from openpyxl.styles.style import StyleId
    from openpyxl.worksheet.header_footer import HeaderFooter

    class DummyWorkbook:

        _guess_types = False
        data_only = False

        def __init__(self):
            self.shared_strings = IndexedList()
            self.shared_strings.add("hello world")
            self.shared_styles = 28*[DummyStyle()]
            self.shared_styles.append(Style())
            self._fonts = IndexedList()
            self._fills = IndexedList()
            self._number_formats = IndexedList()
            self._borders = IndexedList()
            self._alignments = IndexedList()
            self._protections = IndexedList()
            self._cell_styles = IndexedList()
            self.vba_archive = None
            for i in range(29):
                self._cell_styles.add((StyleId(i, i, i, i, i, i)))
            self._cell_styles.add(StyleId(fillId=4, borderId=6, alignmentId=1, protectionId=0))


    class DummyStyle:
        number_format = numbers.FORMAT_GENERAL
        font = ""
        fill = ""
        border = ""
        alignment = ""
        protection = ""

        def copy(self, **kw):
            return selfexc


    class DummyWorksheet:

        encoding = "utf-8"
        title = "Dummy"

        def __init__(self):
            self.parent = DummyWorkbook()
            self.column_dimensions = {}
            self.row_dimensions = {}
            self._styles = {}
            self.cell = None
            self._cells = {}
            self._data_validations = []
            self.header_footer = HeaderFooter()
            self.vba_controls = None

        def _add_cell(self, cell):
            self._cells[cell.coordinate] = cell

        def __getitem__(self, value):
            cell = self._cells.get(value)

            if cell is None:
                cell = Cell(self, 'A', 1)
                self._cells[value] = cell
            return cell

        def get_style(self, coordinate):
            return DummyStyle()

    return DummyWorksheet()


@pytest.fixture
def WorkSheetParser(Worksheet):
    """Setup a parser instance with an empty source"""
    from .. worksheet import WorkSheetParser
    return WorkSheetParser(Worksheet, None, {0:'a'}, {})


@pytest.fixture
def WorkSheetParserKeepVBA(Worksheet):
    """Setup a parser instance with an empty source"""
    Worksheet.parent.vba_archive=True
    from .. worksheet import WorkSheetParser
    return WorkSheetParser(Worksheet, None, {0:'a'}, {})


def test_col_width(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("complex-styles-worksheet.xml", "rb") as src:
        cols = iterparse(src, tag='{%s}col' % SHEET_MAIN_NS)
        for _, col in cols:
            parser.parse_column_dimensions(col)
    assert set(ws.column_dimensions.keys()) == set(['A', 'C', 'E', 'I', 'G'])
    assert ws.column_dimensions['A'].style_id == 0
    assert dict(ws.column_dimensions['A']) == {'max': '1', 'min': '1',
                                               'customWidth': '1',
                                               'width': '31.1640625'}


def test_hidden_col(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("hidden_rows_cols.xml", "rb") as src:
        cols = iterparse(src, tag='{%s}col' % SHEET_MAIN_NS)
        for _, col in cols:
            parser.parse_column_dimensions(col)
    assert 'D' in ws.column_dimensions
    assert dict(ws.column_dimensions['D']) == {'customWidth': '1', 'hidden':
                                               '1', 'max': '4', 'min': '4'}


def test_styled_col(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser
    with open("complex-styles-worksheet.xml", "rb") as src:
        cols = iterparse(src, tag='{%s}col' % SHEET_MAIN_NS)
        for _, col in cols:
            parser.parse_column_dimensions(col)
    assert 'I' in ws.column_dimensions
    cd = ws.column_dimensions['I']
    assert cd.style_id == 28
    assert dict(cd) ==  {'customWidth': '1', 'max': '9', 'min': '9', 'width': '25', 'style':'28'}


def test_hidden_row(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("hidden_rows_cols.xml", "rb") as src:
        rows = iterparse(src, tag='{%s}row' % SHEET_MAIN_NS)
        for _, row in rows:
            parser.parse_row_dimensions(row)
    assert 2 in ws.row_dimensions
    assert dict(ws.row_dimensions[2]) == {'hidden': '1'}


def test_styled_row(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser
    parser.shared_strings = dict((i, i) for i in range(30))
    parser.style_table = ws.parent.shared_styles

    with open("complex-styles-worksheet.xml", "rb") as src:
        rows = iterparse(src, tag='{%s}row' % SHEET_MAIN_NS)
        for _, row in rows:
            parser.parse_row_dimensions(row)
    assert 23 in ws.row_dimensions
    rd = ws.row_dimensions[23]
    assert rd.style_id == 28
    #assert rd.style == Style()
    assert dict(rd) == {'s':'28', 'customFormat':'1'}


def test_sheet_protection(datadir, Worksheet, WorkSheetParser):
    datadir.chdir()
    ws = Worksheet
    parser = WorkSheetParser

    with open("protected_sheet.xml", "rb") as src:
        tree = iterparse(src, tag='{%s}sheetProtection' % SHEET_MAIN_NS)
        for _, tag in tree:
            parser.parse_sheet_protection(tag)
    assert dict(ws.protection) == {
        'autoFilter': '0', 'deleteColumns': '0',
        'deleteRows': '0', 'formatCells': '0', 'formatColumns': '0', 'formatRows':
        '0', 'insertColumns': '0', 'insertHyperlinks': '0', 'insertRows': '0',
        'objects': '0', 'password': 'DAA7', 'pivotTables': '0', 'scenarios': '0',
        'selectLockedCells': '0', 'selectUnlockedCells': '0', 'sheet': '1', 'sort':
        '0'
    }


def test_formula_without_value(Worksheet, WorkSheetParser):
    ws = Worksheet
    parser = WorkSheetParser

    src = """
      <x:c r="A1" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:f>IF(TRUE, "y", "n")</x:f>
        <x:v />
      </x:c>
    """
    element = fromstring(src)

    parser.parse_cell(element)
    assert ws['A1'].data_type == 'f'
    assert ws['A1'].value == '=IF(TRUE, "y", "n")'


def test_formula(Worksheet, WorkSheetParser):
    ws = Worksheet
    parser = WorkSheetParser

    src = """
    <x:c r="A1" t="str" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:f>IF(TRUE, "y", "n")</x:f>
        <x:v>y</x:v>
    </x:c>
    """
    element = fromstring(src)

    parser.parse_cell(element)
    assert ws['A1'].data_type == 'f'
    assert ws['A1'].value == '=IF(TRUE, "y", "n")'


def test_formula_data_only(Worksheet, WorkSheetParser):
    ws = Worksheet
    parser = WorkSheetParser
    parser.data_only = True

    src = """
    <x:c r="A1" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:f>1+2</x:f>
        <x:v>3</x:v>
    </x:c>
    """
    element = fromstring(src)

    parser.parse_cell(element)
    assert ws['A1'].data_type == 'n'
    assert ws['A1'].value == 3


def test_string_formula_data_only(Worksheet, WorkSheetParser):
    ws = Worksheet
    parser = WorkSheetParser
    parser.data_only = True

    src = """
    <x:c r="A1" t="str" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:f>IF(TRUE, "y", "n")</x:f>
        <x:v>y</x:v>
    </x:c>
    """
    element = fromstring(src)

    parser.parse_cell(element)
    assert ws['A1'].data_type == 's'
    assert ws['A1'].value == 'y'


def test_number(Worksheet, WorkSheetParser):
    ws = Worksheet
    parser = WorkSheetParser

    src = """
    <x:c r="A1" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:v>1</x:v>
    </x:c>
    """
    element = fromstring(src)

    parser.parse_cell(element)
    assert ws['A1'].data_type == 'n'
    assert ws['A1'].value == 1


def test_string(Worksheet, WorkSheetParser):
    ws = Worksheet
    parser = WorkSheetParser

    src = """
    <x:c r="A1" t="s" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:v>0</x:v>
    </x:c>
    """
    element = fromstring(src)

    parser.parse_cell(element)
    assert ws['A1'].data_type == 's'
    assert ws['A1'].value == "a"


def test_boolean(Worksheet, WorkSheetParser):
    ws = Worksheet
    parser = WorkSheetParser

    src = """
    <x:c r="A1" t="b" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:v>1</x:v>
    </x:c>
    """
    element = fromstring(src)

    parser.parse_cell(element)
    assert ws['A1'].data_type == 'b'
    assert ws['A1'].value is True


def test_inline_string(Worksheet, WorkSheetParser, datadir):
    ws = Worksheet
    parser = WorkSheetParser
    parser.style_table = ws.parent.shared_styles
    datadir.chdir()

    with open("Table1-XmlFromAccess.xml") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}sheetData/{%s}row/{%s}c" % (SHEET_MAIN_NS, SHEET_MAIN_NS, SHEET_MAIN_NS))
    parser.parse_cell(element)
    assert ws['A1'].data_type == 's'
    assert ws['A1'].value == "ID"


def test_inline_richtext(Worksheet, WorkSheetParser, datadir):
    ws = Worksheet
    parser = WorkSheetParser
    parser.style_table = ws.parent.shared_styles
    datadir.chdir()
    with open("jasper_sheet.xml", "rb") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}sheetData/{%s}row[2]/{%s}c[18]" % (SHEET_MAIN_NS, SHEET_MAIN_NS, SHEET_MAIN_NS))
    assert element.get("r") == 'R2'
    parser.parse_cell(element)
    cell = ws['R2']
    assert cell.data_type == 's'
    assert cell.value == "11 de September de 2014"


def test_data_validation(Worksheet, WorkSheetParser, datadir):
    ws = Worksheet
    parser = WorkSheetParser
    datadir.chdir()

    with open("worksheet_data_validation.xml") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}dataValidations" % SHEET_MAIN_NS)
    parser.parse_data_validation(element)
    dvs = ws._data_validations
    assert len(dvs) == 1


def test_read_autofilter(datadir):
    datadir.chdir()
    wb = load_workbook("bug275.xlsx")
    ws = wb.active
    assert ws.auto_filter.ref == 'A1:B6'


def test_header_footer(WorkSheetParser, datadir):
    parser = WorkSheetParser
    ws = parser.ws
    datadir.chdir()

    with open("header_footer.xml") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}headerFooter" % SHEET_MAIN_NS)
    parser.parse_header_footer(element)

    assert ws.header_footer.hasHeader()
    assert ws.header_footer.left_header.font_name == "Lucida Grande,Standard"
    assert ws.header_footer.left_header.font_color == "000000"
    assert ws.header_footer.left_header.text == "Left top"
    assert ws.header_footer.center_header.text== "Middle top"
    assert ws.header_footer.right_header.text == "Right top"

    assert ws.header_footer.hasFooter()
    assert ws.header_footer.left_footer.text == "Left footer"
    assert ws.header_footer.center_footer.text == "Middle Footer"
    assert ws.header_footer.right_footer.text == "Right Footer"


def test_cell_style(WorkSheetParser, datadir):
    datadir.chdir()
    parser = WorkSheetParser
    ws = parser.ws
    parser.shared_strings[1] = "Arial Font, 10"

    with open("complex-styles-worksheet.xml") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}sheetData/{%s}row[2]/{%s}c[1]" % (SHEET_MAIN_NS, SHEET_MAIN_NS, SHEET_MAIN_NS))
    assert element.get('r') == 'A2'
    parser.parse_cell(element)
    assert ws['A2'].style_id == 2


def test_cell_exotic_style(WorkSheetParser, datadir):
    datadir.chdir()
    parser = WorkSheetParser
    ws = parser.ws
    parser.styles = [None, None, {'pivotButton':True, 'quotePrefix':True}]

    src = """
    <x:c xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="D4" s="2">
    </x:c>
    """

    sheet = fromstring(src)
    parser.parse_cell(sheet)
    assert ws['A1'].pivotButton is None

    cell = ws['D4']
    assert cell.pivotButton is True
    assert cell.quotePrefix is True


def test_sheet_views(WorkSheetParser, datadir):
    datadir.chdir()
    parser = WorkSheetParser

    with open("frozen_view_worksheet.xml") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}sheetViews" % SHEET_MAIN_NS)
    parser.parse_sheet_views(element)
    ws = parser.ws
    view = ws.sheet_view

    assert view.zoomScale == 200
    assert len(view.selection) == 3


def test_legacy_document_keep(WorkSheetParserKeepVBA, datadir):
    parser = WorkSheetParserKeepVBA
    datadir.chdir()

    with open("legacy_drawing_worksheet.xml") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}legacyDrawing" % SHEET_MAIN_NS)
    parser.parse_legacy_drawing(element)
    assert parser.ws.vba_controls == 'vbaControlId'


def test_legacy_document_no_keep(WorkSheetParser, datadir):
    parser = WorkSheetParser
    datadir.chdir()

    with open("legacy_drawing_worksheet.xml") as src:
        sheet = fromstring(src.read())

    element = sheet.find("{%s}legacyDrawing" % SHEET_MAIN_NS)
    parser.parse_legacy_drawing(element)
    assert parser.ws.vba_controls is None
