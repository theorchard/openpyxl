# Copyright (c) 2010-2014 openpyxl

import py.test

from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import safe_iterator, fromstring, iterparse

class DummyWorkbook:

    _guess_types = False
    data_only = False


class DummyWorksheet:

    def __init__(self):
        self.parent = DummyWorkbook()
        self.column_dimensions = {}
        self._styles = {}


def test_parse_col_dimensions(datadir):
    from openpyxl.reader.worksheet import WorkSheetParser

    datadir.chdir()
    ws = DummyWorksheet()

    with open("complex-styles-worksheet.xml") as src:
        parser = WorkSheetParser(ws, src, {}, {})
        tree = iterparse(parser.source)
        for _, tag in tree:
            cols = safe_iterator(tag, '{%s}col' % SHEET_MAIN_NS)
            for col in cols:
                parser.parse_column_dimensions(col)
    assert set(ws.column_dimensions.keys()) == set(['A', 'C', 'E', 'I', 'G'])
    assert dict(ws.column_dimensions['A']) == {'max': '1', 'min': '1',
                                               'customWidth': '1',
                                               'width': '31.1640625'}
