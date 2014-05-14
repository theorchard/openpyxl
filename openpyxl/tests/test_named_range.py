# Copyright (c) 2010-2014 openpyxl

# Python stdlib imports

import pytest

# package imports
from openpyxl.namedrange import split_named_range, NamedRange
from openpyxl.reader.workbook import read_named_ranges
from openpyxl.exceptions import NamedRangeException
from openpyxl.reader.excel import load_workbook


class DummyWS:

    def __init__(self, title):
        self.title = title

    def __str__(self):
        return self.title


class DummyWB:

    def __init__(self, ws):
        self.ws = ws

    def __getitem__(self, key):
        if key == self.ws.title:
            return self.ws

    def get_sheet_names(self):
        return [self.ws.title]


def test_split():
    assert [('My Sheet', '$D$8'), ] == split_named_range("'My Sheet'!$D$8")


def test_split_no_quotes():
    assert [('HYPOTHESES', '$B$3:$L$3'), ] == split_named_range('HYPOTHESES!$B$3:$L$3')


def test_bad_range_name():
    with pytest.raises(NamedRangeException):
        split_named_range('HYPOTHESES$B$3')


def test_range_name_worksheet_special_chars(datadir):

    ws = DummyWS('My Sheeet with a , and \'')
    wb = DummyWB(ws)

    datadir.join("reader").chdir()
    with open('workbook_namedrange.xml') as src:
        content = src.read()
        named_ranges = list(read_named_ranges(content, DummyWB(ws)))
        assert len(named_ranges) == 1
        assert isinstance(named_ranges[0], NamedRange)
        assert [(ws, '$U$16:$U$24'), (ws, '$V$28:$V$36')] == named_ranges[0].destinations


def test_read_named_ranges(datadir):
    ws = DummyWS('My Sheeet')
    datadir.join("reader").chdir()

    with open("workbook.xml") as src:
        content = src.read()
        named_ranges = read_named_ranges(content, DummyWB(ws))
        assert ["My Sheeet!$D$8"] == [str(range) for range in named_ranges]


ranges_counts = (
    (4, 'TEST_RANGE'),
    (3, 'TRAP_1'),
    (13, 'TRAP_2')
)
@pytest.mark.parametrize("count, range_name", ranges_counts)
def test_oddly_shaped_named_ranges(datadir, count, range_name):

    datadir.join("genuine").chdir()
    wb = load_workbook('merge_range.xlsx')
    ws = wb.worksheets[0]
    assert len(ws.range(range_name)) == count


def test_merged_cells_named_range(datadir):
    datadir.join("genuine").chdir()

    wb = load_workbook('merge_range.xlsx')
    ws = wb.worksheets[0]
    cell = ws.range('TRAP_3')
    assert 'B15' == cell.coordinate
    assert 10 == cell.value


def test_print_titles(Workbook):
    wb = Workbook()
    ws1 = wb.create_sheet()
    ws2 = wb.create_sheet()
    ws1.add_print_title(2)
    ws2.add_print_title(3, rows_or_cols='cols')

    def mystr(nr):
        return ','.join(['%s!%s' % (sheet.title, name) for sheet, name in nr.destinations])

    actual_named_ranges = set([(nr.name, nr.scope, mystr(nr)) for nr in wb.get_named_ranges()])
    expected_named_ranges = set([('_xlnm.Print_Titles', ws1, 'Sheet1!$1:$2'),
                                 ('_xlnm.Print_Titles', ws2, 'Sheet2!$A:$C')])
    assert(actual_named_ranges == expected_named_ranges)


@pytest.mark.usefixtures("datadir")
class TestNameRefersToValue:

    def __init__(self, datadir):
        datadir.join("genuine").chdir()
        self.wb = load_workbook('NameWithValueBug.xlsx')
        self.ws = self.wb["Sheet1"]


    def test_has_ranges(self):
        ranges = self.wb.get_named_ranges()
        assert ['MyRef', 'MySheetRef', 'MySheetRef', 'MySheetValue', 'MySheetValue',
                'MyValue'] == [range.name for range in ranges]


    def test_workbook_has_normal_range(self):
        normal_range = self.wb.get_named_range("MyRef")
        assert normal_range.name == "MyRef"
        assert normal_range.destinations == [(self.ws, '$A$1')]
        assert normal_range.scope is None


    def test_workbook_has_value_range(self):
        value_range = self.wb.get_named_range("MyValue")
        assert "MyValue" == value_range.name
        assert "9.99" == value_range.value


    def test_worksheet_range(self):
        range = self.ws.range("MyRef")
        assert range.coordinate == "A1"


    def test_worksheet_range_error_on_value_range(self):
        with pytest.raises(NamedRangeException):
            self.ws.range("MyValue")


    def test_handles_scope(self):
        scoped = [
            ("MySheetRef", "Sheet1"), ("MySheetRef", "Sheet2"),
            ("MySheetValue", "Sheet1"), ("MySheetValue", "Sheet2"),
        ]
        no_scoped = ["MyRef", "MyValue"]
        ranges = self.wb.get_named_ranges()
        assert [(r.name, r.scope.title) for r in ranges if r.scope is not None] == scoped
        assert [r.name for r in ranges if r.scope is None] == no_scoped


    def test_can_be_saved(self, tmpdir):
        tmpdir.chdir()
        FNAME = "foo.xlsx"
        self.wb.save(FNAME)

        wbcopy = load_workbook(FNAME)
        ranges = wbcopy.get_named_ranges()
        names = ["MyRef", "MySheetRef", "MySheetRef", "MySheetValue", "MySheetValue", "MyValue"]
        assert [r.name for r in ranges] == names

        values = ['3.33', '14.4', '9.99']
        assert [r.value for r in ranges if hasattr(r, 'value')] == values
