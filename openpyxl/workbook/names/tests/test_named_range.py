from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# Python stdlib imports

import pytest

from openpyxl.compat import zip

# package imports
from .. named_range import split_named_range, NamedRange, NamedValue
from openpyxl.workbook.names.named_range import read_named_ranges
from openpyxl.utils.exceptions import NamedRangeException
from openpyxl.reader.excel import load_workbook


class DummyWS:

    def __init__(self, title):
        self.title = title

    def __str__(self):
        return self.title


class DummyWB:

    def __init__(self, ws):
        self.ws = ws
        self.worksheets = [ws]

    def __getitem__(self, key):
        if key == self.ws.title:
            return self.ws


@pytest.mark.parametrize("range_string",
                         [
                             "'My Sheet'!$D$8",
                             "Sheet1!$A$1",
                             "[1]Sheet1!$A$1",
                             "[1]!B2range",
                         ])
def test_check_range(range_string):
    from .. named_range import refers_to_range
    assert refers_to_range(range_string) is True


@pytest.mark.parametrize("range_string, result",
                         [
                             ("'My Sheet'!$D$8", [('My Sheet', '$D$8'), ]),
                             ("Sheet1!$A$1", [('Sheet1', '$A$1')]),
                             ("[1]Sheet1!$A$1", [('[1]Sheet1', '$A$1')]),
                             ("[1]!B2range", [('[1]', '')]),
                             ("Sheet1!$C$5:$C$7,Sheet1!$C$9:$C$11,Sheet1!$E$5:$E$7,Sheet1!$E$9:$E$11,Sheet1!$D$8",
                              [('Sheet1', '$C$5:$C$7'),
                               ('Sheet1', '$C$9:$C$11'),
                               ('Sheet1', '$E$5:$E$7'),
                               ('Sheet1', '$E$9:$E$11'),
                               ('Sheet1', '$D$8')
                               ]),
                         ])
def test_split(range_string, result):
    assert list(split_named_range(range_string)) == result


@pytest.mark.parametrize("range_string, external",
                         [
                             ("'My Sheet'!$D$8", False),
                             ("Sheet1!$A$1", False),
                             ("[1]Sheet1!$A$1", True),
                             ("[1]!B2range", True),
                         ])
def test_external_range(range_string, external):
    from .. named_range import external_range
    assert external_range(range_string) is external


def test_dict_interface():
    xlrange = NamedValue("a range", "the value")
    assert dict(xlrange) == {'name':"a range"}


def test_range_scope():
    xlrange = NamedValue("Pi", 3.14)
    xlrange.scope = 1
    assert dict(xlrange) == {'name': 'Pi', 'localSheetId': 1}


def test_destinations_string():
    ws = DummyWS('Sheet1')
    xlrange = NamedRange("TRAP_2", [
        (ws, '$C$5:$C$7'),
        (ws, '$C$9:$C$11'),
        (ws, '$E$5:$E$7'),
        (ws, '$E$9:$E$11'),
        (ws, '$D$8')
                               ])
    assert xlrange.value == "'Sheet1'!$C$5:$C$7,'Sheet1'!$C$9:$C$11,'Sheet1'!$E$5:$E$7,'Sheet1'!$E$9:$E$11,'Sheet1'!$D$8"


def test_split_no_quotes():
    assert [('HYPOTHESES', '$B$3:$L$3'), ] == list(split_named_range('HYPOTHESES!$B$3:$L$3'))


def test_bad_range_name():
    with pytest.raises(NamedRangeException):
        list(split_named_range('HYPOTHESES$B$3'))


def test_range_name_worksheet_special_chars(datadir):

    ws = DummyWS('My Sheeet with a , and \'')

    datadir.chdir()
    with open('workbook_namedrange.xml') as src:
        content = src.read()
        named_ranges = list(read_named_ranges(content, DummyWB(ws)))
        assert len(named_ranges) == 1
        assert isinstance(named_ranges[0], NamedRange)
        assert [(ws, '$U$16:$U$24'), (ws, '$V$28:$V$36')] == named_ranges[0].destinations


def test_read_named_ranges(datadir):
    ws = DummyWS('My Sheeet')
    datadir.chdir()

    with open("workbook.xml") as src:
        content = src.read()
        named_ranges = read_named_ranges(content, DummyWB(ws))
        assert ["My Sheeet!$D$8"] == [str(range) for range in named_ranges]

def test_read_named_ranges_missing_sheet(datadir):
    ws = DummyWS('NOT My Sheeet')
    datadir.chdir()

    with open("workbook.xml") as src:
        content = src.read()
        named_ranges = read_named_ranges(content, DummyWB(ws))
        assert list(named_ranges) == []


def test_read_external_ranges(datadir):
    datadir.chdir()
    ws = DummyWS("Sheet1")
    wb = DummyWB(ws)
    with open("workbook_external_range.xml") as src:
        xml = src.read()
    named_ranges = list(read_named_ranges(xml, wb))
    assert len(named_ranges) == 4
    expected = [
        ("B1namedrange", "'Sheet1'!$A$1"),
        ("references_external_workbook", "[1]Sheet1!$A$1"),
        ("references_nr_in_ext_wb", "[1]!B2range"),
        ("references_other_named_range", "B1namedrange"),
    ]
    for xlr, target in zip(named_ranges, expected):
        assert xlr.name, xlr.value == target


ranges_counts = (
    (4, 'TEST_RANGE'),
    (3, 'TRAP_1'),
    (13, 'TRAP_2')
)
@pytest.mark.parametrize("count, range_name", ranges_counts)
def test_oddly_shaped_named_ranges(datadir, count, range_name):

    datadir.chdir()
    wb = load_workbook('merge_range.xlsx')
    ws = wb.worksheets[0]
    assert len(ws.get_named_range(range_name)) == count


def test_merged_cells_named_range(datadir):
    datadir.chdir()

    wb = load_workbook('merge_range.xlsx')
    ws = wb.worksheets[0]
    cell = ws.get_named_range('TRAP_3')[0]
    assert 'B15' == cell.coordinate
    assert 10 == cell.value


def test_print_titles(Workbook):
    wb = Workbook()
    ws1 = wb.create_sheet()
    ws2 = wb.create_sheet()
    scope1 = ws1._parent.worksheets.index(ws1)
    scope2 = ws2._parent.worksheets.index(ws2)
    ws1.add_print_title(2)
    ws2.add_print_title(3, rows_or_cols='cols')

    def mystr(nr):
        return ','.join(['%s!%s' % (sheet.title, name) for sheet, name in nr.destinations])

    actual_named_ranges = set([(nr.name, nr.scope, mystr(nr)) for nr in wb.get_named_ranges()])
    expected_named_ranges = set([('_xlnm.Print_Titles', scope1, 'Sheet1!$1:$2'),
                                 ('_xlnm.Print_Titles', scope2, 'Sheet2!$A:$C')])
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
        range = self.ws.get_named_range("MyRef")
        assert range.coordinate == "A1"


    def test_worksheet_range_error_on_value_range(self):
        with pytest.raises(NamedRangeException):
            self.ws.get_named_range("MyValue")


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


@pytest.mark.parametrize("value",
                         [
                             "OFFSET(rep!$AK$1,0,0,COUNT(rep!$AK$1),1)",
                             "VLOOKUP(Country!$E$3, Table_Currencies[#All], 2, 9)"
                         ])
def test_formula_names(value):
    from .. named_range import FORMULA_REGEX
    assert FORMULA_REGEX.match(value)


@pytest.mark.parametrize("value",
                         [
                             "OFFSET(rep!$AK$1,0,0,COUNT(rep!$AK$1),1)",
                             "VLOOKUP(Country!$E$3, Table_Currencies[#All], 2, 9)"
                         ])
def test_formula_not_range(value):
    from .. named_range import refers_to_range
    assert refers_to_range(value) is None
