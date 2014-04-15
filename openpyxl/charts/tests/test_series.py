import pytest

@pytest.fixture
def Series():
    """Serie class"""
    from openpyxl.charts import Series
    return Series


@pytest.fixture
def Workbook():
    """Workbook Class"""
    from openpyxl import Workbook
    return Workbook


@pytest.fixture
def ws(Workbook):
    """Empty worksheet titled 'data'"""
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def ten_row_sheet(ws):
    """Worksheet with values 0-9 in the first column"""
    for i in range(10):
        ws.cell(row=i, column=0).value = i
    return ws


@pytest.fixture
def ten_column_sheet(ws):
    """Worksheet with values 0-9 in the first row"""
    ws.append(list(range(10)))
    return ws


@pytest.fixture
def sheet(ten_row_sheet):
    ten_row_sheet.title = "reference"
    return ten_row_sheet


@pytest.fixture
def cell(sheet, Reference):
    return Reference(sheet, (0, 0))


@pytest.fixture
def cell_range(sheet, Reference):
    return Reference(sheet, (0, 0), (9, 0))


@pytest.fixture()
def empty_range(sheet, Reference):
    for i in range(10):
        sheet.cell(row=i, column=1).value = None
    return Reference(sheet, (0, 1), (9, 1))


class TestSerie(object):

    def test_ctor(self, Series, cell):
        series = Series(cell)
        assert series.values == [0]
        assert series.color == None
        assert series.error_bar == None
        assert series.xvalues == None
        assert series.labels == None
        assert series.title == None

    def test_invalid_values(self, Series, cell):
        series = Series(cell)
        with pytest.raises(TypeError):
            series.values = 0

    def test_invalid_xvalues(self, Series, cell):
        series = Series(cell)
        with pytest.raises(TypeError):
            series.xvalues = 0

    def test_color(self, Series, cell):
        series = Series(cell)
        assert series.color == None
        series.color = "blue"
        assert series.color, "blue"
        with pytest.raises(ValueError):
            series.color = None

    def test_min(self, Series, cell, cell_range, empty_range):
        series = Series(cell)
        assert series.min() == 0
        series = Series(cell_range)
        assert series.min() == 0
        series = Series(empty_range)
        assert series.min() == None

    def test_max(self, Series, cell, cell_range, empty_range):
        series = Series(cell)
        assert series.max() == 0
        series = Series(cell_range)
        assert series.max() == 9
        series = Series(empty_range)
        assert series.max() == None

    def test_min_max(self, Series, cell, cell_range, empty_range):
        series = Series(cell)
        assert series.get_min_max() == (0, 0)
        series = Series(cell_range)
        assert series.get_min_max() == (0, 9)
        series = Series(empty_range)
        assert series.get_min_max() == (None, None)

    def test_len(self, Series, cell):
        series = Series(cell)
        assert len(series) == 1

    def test_error_bar(self, Series, ErrorBar, cell):
        series = Series(cell)
        series.error_bar = ErrorBar(None, cell)
        assert series.get_min_max() == (0, 0)



