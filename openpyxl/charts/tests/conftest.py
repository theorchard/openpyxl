import pytest

# Charts objects under test

@pytest.fixture
def Chart():
    """Chart class"""
    from openpyxl.charts.chart import Chart
    return Chart


@pytest.fixture
def GraphChart():
    """GraphicChart class"""
    from openpyxl.charts.graph import GraphChart
    return GraphChart


@pytest.fixture
def Axis():
    """Axis class"""
    from openpyxl.charts.axis import Axis
    return Axis


@pytest.fixture
def PieChart():
    """PieChart class"""
    from openpyxl.charts import PieChart
    return PieChart


@pytest.fixture
def LineChart():
    """LineChart class"""
    from openpyxl.charts import LineChart
    return LineChart


@pytest.fixture
def BarChart():
    """BarChart class"""
    from openpyxl.charts import BarChart
    return BarChart


@pytest.fixture
def ScatterChart():
    """ScatterChart class"""
    from openpyxl.charts import ScatterChart
    return ScatterChart


@pytest.fixture
def Reference():
    """Reference class"""
    from openpyxl.charts import Reference
    return Reference


@pytest.fixture
def Series():
    """Serie class"""
    from openpyxl.charts import Series
    return Series


@pytest.fixture
def ErrorBar():
    """ErrorBar class"""
    from openpyxl.charts import ErrorBar
    return ErrorBar


# Utility fixtures

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
        ws.append([i])
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
    return Reference(sheet, (1, 1))


@pytest.fixture
def cell_range(sheet, Reference):
    return Reference(sheet, (1, 1), (10, 1))


@pytest.fixture()
def empty_range(sheet, Reference):
    for i in range(10):
        sheet.cell(row=i+1, column=2).value = None
    return Reference(sheet, (1, 2), (10, 2))


@pytest.fixture()
def missing_values(sheet, Reference):
    vals = [None, None, 1, 2, 3, 4, 5, 6, 7, 8]
    for idx, val in enumerate(vals, 1):
        sheet.cell(row=idx, column=3).value = val
    return Reference(sheet, (1, 3), (10, 3))


@pytest.fixture()
def series(cell_range, Series):
    return Series(values=cell_range)


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    import os
    here = os.path.split(__file__)[0]
    DATADIR = os.path.join(here, "data")
    from py._path.local import LocalPath
    return LocalPath(DATADIR)
