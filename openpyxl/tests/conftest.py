# Fixtures (pre-configured objects) for tests
import pytest

# objects under test

@pytest.fixture
def NumberFormat():
    """NumberFormat Class"""
    from openpyxl.styles import NumberFormat
    return NumberFormat


@pytest.fixture
def Image():
    """Image class"""
    from openpyxl.drawing import Image
    return Image


# Styles

@pytest.fixture
def FormatRule():
    """Formatting rule class"""
    from openpyxl.formatting.rules import FormatRule
    return FormatRule


# utility fixtures

@pytest.fixture
def ws(Workbook):
    """Empty worksheet titled 'data'"""
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    from openpyxl.tests.helper import DATADIR
    from py._path.local import LocalPath
    return LocalPath(DATADIR)
