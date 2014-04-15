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


from openpyxl.xml.functions import Element

@pytest.fixture
def root_xml():
    """Root XML element <test>"""
    return Element("test")

@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    from openpyxl.tests.helper import DATADIR
    from py._path.local import LocalPath
    return LocalPath(DATADIR)

### Markers ###

def pytest_runtest_setup(item):
    if isinstance(item, item.Function):
        try:
            from PIL import Image
        except ImportError:
            Image = False
        if item.get_marker("pil_required") and Image is False:
            pytest.skip("PIL must be installed")
        elif item.get_marker("pil_not_installed") and Image:
            pytest.skip("PIL is installed")
        elif item.get_marker("not_py33"):
            pytest.skip("Ordering is not a given in Python 3")
        elif item.get_marker("lxml_required"):
            pytest.skip("LXML is required for some features such as schema validation")
