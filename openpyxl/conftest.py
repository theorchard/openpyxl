import pytest

# Global objects under tests

@pytest.fixture
def Workbook():
    """Workbook Class"""
    from openpyxl import Workbook
    return Workbook


@pytest.fixture
def Worksheet():
    """Worksheet Class"""
    from openpyxl.worksheet import Worksheet
    return Worksheet

