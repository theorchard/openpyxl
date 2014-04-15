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


# Global fixtures

@pytest.fixture
def root_xml():
    """Root XML element <test>"""
    from openpyxl.xml.functions import Element
    return Element("test")

