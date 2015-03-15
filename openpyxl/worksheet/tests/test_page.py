from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl.xml.functions import tostring
from openpyxl.tests.helper import compare_xml

@pytest.fixture
def PageMargins():
    from .. page import PageMargins
    return PageMargins


def test_empty_ctor(PageMargins):
    pm = PageMargins()
    assert pm.left == 0.75
    assert pm.right == 0.75
    assert pm.top == 1
    assert pm.bottom == 1
    assert pm.header == 0.5
    assert pm.footer == 0.5


def test_dict_interface(PageMargins):
    pm = PageMargins()
    assert dict(pm) == {'bottom': '1', 'footer': '0.5', 'header': '0.5',
                        'left': '0.75', 'right': '0.75', 'top': '1'}


@pytest.fixture
def PageSetup():
    from .. page import PageSetup
    return PageSetup


def test_page_setup(PageSetup):
    p = PageSetup()
    assert dict(p) == {}
    p.scale = 1
    assert p.scale == 1
    p.paperHeight = "24.73mm"
    assert p.paperHeight == "24.73mm"
    assert p.cellComments == None
    p.orientation = "default"
    assert p.orientation == "default"
    p.id = 'a12'
    assert dict(p) == {'scale':'1', 'paperHeight': '24.73mm', 'orientation': 'default', 'id':'a12'}


def test_wrong_page_setup(PageSetup):
    p = PageSetup()
    """ providing a not standard parameter """
    with pytest.raises(ValueError):
        p.orientation = "tagada"


@pytest.fixture
def PrintOptions():
    from .. page import PrintOptions
    return PrintOptions


def test_print_options(PrintOptions):
    p = PrintOptions()
    assert dict(p) == {}
    p.horizontalCentered = True
    p.verticalCentered = True
    assert dict(p) == {'verticalCentered': '1', 'horizontalCentered': '1'}


def test_page_margins(PageMargins):
    page_margins = PageMargins()
    page_margins.left = 2.0
    page_margins.right = 2.0
    page_margins.top = 2.0
    page_margins.bottom = 2.0
    page_margins.header = 1.5
    page_margins.footer = 1.5
    xml = tostring(page_margins.to_tree())
    expected = """
    <pageMargins xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" left="2" right="2" top="2" bottom="2" header="1.5" footer="1.5"/>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_printer_settings(PageSetup):
    page_setup = PageSetup()
    page_setup.orientation = "landscape"
    page_setup.paperSize = 3
    page_setup.fitToHeight = False
    page_setup.fitToWidth = True
    xml = tostring(page_setup.to_tree())
    expected = """
    <pageSetup orientation="landscape" paperSize="3" fitToHeight="0" fitToWidth="1"/>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_print_options(PrintOptions):
    print_options = PrintOptions()
    print_options.horizontalCentered = True
    print_options.verticalCentered = True
    xml = tostring(print_options.to_tree())
    expected = """
    <printOptions horizontalCentered="1" verticalCentered="1"/>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
