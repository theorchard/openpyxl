# Copyright (c) 2010-2014 openpyxl

import pytest


def test_untuple():
    from .. page import flatten
    test = Page_untuple(4)
    assert test == 4
    test = Page_untuple((4,))
    assert test == 4


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
    assert dict(p) == {'scale':'1', 'paperHeight': '24.73mm', 'orientation': 'default', '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id':'a12'}


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
