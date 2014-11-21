# Copyright (c) 2010-2014 openpyxl

import pytest

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
