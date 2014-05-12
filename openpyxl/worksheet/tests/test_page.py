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

