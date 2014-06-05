# Copyright (c) 2010-2014 openpyxl

import pytest

@pytest.fixture
def Alignment():
    from .. alignment import Alignment
    return Alignment


def test_ctor(Alignment):
    al = Alignment()
    assert dict(al) ==  {'horizontal': 'general', 'vertical': 'bottom'}
