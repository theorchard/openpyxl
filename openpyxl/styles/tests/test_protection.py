# Copyright (c) 2010-2015 openpyxl

import pytest

@pytest.fixture
def Protection():
    from .. protection import Protection
    return Protection


def test_default(Protection):
    pt = Protection()
    assert dict(pt) == {'hidden':'0', 'locked':'1'}


def test_round_trip(Protection):
    args = {'hidden':'1', 'locked':'1'}
    pt = Protection(**args)
    assert dict(pt) == args
