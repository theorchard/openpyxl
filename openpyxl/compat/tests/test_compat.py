# Copyright (c) 2010-2014 openpyxl
import pytest


@pytest.mark.parametrize("value, result",
                         [
                          ('s', 's'),
                          (2.0/3, '0.6666666666666666'),
                          (1, '1'),
                          (None, 'None')
                         ]
                         )
def test_safe_string(value, result):
    from openpyxl.compat import safe_string
    assert safe_string(value) == result
    v = safe_string('s')
    assert v == 's'


@pytest.fixture
def dictionary():
    return {'1':1, 'a':'b', 3:'d'}


def test_iterkeys(dictionary):
    from openpyxl.compat import iterkeys
    d = dictionary
    assert set(iterkeys(d)) == set(['1', 'a', 3])


def test_iteritems(dictionary):
    from openpyxl.compat import iteritems
    d = dictionary
    assert set(iteritems(d)) == set([(3, 'd'), ('1', 1), ('a', 'b')])


def test_itervalues(dictionary):
    from openpyxl.compat import itervalues
    d = dictionary
    assert set(itervalues(d)) == set([1, 'b', 'd'])
