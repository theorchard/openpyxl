import pytest


@pytest.fixture
def list():
    from . import IndexedList
    return IndexedList


def test_ctor(list):
    l = list(['b', 'a'])
    assert l == ['a', 'b']


def test_function(list):
    l = list()
    l.append('b')
    l.append('a')
    assert l == ['b', 'a']
    assert l.values == ['a', 'b']


def test_index(list):
    l = list(['a', 'b'])
    l.append('a')
    assert l == ['a', 'b', 'a']
    l.append('c')
    assert l.index('c') == 2


def test_table_builder(list):
    sb = list()
    result = {'a':0, 'b':1, 'c':2, 'd':3}

    for letter in sorted(result.keys()):
        for x in range(5):
            sb.append(letter)
        assert sb.index(letter) == result[letter]
    assert sb.values == ['a', 'b', 'c', 'd']
