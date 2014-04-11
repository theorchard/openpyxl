import pytest


@pytest.fixture
def list():
    from . import IndexedList
    return IndexedList


def test_ctor(list):
    l = list(['b', 'a'])
    assert l == ['b', 'a']


def test_function(list):
    l = list()
    l.append('b')
    l.append('a')
    assert l == ['b', 'a']


def test_index(list):
    l = list(['a', 'b'])
    l.append('a')
    assert l == ['a', 'b']
    l.append('c')
    assert l.index('c') == 2


def test_table_builder(list):
    sb = list()
    result = {'a':0, 'b':1, 'c':2, 'd':3}

    for letter in sorted(result.keys()):
        for x in range(5):
            sb.append(letter)
        assert sb.index(letter) == result[letter]
    assert sb == ['a', 'b', 'c', 'd']
