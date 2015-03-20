from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest


@pytest.fixture
def dummy_object():
    class Dummy:

        def __init__(self, a, b, c):
            self.a = a
            self.b = b
            self.c = c

        def __repr__(self):
            return "dummy object"

        def copy(self, **kw):
            items = self.__dict__
            items.update(kw)
            return self.__class__(**items)

    return Dummy(a=1, b=2, c=3)


@pytest.fixture
def proxy(dummy_object):
    from .. proxy import StyleProxy
    return StyleProxy(dummy_object)


def test_ctor(proxy):
    assert proxy.a == 1
    assert proxy.b == 2
    assert proxy.c == 3


def test_non_writable(proxy):
    with pytest.raises(AttributeError):
        proxy.a = 5


def test_repr(proxy):
    assert repr(proxy) == "dummy object"


def test_copy(proxy):
    cp = proxy.copy(a='a')
    assert cp.a == 'a'
    assert cp.b == 2
    assert cp.c == 3


def test_invalid_proxy():
    from .. proxy import StyleProxy

    class Dummy:

        def __init__(self, a, b, c):
            self.a = a
            self.b = b
            self.c = c

    dummy = Dummy(1, 2, 3)

    with pytest.raises(TypeError):
        sp = StyleProxy(dummy)
