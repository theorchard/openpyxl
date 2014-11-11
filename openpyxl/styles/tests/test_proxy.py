from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest


@pytest.fixture
def dummy_object():
    class Dummy:

        def __init__(self):
            self.a = 1
            self.b = 2
            self.c = 3

        def __repr__(self):
            return "dummy object"

    return Dummy()


@pytest.fixture
def proxy(dummy_object):
    from .. proxy import Proxy
    return Proxy(dummy_object)


def test_ctor(proxy):
    assert proxy.a == 1
    assert proxy.b == 2
    assert proxy.c == 3


def test_non_writable(proxy):
    with pytest.raises(AttributeError):
        proxy.a = 5


def test_repr(proxy):
    assert repr(proxy) == "dummy object"
