from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest


@pytest.fixture
def BoundDictionary():
    from .. bound_dictionary import BoundDictionary
    return BoundDictionary


@pytest.mark.parametrize("default", (None, int))
def test_ctor(BoundDictionary, default):
    bd = BoundDictionary("parent", default)
    assert bd.reference == "parent"
    assert bd.default_factory == default


def test_coupling(BoundDictionary):

    class Child:

        def __init__(self, parent, index=None):
            self.parent = parent
            self.index = index

    class Parent:

        def __init__(self):
            self.children = BoundDictionary("index", self._add_child)

        def _add_child(self):
            return Child(self)

    p = Parent()
    child = p.children['A']
    assert child.parent == p
    assert child.index == 'A'
