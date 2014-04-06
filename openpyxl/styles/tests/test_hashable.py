from openpyxl.styles import HashableObject
import pytest


@pytest.fixture
def Immutable():

    class Immutable(HashableObject):

        __fields__ = ('value',)

        def __init__(self, value=None):
            self.value = value

    return Immutable


class TestHashable:

    def test_ctor(self, Immutable):
        d = Immutable()
        d.value = 1
        assert d.value == 1

    def test_cannot_change_attrs(self, Immutable):
        d = Immutable()
        d.value = 1
        with pytest.raises(TypeError):
            d.value = 2

    def test_cannot_delete_attrs(self, Immutable):
        d = Immutable()
        d.value = 1
        with pytest.raises(TypeError):
            del d['value']

    def test_copy(self, Immutable):
        d = Immutable()
        d.value = 1
        #c = d.copy()

    def test_hash(self, Immutable):
        d = Immutable()
        assert hash(d) == 8982499335057739900

    def test_str(self, Immutable):
        d = Immutable()
        assert str(d) == "Immutable(value=None)"

    def test_repr(self, Immutable):
        d = Immutable()
        assert repr(d) == ""

    def test_eq(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(1)
        assert d1 is not d2
        assert d1 == d2

    def test_ne(self, Immutable):
        d1 = Immutable(1)
        d2 = Immutable(2)
        assert d1 != d2
