import pytest


class TestDescriptor:

    from . import Descriptor

    class Dummy:
        pass

    def test_ctor(self):
        d = self.Descriptor('key', size=1)
        assert d.name == 'key'
        assert d.size == 1

    def test_setter(self):
        d = self.Descriptor('key')
        client = self.Dummy()
        d.__set__(client, 42)
        assert client.key == 42


@pytest.fixture
def boolean():

    from . import Bool, Strict

    class Dummy(Strict):

        value = Bool()

    return Dummy()


class TestBool:

    def test_valid(self, boolean):
        boolean.value = True
        assert boolean.value


    def test_invalid(self, boolean):
        with pytest.raises(TypeError):
            boolean.value = 1


@pytest.fixture
def maximum():
    from . import Max, Strict

    class Dummy(Strict):

        value = Max(max=5)

    return Dummy()


class TestMax:

    def test_ctor(self):
        from . import Strict, Max

        with pytest.raises(TypeError):
            class Dummy(Strict):
                value = Max()

    def test_valid(self, maximum):
        maximum.value = 4
        assert maximum.value == 4

    def test_invalid(self, maximum):
        with pytest.raises(ValueError):
            maximum.value = 6


@pytest.fixture
def minimum():
    from . import Min, Strict

    class Dummy(Strict):

        value = Min(min=0)

    return Dummy()


class TestMin:

    def test_ctor(self):
        from . import Strict, Min

        with pytest.raises(TypeError):
            class Dummy(Strict):
                value = Min()


    def test_valid(self, minimum):
        minimum.value = 2
        assert minimum.value == 2


    def test_invalid(self, minimum):
        with pytest.raises(ValueError):
            minimum.value = -1


@pytest.fixture
def min_max():
    from . import MinMax, Strict

    class Dummy(Strict):

        value = MinMax(min=-1, max=1)

    return Dummy()


class TestMinMax:

    def test_ctor(self):
        from . import MinMax, Strict

        with pytest.raises(TypeError):

            class Dummy(Strict):
                value = MinMax(min=-10)

        with pytest.raises(TypeError):

            class Dummy(Strict):
                value = MinMax(max=10)


    def test_valid(self, min_max):
        min_max.value = 1
        assert min_max.value == 1


    def test_invalid(self, min_max):
        with pytest.raises(ValueError):
            min_max.value = 2
