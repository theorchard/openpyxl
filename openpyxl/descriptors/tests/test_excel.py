from __future__ import absolute_import

import pytest

from .. import Strict


@pytest.fixture
def UniversalMeasure():
    from ..excel import UniversalMeasure

    class Dummy(Strict):

        value = UniversalMeasure()

    return Dummy()


class TestUniversalMeasure:

    @pytest.mark.parametrize("value",
                             ["24.73mm", "0cm", "24pt", '999pc', "50pi"]
                             )
    def test_valid(self, UniversalMeasure, value):
        UniversalMeasure.value = value
        assert UniversalMeasure.value == value

    @pytest.mark.parametrize("value",
                             [24.73, '24.73zz', "24.73 mm", None, "-24.73cm"]
                             )
    def test_invalid(self, UniversalMeasure, value):
        with pytest.raises(ValueError):
            UniversalMeasure.value = "{0}".format(value)


@pytest.fixture
def HexBinary():
    from ..excel import HexBinary

    class Dummy(Strict):

        value = HexBinary()

    return Dummy()


class TestHexBinary:

    @pytest.mark.parametrize("value",
                             ["aa35efd", "AABBCCDD"]
                             )
    def test_valid(self, HexBinary, value):
        HexBinary.value = value
        assert HexBinary.value == value


    @pytest.mark.parametrize("value",
                             ["GGII", "35.5"]
                             )
    def test_invalid(self, HexBinary, value):
        with pytest.raises(ValueError):
            HexBinary.value = value


@pytest.fixture
def TextPoint():
    from ..excel import TextPoint

    class Dummy(Strict):

        value = TextPoint()

    return Dummy()


class TestTextPoint:

    @pytest.mark.parametrize("value",
                             [-400000, "400000", 0]
                             )
    def test_valid(self, TextPoint, value):
        TextPoint.value = value
        assert TextPoint.value == int(value)

    def test_invalid_value(self, TextPoint):
        with pytest.raises(ValueError):
            TextPoint.value = -400001

    def test_invalid_type(self, TextPoint):
        with pytest.raises(TypeError):
            TextPoint.value = "40pt"


@pytest.fixture
def Percentage():
    from ..excel import Percentage

    class Dummy(Strict):

        value = Percentage()

    return Dummy()


class TestPercentage:

    @pytest.mark.parametrize("value",
                             ["15%", "15.5%"]
                             )
    def test_valid(self, Percentage, value):
        Percentage.value = value
        assert Percentage.value == value

    @pytest.mark.parametrize("value",
                             ["15", "101%", "-1%"]
                             )
    def test_valid(self, Percentage, value):
        with pytest.raises(ValueError):
            Percentage.value = value
