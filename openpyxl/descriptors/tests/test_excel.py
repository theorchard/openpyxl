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
