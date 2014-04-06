from openpyxl.styles.colors import Color
import pytest


class TestColor:

    def test_ctor(self):
        c = Color()
        assert c.index == "FF000000"

    def test_validation(self):
        c = Color()
        with pytest.raises(TypeError):
            c.index = 4
