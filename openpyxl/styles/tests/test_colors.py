from openpyxl.styles.colors import Color
import pytest


class TestColor:

    def test_ctor(self):
        c = Color()
        assert c.value == "00000000"
        assert c.type == "rgb"

    def test_validation(self):
        c = Color()
        with pytest.raises(TypeError):
            c.value = 4
