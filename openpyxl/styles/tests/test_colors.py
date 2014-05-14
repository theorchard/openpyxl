from openpyxl.styles.colors import Color
import pytest


class TestColor:

    def test_ctor(self):
        c = Color()
        assert c.value == "00000000"
        assert c.type == "rgb"
        assert dict(c) == {'rgb': '00000000'}

    def test_rgb(self):
        c = Color(rgb="FFFFFFFF")
        assert c.value == "FFFFFFFF"
        assert c.type == "rgb"
        assert dict(c) == {'rgb': 'FFFFFFFF'}

    def test_indexed(self):
        c = Color(indexed=4)
        assert c.value == 4
        assert c.type == "indexed"
        assert dict(c) == {'indexed': "4"}

    def test_auto(self):
        c = Color(auto=1)
        assert c.type is "auto"
        assert c.value is True
        assert dict(c) == {'auto': "1"}

    def test_theme(self):
        c = Color(theme="1")
        assert c.value == 1
        assert c.type == "theme"
        assert dict(c) ==  {'theme': "1"}

    def test_tint(self):
        c = Color(tint=0.5)
        assert c.tint == 0.5
        assert dict(c) == {'rgb': '00000000', 'tint': "0.5"}

    def test_highlander(self):
        c = Color(rgb="FFFFFFF", indexed=4, theme=2, auto=False)
        assert c.value == 4
        assert c.type == "indexed"

    def test_validation(self):
        c = Color()
        with pytest.raises(TypeError):
            c.value = 4
