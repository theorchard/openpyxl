from openpyxl.styles.fonts import Font

import pytest

class TestFont:

    def test_ctor(self):
        f = Font()
        assert f.name == 'Calibri'
        assert f.size == 11
        assert f.bold is False
        assert f.italic is False
        assert f.underline == 'none'
        assert f.strikethrough is False
        assert f.color.index == '00000000'
