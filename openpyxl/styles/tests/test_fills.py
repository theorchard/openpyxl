# Copyright (c) 2010-2014 openpyxl

import pytest

from openpyxl.styles.colors import BLACK, WHITE, Color


@pytest.fixture
def GradientFill():
    from openpyxl.styles.fills import GradientFill
    return GradientFill


class TestGradientFill:

    def test_empty_ctor(self, GradientFill):
        gf = GradientFill()
        assert gf.fill_type == 'linear'
        assert gf.degree == 0
        assert gf.left == 0
        assert gf.right == 0
        assert gf.top == 0
        assert gf.bottom == 0
        assert gf.stop == ()


    def test_ctor(self, GradientFill):
        gf = GradientFill(degree=90, left=1, right=2, top=3, bottom=4)
        assert gf.degree == 90
        assert gf.left == 1
        assert gf.right == 2
        assert gf.top == 3
        assert gf.bottom == 4


    def test_sequence(self, GradientFill):
        colors = [Color(BLACK), Color(WHITE)]
        gf = GradientFill(stop=colors)
        assert gf.stop == colors


    def test_invalid_sequence(self, GradientFill):
        colors = [BLACK, WHITE]
        with pytest.raises(TypeError):
            gf = GradientFill(stop=colors)


    def test_dict_interface(self, GradientFill):
        gf = GradientFill(degree=90, left=1, right=2, top=3, bottom=4)
        assert dict(gf) == {'bottom': "4", 'degree': "90", 'left':"1",
                            'right': "2", 'top': "3", 'type': 'linear'}

