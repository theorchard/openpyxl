import pytest


class TestScatterChart:

    def test_ctor(self, ScatterChart):
        c = ScatterChart()
        assert c.TYPE == "scatterChart"
        assert c.x_axis.type == "valAx"
        assert c.y_axis.type == "valAx"
