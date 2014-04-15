import pytest


class TestLineChart:

    def test_ctor(self, LineChart):
        c = LineChart()
        assert c.TYPE == "lineChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"
