import pytest


class TestBarChart(object):

    def test_ctor(self, BarChart):
        c = BarChart()
        assert c.TYPE == "barChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"
