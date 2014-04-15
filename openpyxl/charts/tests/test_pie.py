import pytest


class TestPieChart:

    def test_ctor(self, PieChart):
        c = PieChart()
        assert c.TYPE, "pieChart"
