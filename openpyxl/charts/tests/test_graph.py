from datetime import date
import pytest


class TestGraphChart:

    def test_ctor(self, GraphChart, Axis):
        c = GraphChart()
        assert isinstance(c.x_axis, Axis)
        assert isinstance(c.y_axis, Axis)

    def test_get_x_unit(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        assert c.get_x_units() == 10

    def test_get_y_unit(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        c.y_axis.max = 10
        assert c.get_y_units() == 190500

    def test_get_y_char(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        assert c.get_y_chars() == 1

    def test_compute_series_extremes(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        mini, maxi = c._get_extremes()
        assert mini == 0
        assert maxi == 9

    def test_compute_series_max_dates(self, ws, Reference, Series, GraphChart):
        for i in range(1, 10):
            ws.append([date(2013, i, 1)])
        c = GraphChart()
        ref = Reference(ws, (1, 1), (10, 1))
        series = Series(ref)
        c.append(series)
        mini, maxi = c._get_extremes()
        assert mini == 0
        assert maxi == 41518.0

    def test_override_axis(self, GraphChart, series):
        c = GraphChart()
        c.add_serie(series)
        c.compute_axes()
        assert c.y_axis.min == 0
        assert c.y_axis.max == 10
        c.y_axis.min = -1
        c.y_axis.max = 5
        assert c.y_axis.min == -2
        assert c.y_axis.max == 6


