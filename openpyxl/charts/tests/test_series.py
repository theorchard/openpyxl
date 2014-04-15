import pytest


class TestSeries:

    def test_ctor(self, Series, cell):
        series = Series(cell)
        assert series.values == [0]
        assert series.color == None
        assert series.error_bar == None
        assert series.xvalues == None
        assert series.labels == None
        assert series.title == None

    def test_invalid_values(self, Series, cell):
        series = Series(cell)
        with pytest.raises(TypeError):
            series.values = 0

    def test_invalid_xvalues(self, Series, cell):
        series = Series(cell)
        with pytest.raises(TypeError):
            series.xvalues = 0

    def test_color(self, Series, cell):
        series = Series(cell)
        assert series.color == None
        series.color = "blue"
        assert series.color == "blue"
        with pytest.raises(ValueError):
            series.color = None

    def test_min(self, Series, cell, cell_range, empty_range):
        series = Series(cell)
        assert series.min() == 0
        series = Series(cell_range)
        assert series.min() == 0
        series = Series(empty_range)
        assert series.min() == None

    def test_max(self, Series, cell, cell_range, empty_range):
        series = Series(cell)
        assert series.max() == 0
        series = Series(cell_range)
        assert series.max() == 9
        series = Series(empty_range)
        assert series.max() == None

    def test_min_max(self, Series, cell, cell_range, empty_range):
        series = Series(cell)
        assert series.get_min_max() == (0, 0)
        series = Series(cell_range)
        assert series.get_min_max() == (0, 9)
        series = Series(empty_range)
        assert series.get_min_max() == (None, None)

    def test_len(self, Series, cell):
        series = Series(cell)
        assert len(series) == 1

    def test_error_bar(self, Series, ErrorBar, cell):
        series = Series(cell)
        series.error_bar = ErrorBar(None, cell)
        assert series.get_min_max() == (0, 0)
