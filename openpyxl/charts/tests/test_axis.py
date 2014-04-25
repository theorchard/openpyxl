import pytest


@pytest.mark.parametrize("value, result",
                         [
                          (1, None),
                          (0.9, 10),
                          (0.09, 100),
                          (-0.09, 100)
                         ]
                         )
def test_less_than_one(value, result):
    from openpyxl.charts.axis import less_than_one
    assert less_than_one(value) == result

def test_axis_ctor(Axis):
    axis = Axis()
    assert axis.title == ""
    assert axis.auto_axis is True
    with pytest.raises(ZeroDivisionError):
        axis.max == 0
    with pytest.raises(ZeroDivisionError):
        axis.min == 0
    with pytest.raises(ZeroDivisionError):
        axis.unit == 0


@pytest.mark.parametrize("set_max, set_min, min, max, unit",
                         [
                         (10, 0, 0, 12, 2),
                         (5, 0, 0, 6, 1),
                         (50000, 0, 0, 60000, 12000),
                         (1, 0, 0, 2, 1),
                         (0.9, 0, 0, 1, 0.2),
                         (0.09, 0, 0, 0.1, 0.02),
                         (0, -0.09, -0.1, 0, 0.02),
                         (8, -2, -3, 10, 2)
                         ]
                         )
def test_scaling(Axis, set_max, set_min, min, max, unit):
    axis = Axis()
    axis.max = set_max
    axis.min = set_min
    assert axis.min == min
    assert axis.max == max
    assert axis.unit == unit

