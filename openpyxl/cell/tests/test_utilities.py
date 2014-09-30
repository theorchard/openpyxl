from ..cell import get_column_interval


def test_column_interval():
    expected = ['A', 'B', 'C', 'D']
    assert get_column_interval('A', 'D') == expected
    assert get_column_interval('A', 4) == expected
