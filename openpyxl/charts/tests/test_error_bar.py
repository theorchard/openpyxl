import pytest


class TestErrorBar:

    def test_ctor(self, ErrorBar):
        with pytest.raises(TypeError):
            ErrorBar(None, range(10))
