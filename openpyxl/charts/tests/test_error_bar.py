import pytest


class TestErrorBar(object):

    def test_ctor(self, ErrorBar):
        with pytest.raises(TypeError):
            ErrorBar(None, range(10))
