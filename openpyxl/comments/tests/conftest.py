# Copyright (c) 2010-2014 openpyxl
import pytest


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    import os
    here = os.path.split(__file__)[0]
    DATADIR = os.path.join(here, "data")
    from py._path.local import LocalPath
    return LocalPath(DATADIR)


# objects under test
