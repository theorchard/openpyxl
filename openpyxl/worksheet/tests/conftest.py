# Fixtures (pre-configured objects) for tests
import pytest


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    import os
    from py._path.local import LocalPath
    here = os.path.split(__file__)[0]
    DATADIR = os.path.join(here, "data")
    return LocalPath(DATADIR)
