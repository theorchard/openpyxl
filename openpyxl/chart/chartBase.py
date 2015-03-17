"""
Collection of utility primitives for charts.
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    Float,
    Typed,
    MinMax
)
from .shapes import ShapeProperties


class AxDataSource(Serialisable):

    pass

class NumDataSource(Serialisable):

    pass


class GapAmount(Serialisable):

    # needs to serialise to %
    val = MinMax(min=0, max=500)

    def __init__(self,
                 val=150,
                ):
        self.val = val


class Overlap(Serialisable):

    # needs to serialise to %

    val = MinMax(min=0, max=150)

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ChartLines(Serialisable):

    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('spPr',)

    def __init__(self,
                 spPr=None,
                ):
        self.spPr = spPr

