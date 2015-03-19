from __future__ import absolute_import
#copyright openpyxl 2010-2015

"""
Excel specific descriptors
"""

from openpyxl.compat import basestring
from . import MatchPattern, MinMax, Integer, String, Typed
from .serialisable import Serialisable


class HexBinary(MatchPattern):

    pattern = "[0-9a-fA-F]+$"


class UniversalMeasure(MatchPattern):

    pattern = "[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"


class TextPoint(MinMax):
    """
    Size in hundredths of points.
    In theory other units of measurement can be used but these are unbounded
    """
    expected_type = int

    min = -400000
    max = 400000


Coordinate = Integer


class Percentage(MatchPattern):

    pattern = "((100)|([0-9][0-9]?))(\.[0-9][0-9]?)?%"


class Extension(Serialisable):

    uri = Typed(expected_type=String, )

    def __init__(self,
                 uri=None,
                ):
        self.uri = uri


class ExtensionList(Serialisable):

    ext = Typed(expected_type=Extension, allow_none=True)

    def __init__(self,
                 ext=None,
                ):
        self.ext = ext
