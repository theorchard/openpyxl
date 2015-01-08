from __future__ import absolute_import
#copyright openpyxl 2010-2015

"""
Excel specific descriptors
"""

from . import MatchPattern, MinMax


class HexBinary(MatchPattern):

    pattern = "[0-9a-fA-F]+$"


class UniversalMeasure(MatchPattern):

    pattern = "[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"


class TextPoint(UniversalMeasure, MinMax):

    min = -40000
    max = 40000


class Coordinate(MinMax):

    """union of unqualified coordinate and universal measure types
    see worksheet properties for universal measure
    """

    min= -27273042329600
    max= 27273042316900


class Percentage(MatchPattern):

    pattern = "((100)|([0-9][0-9]?))(\.[0-9][0-9]?)?%"
