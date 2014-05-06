from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.compat import safe_string
from openpyxl.descriptors import Set, Typed, Bool

from .colors import Color
from .hashable import HashableObject
from .descriptors import Color
from .border import Border


DIAGONAL_NONE = 0
DIAGONAL_UP = 1
DIAGONAL_DOWN = 2
DIAGONAL_BOTH = 3
diagonals = (DIAGONAL_NONE, DIAGONAL_UP, DIAGONAL_DOWN, DIAGONAL_BOTH)


class Borders(HashableObject):
    """Border positioning for use in styles."""


    __fields__ = ('left',
                  'right',
                  'top',
                  'bottom',
                  'diagonal',
                  'diagonal_direction',
                  'vertical',
                  'horizontal')

    # child elements
    left = Typed(expected_type=Border)
    right = Typed(expected_type=Border)
    top = Typed(expected_type=Border)
    bottom = Typed(expected_type=Border)
    diagonal = Typed(expected_type=Border)
    vertical = Typed(expected_type=Border, allow_none=True)
    horizontal = Typed(expected_type=Border, allow_none=True)
    # attributes
    outline = Bool()
    diagonalUp = Bool()
    diagonalDown = Bool()

    def __init__(self, left=Border(), right=Border(), top=Border(),
                 bottom=Border(), diagonal=Border(), diagonal_direction=DIAGONAL_NONE,
                 vertical=None, horizontal=None, diagonalUp=False, diagonalDown=False,
                 outline=True):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.diagonal = diagonal
        self.vertical = vertical
        self.horizontal = horizontal
        self.diagonal_direction = diagonal_direction
        self.diagonalUp = diagonalUp
        self.diagonalDown = diagonalDown
        self.outline = outline

    @property
    def children(self):
        for key in ('left', 'right', 'top', 'bottom', 'diagonal', 'vertical',
                    'horizontal'):
            value = getattr(self, key)
            if value is not None:
                yield key, value

    def __iter__(self):
        """
        Unset outline defaults to True, others default to False
        """
        for key in ('diagonalUp', 'diagonalDown', 'outline'):
            value = getattr(self, key)
            if (key == "outline" and not value
                or key != "outline" and value):
                yield key, safe_string(value)
