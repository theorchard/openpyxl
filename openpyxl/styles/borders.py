from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.compat import safe_string
from openpyxl.descriptors import Set, Typed, Bool, Alias

from .colors import Color
from .hashable import HashableObject


BORDER_NONE = None
BORDER_DASHDOT = 'dashDot'
BORDER_DASHDOTDOT = 'dashDotDot'
BORDER_DASHED = 'dashed'
BORDER_DOTTED = 'dotted'
BORDER_DOUBLE = 'double'
BORDER_HAIR = 'hair'
BORDER_MEDIUM = 'medium'
BORDER_MEDIUMDASHDOT = 'mediumDashDot'
BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot'
BORDER_MEDIUMDASHED = 'mediumDashed'
BORDER_SLANTDASHDOT = 'slantDashDot'
BORDER_THICK = 'thick'
BORDER_THIN = 'thin'


class Side(HashableObject):

    spec = """Actually to BorderPr 18.8.6"""

    """Border options for use in styles.
    Caution: if you do not specify a border_style, other attributes will
    have no effect !"""

    __fields__ = ('style',
                  'color')

    color = Typed(expected_type=Color, allow_none=True)
    style = Set(values=(BORDER_NONE, BORDER_DASHDOT, BORDER_DASHDOTDOT,
                        BORDER_DASHED, BORDER_DOTTED, BORDER_DOUBLE, BORDER_HAIR, BORDER_MEDIUM,
                        BORDER_MEDIUMDASHDOT, BORDER_MEDIUMDASHDOTDOT, BORDER_MEDIUMDASHED,
                        BORDER_SLANTDASHDOT, BORDER_THICK, BORDER_THIN))
    border_style = Alias('style')

    def __init__(self, style=None, color=None, border_style=None):
        if border_style is not None:
            style = border_style
        self.style = style
        self.color = color

    def __iter__(self):
        for key in ("style",):
            value = getattr(self, key)
            if value:
                yield key, safe_string(value)


class Border(HashableObject):
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
    left = Typed(expected_type=Side)
    right = Typed(expected_type=Side)
    top = Typed(expected_type=Side)
    bottom = Typed(expected_type=Side)
    diagonal = Typed(expected_type=Side, allow_none=True)
    vertical = Typed(expected_type=Side, allow_none=True)
    horizontal = Typed(expected_type=Side, allow_none=True)
    # attributes
    outline = Bool()
    diagonalUp = Bool()
    diagonalDown = Bool()

    def __init__(self, left=Side(), right=Side(), top=Side(),
                 bottom=Side(), diagonal=Side(), diagonal_direction=None,
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
