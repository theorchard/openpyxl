from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from abc import abstractmethod, abstractproperty
from collections import namedtuple

from openpyxl.compat.abc import ABC
from openpyxl.utils.indexed_list import IndexedList

from .numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_REVERSE


class StyleProxy(object):
    """
    Proxy formatting objects so that they cannot be altered
    """

    __slots__ = ('__target')

    def __init__(self, target):
        if not hasattr(target, 'copy'):
            raise TypeError("Proxied objects must have a copy method.")
        self.__target = target


    def __repr__(self):
        return repr(self.__target)


    def __getattr__(self, attr):
        return getattr(self.__target, attr)


    def __setattr__(self, attr, value):
        if attr != "_StyleProxy__target":
            raise AttributeError("Style objects are immutable and cannot be changed."
                                 "Reassign the style with a copy")
        super(StyleProxy, self).__setattr__(attr, value)


    def copy(self, **kw):
        """Return a copy of the proxied object. Keyword args will be passed through"""
        return self.__target.copy(**kw)


    def __eq__(self, other):
        return self.__target == other


    def __ne__(self, other):
        return not self == other



StyleId = namedtuple("StyleId", "alignment border fill font number_format protection")



class StyledObject(object):
    """
    Mixin Class for stylable objects implementing proxy and lookup functions
    """

    @abstractmethod
    def __init__(self):
        self._font_id = 0
        self._fill_id = 0
        self._border_id = 0
        self._alignment_id = 0
        self._protection_id = 0
        self._number_format_id = 0
        self._style_id = 0

    @abstractproperty
    def _fonts(self):
        return IndexedList()

    @property
    def font(self):
        fo = self._fonts[self._font_id]
        if fo is not None:
            return StyleProxy(fo)

    @font.setter
    def font(self, value):
        self._font_id = self._fonts.add(value)


    @abstractproperty
    def _fills(self):
        return IndexedList()

    @property
    def fill(self):
        fo = self._fills[self._fill_id]
        return StyleProxy(fo)

    @fill.setter
    def fill(self, value):
        self._fill_id = self._fills.add(value)


    @abstractproperty
    def _borders(self):
        return IndexedList()

    @property
    def border(self):
        fo = self._borders[self._border_id]
        return StyleProxy(fo)

    @border.setter
    def border(self, value):
        self._border_id = self._borders.add(value)


    @abstractproperty
    def _alignments(self):
        return IndexedList()

    @property
    def alignment(self):
        fo = self._alignments[self._alignment_id]
        return StyleProxy(fo)

    @alignment.setter
    def alignment(self, value):
        self._alignment_id = self._alignments.add(value)


    @abstractproperty
    def _protections(self):
        return IndexedList()

    @property
    def protection(self):
        fo = self._protections[self._protection_id]
        return StyleProxy(fo)

    @protection.setter
    def protection(self, value):
        self._protection_id = self._protections.add(value)


    @abstractproperty
    def _styles(self):
        return IndexedList()

    @property
    def style(self):
        fo = self._styles[self._style_id]
        if fo is not None:
            return StyleProxy(fo)

    @style.setter
    def style(self, value):
        self._style_id = self._styles.add(value)


    @abstractproperty
    def _cell_styles(self):
        return IndexedList()

    @property
    def style_id(self):
        style = StyleId(self._alignment_id,
                        self._border_id,
                        self._fill_id,
                        self._font_id,
                        self._number_format_id,
                        self._protection_id)
        return self._cell_styles.add(style)

    @property
    def has_style(self):
        return self._alignment_id \
               or self._border_id \
               or self._fill_id \
               or self._font_id \
               or self._number_format_id \
               or self._protection_id \

    @abstractproperty
    def _number_formats(self):
        return IndexedList()

    @property
    def number_format(self):
        if self._number_format_id < 164:
            return BUILTIN_FORMATS.get(self._number_format_id, "General")
        return self._number_formats[self._number_format_id - 164]

    @number_format.setter
    def number_format(self, value):
        _id = BUILTIN_FORMATS_REVERSE.get(value)
        if _id is None:
            _id = self._number_formats.add(value) + 164
        self._number_format_id = _id
