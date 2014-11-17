from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


class Comment(object):
    __slots__ = ('_parent',
                 '_text',
                 '_author',
                 '_width',
                 '_height')

    def __init__(self, text, author):
        self._text = text
        self._author = author
        self._parent = None
        self._width = '108pt'
        self._height = '59.25pt'

    @property
    def author(self):
        """ The name recorded for the author

            :rtype: string
        """
        return self._author
    @author.setter
    def author(self, value):
        self._author = value

    @property
    def text(self):
        """ The text of the commment

            :rtype: string
        """
        return self._text
    @text.setter
    def text(self, value):
        self._text = value
