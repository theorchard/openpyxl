from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.compat import safe_string
from openpyxl.descriptors import Bool

from .hashable import HashableObject


class Protection(HashableObject):
    """Protection options for use in styles."""

    __fields__ = ('locked',
                  'hidden')
    locked = Bool()
    hidden = Bool()

    def __init__(self, locked=True, hidden=False):
        self.locked = locked
        self.hidden = hidden

    def __iter__(self):
        """
        Dictionary interface for easier serialising.
        All values converted to strings
        """
        for key in self.__fields__:
            value = getattr(self, key)
            yield key, safe_string(value)
