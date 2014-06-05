from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from .hashable import HashableObject


class Protection(HashableObject):
    """Protection options for use in styles."""
    PROTECTION_INHERIT = 'inherit'
    PROTECTION_PROTECTED = True
    PROTECTION_UNPROTECTED = False

    __fields__ = ('locked',
                  'hidden')
    __slots__ = __fields__

    def __init__(self, locked=PROTECTION_INHERIT, hidden=PROTECTION_INHERIT):
        self.locked = locked
        self.hidden = hidden
