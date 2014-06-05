from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.descriptors import Strict, Bool, String, Alias, Integer
from openpyxl.compat import safe_string

from .password_hasher import hash_password


class SheetProtection(Strict):
    """
    Information about protection of various aspects of a sheet.
    True values mean that protection for the object or action is active
    This is the **default** when protection is active, ie. users cannot do something
    """

    sheet = Bool()
    enabled = Alias('sheet')
    objects = Bool()
    scenarios = Bool()
    formatCells = Bool()
    formatColumns = Bool()
    formatRows = Bool()
    insertColumns = Bool()
    insertRows = Bool()
    insertHyperlinks = Bool()
    deleteColumns = Bool()
    deleteRows = Bool()
    selectLockedCells = Bool()
    selectUnlockedCells = Bool()
    sort = Bool()
    autoFilter = Bool()
    pivotTables = Bool()
    password = String()
    saltValue = String(allow_none=True)
    spinCount = Integer(allow_none=True)
    algorithmName = String(allow_none=True)

    _password = None


    def __init__(self, sheet=False, objects=False, scenarios=False,
                 formatCells=True, formatRows=True, formatColumns=True,
                 insertColumns=True, insertRows=True, insertHyperlinks=True,
                 deleteColumns=True, deleteRows=True, selectLockedCells=False,
                 selectUnlockedCells=False, sort=True, autoFilter=True, pivotTables=True,
                 password=None, algorithmName=None, saltValue=None, spinCount=None):
        self.sheet = sheet
        self.objects = objects
        self.scenarios = scenarios
        self.formatCells = formatCells
        self.formatColumns = formatColumns
        self.formatRows = formatRows
        self.insertColumns = insertColumns
        self.insertRows = insertRows
        self.insertHyperlinks = insertHyperlinks
        self.deleteColumns = deleteColumns
        self.deleteRows = deleteRows
        self.selectLockedCells = selectLockedCells
        self.selectUnlockedCells = selectUnlockedCells
        self.sort = sort
        self.autoFilter = autoFilter
        self.pivotTables = pivotTables
        if password is not None:
            self.set_password(password)
        self.algorithmName = algorithmName
        self.saltValue = saltValue
        self.spinCount = spinCount


    def set_password(self, value='', already_hashed=False):
        """Set a password on this sheet."""
        if not already_hashed:
            value = hash_password(value)
        self._password = value
        self.enable()

    @property
    def password(self):
        """Return the password value, regardless of hash."""
        return self._password

    @password.setter
    def _set_raw_password(self, value):
        """Set a password directly, forcing a hash step."""
        self.set_password(value, already_hashed=False)


    def enable(self):
        self.sheet = True


    def disable(self):
        self.sheet = False


    def __iter__(self):
        for key in ('sheet', 'objects', 'scenarios', 'formatCells',
                  'formatRows', 'formatColumns', 'insertColumns', 'insertRows',
                  'insertHyperlinks', 'deleteColumns', 'deleteRows',
                  'selectLockedCells', 'selectUnlockedCells', 'sort', 'autoFilter',
                  'pivotTables', 'password', 'algorithmName', 'saltValue', 'spinCount'):
            value = getattr(self, key)
            if value is not None:
                yield key, safe_string(value)
