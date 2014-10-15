from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import datetime

from openpyxl.compat import safe_string
from openpyxl.date_time import CALENDAR_WINDOWS_1900
from openpyxl.descriptors import Strict, String, Typed, Sequence, Alias


class DocumentProperties(Strict):
    """High-level properties of the document.
    Defined in ECMA-376 Par2 Annex D
    """

    category = String(allow_none=True)
    contentStatus = String(allow_none=True)
    _keywords = Sequence(allow_none=True)
    lastModifiedBy = String(allow_none=True)
    lastPrinted = String(allow_none=True)
    revision = String(allow_none=True)
    version = String(allow_none=True)
    last_modified_by = Alias("lastModifiedBy")

    # Dublin Core Properties
    created = Typed(expected_type=datetime.datetime, allow_none=True)
    creator = String(allow_none=True)
    description = String(allow_none=True)
    identifier = String(allow_none=True)
    language = String(allow_none=True)
    modified = Typed(expected_type=datetime.datetime, allow_none=True)
    subject = String(allow_none=True)
    title = String(allow_none=True)

    __fields__ = ("category", "contentStatus", "keywords", "lastModifiedBy",
                "lastPrinted", "revision", "version", "created", "creator", "description",
                "identifier", "language", "modified", "subject", "title")

    def __init__(self,
                 category=None,
                 contentStatus=None,
                 keywords=[],
                 lastModifiedBy=None,
                 lastPrinted=None,
                 revision=None,
                 version=None,
                 created=datetime.datetime.now(),
                 creator="openpyxl",
                 description=None,
                 identifier=None,
                 language=None,
                 modified=datetime.datetime.now(),
                 subject=None,
                 title=None,
                 ):
        self.contentStatus = contentStatus
        self.lastPrinted = lastPrinted
        self.revision = revision
        self.version = version
        self.creator = creator
        self.lastModifiedBy = lastModifiedBy
        self.creator = creator
        self.modified = modified
        self.created = created
        self.title = title
        self.subject = subject
        self.description = description
        self.identifier = identifier
        self.language = language
        self._keywords = keywords
        self.category = category
        self.excel_base_date = CALENDAR_WINDOWS_1900


    @property
    def keywords(self):
        """Return keywords as string or None if emtpy
        """
        if self._keywords:
            return ", ".join([self._keywords])


    def __iter__(self):
        for attr in self.__fields__:
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


class DocumentSecurity(object):
    """Security information about the document."""

    def __init__(self):
        self.lock_revision = False
        self.lock_structure = False
        self.lock_windows = False
        self.revision_password = ''
        self.workbook_password = ''
