from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import datetime

from openpyxl.date_time import CALENDAR_WINDOWS_1900


class DocumentProperties(object):
    """High-level properties of the document."""

    def __init__(self):
        self.creator = 'Unknown'
        self.last_modified_by = self.creator
        self.modified = self.created = datetime.datetime.now()
        self.title = 'Untitled'
        self.subject = ''
        self.description = ''
        self.keywords = ''
        self.category = ''
        self.company = 'Microsoft Corporation'
        self.excel_base_date = CALENDAR_WINDOWS_1900



class DocumentSecurity(object):
    """Security information about the document."""

    def __init__(self):
        self.lock_revision = False
        self.lock_structure = False
        self.lock_windows = False
        self.revision_password = ''
        self.workbook_password = ''
