from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import datetime

from openpyxl.compat import safe_string
from openpyxl.date_time import CALENDAR_WINDOWS_1900, datetime_to_W3CDTF
from openpyxl.descriptors import Strict, String, Typed, Sequence, Alias
from openpyxl.xml.functions import ElementTree, Element, SubElement, tostring
from openpyxl.xml.constants import COREPROPS_NS, DCORE_NS, XSI_NS, DCTERMS_NS, DCTERMS_PREFIX


class DocumentProperties(Strict):
    """High-level properties of the document.
    Defined in ECMA-376 Par2 Annex D
    """

    category = String(allow_none=True)
    contentStatus = String(allow_none=True)
    keywords = Sequence(allow_none=True)
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

    __fields__ = ("category", "contentStatus", "lastModifiedBy",
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
        self.keywords = keywords
        self.category = category

    def __iter__(self):
        for attr in self.__fields__:
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


def write_properties(props):
    """Write the core properties to xml."""
    root = Element('{%s}coreProperties' % COREPROPS_NS)
    SubElement(root, '{%s}creator' % DCORE_NS).text = props.creator
    SubElement(root, '{%s}lastModifiedBy' % COREPROPS_NS).text = props.lastModifiedBy
    SubElement(root, '{%s}created' % DCTERMS_NS,
               {'{%s}type' % XSI_NS: '%s:W3CDTF' % DCTERMS_PREFIX}).text = \
                   datetime_to_W3CDTF(props.created)
    SubElement(root, '{%s}modified' % DCTERMS_NS,
               {'{%s}type' % XSI_NS: '%s:W3CDTF' % DCTERMS_PREFIX}).text = \
                   datetime_to_W3CDTF(props.modified)
    SubElement(root, '{%s}title' % DCORE_NS).text = props.title
    SubElement(root, '{%s}description' % DCORE_NS).text = props.description
    SubElement(root, '{%s}subject' % DCORE_NS).text = props.subject
    node = SubElement(root, '{%s}keywords' % COREPROPS_NS)
    for kw in props.keywords:
        SubElement(node, "{%s}keyword").text = kw
    SubElement(root, '{%s}category' % COREPROPS_NS).text = props.category
    return tostring(root)



class DocumentSecurity(object):
    """Security information about the document."""

    def __init__(self):
        self.lock_revision = False
        self.lock_structure = False
        self.lock_windows = False
        self.revision_password = ''
        self.workbook_password = ''
