from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Serialise external links"""


from openpyxl.xml.constants import SHEET_MAIN_NS, REL_NS, PKG_REL_NS
from openpyxl.xml.functions import Element, SubElement

from openpyxl.writer.workbook import RelationElement


def write_external_link(links):
    """Serialise links to ranges in a single external worbook"""
    root = Element("{%s}externalLink" % SHEET_MAIN_NS)
    book =  SubElement(root, "{%s}externalBook" % SHEET_MAIN_NS, {'{%s}id' % REL_NS:'rId1'})
    external_ranges = SubElement(book, "{%s}definedNames" % SHEET_MAIN_NS)
    for l in links:
        external_ranges.append(Element("{%s}definedName" % SHEET_MAIN_NS, dict(l)))
    return root


def write_external_book_rel(book):
    """Serialise link to external file"""
    root = Element("{%s}Relationships" % PKG_REL_NS)
    attrs = {"Id":"rId1", "Target":book.Target, "TargetMode":book.TargetMode,
             "Type":book.Type}
    root.append(RelationElement(attrs))
    return root
