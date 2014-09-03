from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Serialise external links"""


from openpyxl.xml.constants import SHEET_MAIN_NS, PKG_REL_NS
from openpyxl.xml.functions import Element, SubElement


def write_external_link(book, links):
    """Serialise links to ranges in a single external worbook"""
    tree = Element("{%s}externalLink" % SHEET_MAIN_NS)
    tree.append(Element("{%s}externalBook", {'{%s}id' % PKG_REL_NS:'rId1'}))
    external_ranges = SubElement(tree, "definedNames")
    for l in links:
        external_ranges.append(Element("{%s}definedName"), dict(l))

