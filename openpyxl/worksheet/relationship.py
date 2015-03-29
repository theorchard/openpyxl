from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors import String, Set, NoneSet
from openpyxl.descriptors.serialisable import Serialisable

from openpyxl.xml.constants import REL_NS, PKG_REL_NS
from openpyxl.xml.functions import Element, SubElement, tostring


class Relationship(Serialisable):
    """Represents many kinds of relationships."""
    # TODO: Use this object for workbook relationships as well as
    # worksheet relationships

    type = String()
    target = String(allow_none=True)
    targetMode = NoneSet(values=("external"))
    id = String(allow_none=True)


    def __init__(self, type, target=None, targetMode=None, id=None):
        self.type = "%s/%s" % (REL_NS, type)
        self.target = target
        self.target_mode = targetMode
        self.id = id

    def __repr__(self):
        root = Element("{%s}Relationships" % PKG_REL_NS)
        SubElement(root, "{%s}Relationship" % PKG_REL_NS, self.__dict__)
        return tostring(root).decode("utf-8")

