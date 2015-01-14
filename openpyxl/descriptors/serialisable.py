from __future__ import absolute_import
# copyright openpyxl 2010-2015

from . import _Serialiasable

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement, safe_iterator, localname


class Serialisable(_Serialiasable):
    """
    Objects can serialise to XML their attributes and child objects.
    The following class attributes are created by the metaclass at runtime:
    __attrs__ = attributes
    __nested__ = single-valued child treated as an attribute
    __elements__ = child elements
    """

    __attrs__ = None
    __nested__ = None
    __elements__ = None

    @property
    def tagname(self):
        raise(NotImplementedError)


    @classmethod
    def create(cls, node):
        """
        Create object from XML
        """
        attrib = dict(node.attrib)
        for el in node:
            tag = localname(el)
            desc = getattr(cls, tag, None)
            if desc is None:
                continue
            if tag in cls.__nested__:
                attrib[tag] = cls._create_nested(el, tag)
            else:
                attrib[tag] = desc.expected_type.create(el)
        return cls(**attrib)


    @classmethod
    def _create_nested(cls, el, tag):
        """
        Allow special handling of nested attributes in subclasses.
        Default for child elements without a val attribute is True
        """
        return el.get("val", True)


    def serialise(self, tagname=None):
        if tagname is None:
            tagname = self.tagname
        attrs = dict(self)
        el = Element(tagname, attrs)
        for n in self.__nested__:
            value = getattr(self, n)
            if isinstance(value, tuple):
                if hasattr(el, 'extend'):
                    el.extend(self._serialise_nested(value))
                else: # py26 nolxml
                    for _ in self._serialise_nested(value):
                        el.append(_)
            elif value:
                SubElement(el, n, val=safe_string(value))
        for c in self.__elements__:
            obj = getattr(self, c)
            if obj is not None:
                el.append(obj.serialise())
        return el


    def _serialise_nested(self, sequence):
        """
        Allow special handling of sequences which themselves are not directly serialisable
        """
        for obj in sequence:
            yield obj.serialise()


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)
