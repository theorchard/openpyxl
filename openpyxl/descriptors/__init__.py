from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from .base import *


class MetaStrict(type):

    def __new__(cls, clsname, bases, methods):
        for k, v in methods.items():
            if isinstance(v, Descriptor):
                v.name = k
        return type.__new__(cls, clsname, bases, methods)

Strict = MetaStrict('Strict', (object,), {}
               )

del MetaStrict
