from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


class Proxy(object):
    """
    Proxy formatting objects so that they cannot be altered
    """


    __slots__ = ('__target')

    def __init__(self, target):
        self.__target = target

    def __repr__(self):
        return repr(self.__target)

    def __getattr__(self, attr):
        return getattr(self.__target, attr)
