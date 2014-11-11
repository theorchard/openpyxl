from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


"""
Proxy formatting objects so that they cannot be altered
"""


class Proxy(object):

    __slots__ = ('__target')

    def __init__(self, target):
        self.__target = target


    def copy(self):
        pass
