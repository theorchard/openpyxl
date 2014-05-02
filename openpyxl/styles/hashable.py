from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

import inspect
from openpyxl.compat import unicode, basestring
from openpyxl.descriptors import Descriptor, Strict

BASE_TYPES = (str, unicode, float, int)


class HashableObject(Strict):
    """Define how to hash property classes."""
    __fields__ = ()
    __slots__ = __fields__
    __base__ = False

    @property
    def __defaults__(self):
        spec = inspect.getargspec(self.__class__.__init__)
        return dict(zip(spec.args[1:], spec.defaults))

    def copy(self, **kwargs):
        current = dict([(x, getattr(self, x)) for x in self.__fields__])
        current.update(kwargs)
        return self.__class__(**current)

    def __setattr__(self, *args, **kwargs):
        name, value = args
        if hasattr(self, name) and (
            getattr(self, name) is not None
            and not isinstance(getattr(self, name), Descriptor)):
            raise TypeError('cannot set %s attribute' % name)
        return object.__setattr__(self, *args, **kwargs)

    def __delattr__(self, *args, **kwargs):
        raise TypeError('cannot delete %s attribute' % args[0])

    def __print__(self, defaults=False):
        if defaults:
            print_func = str
        else:
            print_func = repr
        pieces = []
        default_values = self.__defaults__
        for k in self.__fields__:
            value = getattr(self, k)
            if not defaults and value == default_values[k]:
                continue
            if isinstance(value, basestring):
                print_func = repr  # keep quotes around strings
            pieces.append('%s=%s' % (k, print_func(value)))
        if pieces or self.__base__:
            return '%s(%s)' % (self.__class__.__name__, ', '.join(pieces))
        else:
            return ''

    def __repr__(self):
        return self.__print__(defaults=False)

    def __str__(self):
        return self.__print__(defaults=True)

    @property
    def __key(self):
        """Use a tuple of fields as the basis for a key"""
        return [getattr(self, x) for x in self.__fields__]

    def __hash__(self):
        return hash(str(self.__key))

    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return self.__key == other.__key
        return self.__key == other

    def __ne__(self, other):
        return not self == other
