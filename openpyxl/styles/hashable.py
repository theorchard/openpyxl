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

BASE_TYPES = (str, unicode, float, int)


class HashableObject(object):
    """Define how to hash property classes."""
    __fields__ = None
    __check__ = {}

    def _typecheck(self, name, value):
        expected = self.__check__.get(name)
        if expected:
            if not isinstance(value, expected):
                msg = '%s should be a %s, not %s' % (name, expected,
                                                     value.__class__.__name__)
                raise TypeError(msg)
        else:
            if value is not None and not isinstance(value, BASE_TYPES):
                raise TypeError('%s cannot be a %s' % (name,
                                                       value.__class__.__name__))

    def copy(self, **kwargs):
        current = dict([(x, getattr(self, x)) for x in self.__fields__])
        current.update(kwargs)
        return self.__class__(**current)

    def __setattr__(self, *args, **kwargs):
        name, value = args
        self._typecheck(name, value)
        if hasattr(self, name) and getattr(self, name) is not None:
            raise TypeError('cannot set %s attribute' % name)
        return object.__setattr__(self, *args, **kwargs)

    def __delattr__(self, *args, **kwargs):
        raise TypeError('cannot delete %s attribute' % args[0])

    def __repr__(self):
        return '%s(%s)' % (self.__class__.__name__,
                           ', '.join(['%s=%s' % (k, repr(getattr(self, k)))
                                        for k in self.__fields__]))

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
