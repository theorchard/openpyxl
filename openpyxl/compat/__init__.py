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
from tempfile import NamedTemporaryFile

from .strings import (
    basestring,
    unicode,
    bytes,
    file,
    tempfile,
    safe_string
    )
from .numbers import long, NUMERIC_TYPES
from .itertools import (
    range,
    ifilter,
    iteritems,
    iterkeys,
    itervalues,
    izip
)
try:
    from functools import lru_cache
except ImportError:
    from .functools import lru_cache

# Python 2.6
try:
    from collections import OrderedDict
except ImportError:
    from .odict import OrderedDict

import warnings
from functools import wraps
import inspect


class DummyCode:

    pass


class deprecated(object):

    def __init__(self, reason):
        if inspect.isclass(reason) or inspect.isfunction(reason):
            raise TypeError("Reason for deprecation must be supplied")
        self.reason = reason

    def __call__(self, obj, *args, **kwargs):
        @wraps(obj)
        def new_func(*args, **kwargs):
            msg = "Call to deprecated function or class {0} ({1})".format(obj.__name__,
                                                               self.reason)
            if inspect.isfunction(obj):
                _code = self._wrap_function(obj)
            elif inspect.isclass(obj):
                _code = self._wrap_class(obj)

            warnings.warn_explicit(
                '{0}.'.format(msg),
                category=UserWarning,
                filename=_code.co_filename,
                lineno=_code.co_firstlineno + 1
            )
            return obj(*args, **kwargs)
        return new_func

    def _wrap_function(self, obj):
        if hasattr(obj, 'func_code'):
            _code = obj.func_code
        else:
            _code = obj.__code__
        return _code

    def _wrap_class(self, obj):
        _code = DummyCode()
        _code.co_filename = obj.__module__
        _code.co_firstlineno = 0
        return _code
