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

"""Implements the lxml.etree.xmlfile API using the standard library xml.etree"""

from xml.etree import ElementTree
from xml.etree.ElementTree import Element
from contextlib import contextmanager  

class _IncrementalFileWriter(object):
    def __init__(self, output_file):
        self._element_stack = []
        self._top_element = None
        self._file = output_file
    
    @contextmanager
    def element(self, tag, attrib=None, nsmap=None, **_extra):
        """Create a new xml element using a context manager.
        The elements are written when the top level context is left."""
        
        # __enter__ part
        self._top_element = Element(tag)
        self._top_element.text = ''
        self._element_stack.append(self._top_element)
        
        if attrib is not None:
            self._top_element.attrib = attrib
            
        yield
        
        # __exit__ part
        self._top_element = self._element_stack.pop()        
        if self._element_stack:     
            self._element_stack[-1].append(self._top_element)
            self._top_element = self._element_stack[-1]
        else:
            self._file.write(ElementTree.tostring(self._top_element))
        
    def write(self, arg):
        """Write a string or subelement."""       
        if isinstance(arg, str):
            self._top_element.text += arg
        elif isinstance(arg, Element):
            self._top_element.append(arg)
        else:
            raise RuntimeError()
        
    def __enter__(self):
        pass
    def __exit__(self, type, value, traceback):
        pass
    
class xmlfile(object):
    def __init__(self, output_file, buffered=False):
        self._file = output_file
    def __enter__(self):
        return _IncrementalFileWriter(self._file)
        pass
    def __exit__(self, type, value, traceback):
        pass