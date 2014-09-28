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

#from openpyxl.xml.functions import Element, tostring
from xml.etree.ElementTree import Element, tostring
from contextlib import contextmanager  

class LxmlSyntaxError(Exception):
    pass

class _FakeIncrementalFileWriter(object):
    """Replacement for _IncrementalFileWriter of lxml.
       Uses ElementTree to build xml in memory."""
    def __init__(self, output_file):
        self._element_stack = []
        self._top_element = None
        self._file = output_file
        self._have_root = False
    
    @contextmanager
    def element(self, tag, attrib=None, nsmap=None, **_extra):
        """Create a new xml element using a context manager.
        The elements are written when the top level context is left."""
        
        # __enter__ part
        self._have_root = True
        self._top_element = Element(tag)
        self._top_element.text = ''
        self._top_element.tail = ''
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
            self._write_element(self._top_element)
            self._top_element = None
        
    def write(self, arg):
        """Write a string or subelement."""
                       
        if isinstance(arg, str):
            # it is not allowed to write a string outside of an element
            if self._top_element is None:
                raise LxmlSyntaxError() 
                                      
            if len(self._top_element) == 0:
                # element has no children: add string to text
                self._top_element.text += arg
            else:
                # element has children: add string to tail of last child
                self._top_element[-1].tail += arg
                
        else:
            if self._top_element is not None:
                self._top_element.append(arg)
            elif not self._have_root:
                self._write_element(arg)
            else:
                raise LxmlSyntaxError()
        
    def _write_element(self, element):
        xml = tostring(element)
        self._file.write(xml)
        
    def __enter__(self):
        pass
    def __exit__(self, type, value, traceback):
        # without root the xml document is incomplete
        if not self._have_root:
            raise LxmlSyntaxError()
    
class xmlfile(object):
    """Context manager that can replace lxml.etree.xmlfile."""
    def __init__(self, output_file, buffered=False, encoding=None, close=False):        
        if isinstance(output_file, str):
            self._file = open(output_file, 'w')
            self._close = True
        else:
            self._file = output_file
            self._close = close            
            
    def __enter__(self):
        return _FakeIncrementalFileWriter(self._file)

    def __exit__(self, type, value, traceback):
        if self._close == True:
            self._file.close()