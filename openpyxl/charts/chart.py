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

from openpyxl.drawing import Drawing, Shape

from .legend import Legend
from .series import Series


class Chart(object):
    """ raw chart class """

    GROUPING = 'standard'
    TYPE = None

    def mymax(self, values):
        return max([x for x in values if x is not None])

    def mymin(self, values):
        return min([x for x in values if x is not None])

    def __init__(self):

        self.series = []
        self._series = self.series # backwards compatible

        # public api
        self.legend = Legend()
        self.show_legend = True
        self.lang = 'en-GB'
        self.title = ''
        self.print_margins = dict(b=.75, l=.7, r=.7, t=.75, header=0.3, footer=.3)

        # the containing drawing
        self.drawing = Drawing()
        self.drawing.left = 10
        self.drawing.top = 400
        self.drawing.height = 400
        self.drawing.width = 800

        # the offset for the plot part in percentage of the drawing size
        self.width = .6
        self.height = .6
        self._margin_top = 1
        self._margin_top = self.margin_top
        self._margin_left = 0

        # the user defined shapes
        self.shapes = []
        self._shapes = self.shapes # backwards compatible

    def append(self, obj):
        """Add a series or a shape"""
        if isinstance(obj, Series):
            self.series.append(obj)
        elif isinstance(obj, Shape):
            self.shapes.append(obj)

    add_shape = add_serie = add_series = append

    def __iter__(self):
        return iter(self.series)

    def get_y_chars(self):
        """ estimate nb of chars for y axis """
        _max = max([s.max() for s in self])
        return len(str(int(_max)))

    @property
    def margin_top(self):
        """ get margin in percent """

        return min(self._margin_top, self._get_max_margin_top())

    @margin_top.setter
    def margin_top(self, value):
        """ set base top margin"""
        self._margin_top = value

    def _get_max_margin_top(self):

        mb = Shape.FONT_HEIGHT + Shape.MARGIN_BOTTOM
        plot_height = self.drawing.height * self.height
        return float(self.drawing.height - plot_height - mb) / self.drawing.height

    @property
    def margin_left(self):

        return max(self._get_min_margin_left(), self._margin_left)

    @margin_left.setter
    def margin_left(self, value):
        self._margin_left = value

    def _get_min_margin_left(self):

        ml = (self.get_y_chars() * Shape.FONT_WIDTH) + Shape.MARGIN_LEFT
        return float(ml) / self.drawing.width
