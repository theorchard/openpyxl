from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from .graph import GraphChart


class BarChart(GraphChart):

    TYPE = "barChart"
    GROUPING = "clustered"
