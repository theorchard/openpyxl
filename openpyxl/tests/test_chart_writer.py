import pytest
import os

from openpyxl.xml.functions import Element, fromstring, safe_iterator
from openpyxl.xml.constants import CHART_NS

from openpyxl.writer.charts import (ChartWriter,
                                    PieChartWriter,
                                    LineChartWriter,
                                    BarChartWriter,
                                    ScatterChartWriter,
                                    BaseChartWriter
                                    )
from openpyxl.styles import Color
