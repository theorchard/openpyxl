from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
    Set,
    Float,
)
from openpyxl.descriptors.excel import ExtensionList


class DLbl(Serialisable):

    idx = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 idx=None,
                 extLst=None,
                ):
        self.idx = idx
        self.extLst = extLst


class DLbls(Serialisable):

    dLbl = Typed(expected_type=DLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 dLbl=None,
                 extLst=None,
                ):
        self.dLbl = dLbl
        self.extLst = extLst
