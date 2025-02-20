# LoPolyfill - Python - A set of 24.8 functions made availaible for 7.2
# Copyright (C) 2025 Julien Férard.
#
# LoPolyfill is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# LoPolyfill is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# IMPORTANT: The documentation of the provided functions and their parameters is
# taken from the LibreOffice help pages ( Mozilla Public License v2.0).
import logging
from pathlib import Path
from typing import Any, Sequence, Tuple, List

import unohelper
from com.github.jferard.lopolyfill import XLoPolyfill
from com.sun.star.beans import XPropertySet
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.i18n import XCollator

from lopolyfill_funcs import (
    LopFilter, LopRandarray, LopSequence, LopSort, LopUnique, LopXMatch,
    LopArrayHandling, DataArray, DataRow)


class LoPolyfillImpl(unohelper.Base, XLoPolyfill):
    def __init__(self, ctx):
        self.ctxt = ctx

    # FILTER https://help.libreoffice.org/master/en-US/text/scalc/01/func_filter.html
    def lopFilter(
            self, inRange: DataArray,
            criteria: DataArray, defaultValue: Any
    ) -> DataArray:
        return LopFilter(IllegalArgumentException).execute(
            inRange, criteria, defaultValue)

    # RANDARRAY https://help.libreoffice.org/master/en-US/text/scalc/01/func_randarray.html
    def lopRandarray(self, rows: Any, columns: Any, minValue: Any,
                     maxValue: Any, integers: Any
                     ) -> DataArray:
        return LopRandarray(IllegalArgumentException).execute(
            rows, columns, minValue, maxValue, integers)

    # SEQUENCE https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_sequence.html
    def lopSequence(
            self, rows: int, columns: int, start: Any, step: Any
    ) -> DataArray:
        return LopSequence(IllegalArgumentException).execute(
            rows, columns, start, step)

    def lopSort(
            self,
            oDoc: XPropertySet,
            inRange: DataArray,
            sortIndex: Any, sortOrder: Any, byCol: Any
    ) -> DataArray:
        oCollator = self._get_collator_from_doc(oDoc)
        return LopSort(oCollator, IllegalArgumentException).sort(
            inRange, sortIndex, sortOrder, byCol)

    def lopSortBy(
            self,
            oDoc: XPropertySet,
            inRange: DataArray,
            sortByRange1: DataArray, sortOrder1: int,
            sortByRange2: Any, sortOrder2: Any,
            sortByRange3: Any, sortOrder3: Any,
            sortByRange4: Any, sortOrder4: Any,
            sortByRange5: Any, sortOrder5: Any,
            sortByRange6: Any, sortOrder6: Any,
            sortByRange7: Any, sortOrder7: Any,
            sortByRange8: Any, sortOrder8: Any,
            sortByRange9: Any, sortOrder9: Any,
            sortByRange10: Any, sortOrder10: Any,
            sortByRange11: Any, sortOrder11: Any,
            sortByRange12: Any, sortOrder12: Any,
            sortByRange13: Any, sortOrder13: Any,
            sortByRange14: Any, sortOrder14: Any,
            sortByRange15: Any, sortOrder15: Any,
    ) -> DataArray:
        oCollator = self._get_collator_from_doc(oDoc)
        return LopSort(oCollator, IllegalArgumentException).sort_by(
            inRange,
            sortByRange1, sortOrder1, sortByRange2, sortOrder2,
            sortByRange3, sortOrder3, sortByRange4, sortOrder4,
            sortByRange5, sortOrder5, sortByRange6, sortOrder6,
            sortByRange7, sortOrder7, sortByRange8, sortOrder8,
            sortByRange9, sortOrder9, sortByRange10, sortOrder10,
            sortByRange11, sortOrder11, sortByRange12, sortOrder12,
            sortByRange13, sortOrder13, sortByRange14, sortOrder14,
            sortByRange15, sortOrder15
        )

    def lopUnique(
            self, inRange: DataArray, byCol: Any, uniqueness: Any
    ) -> DataArray:
        return LopUnique(IllegalArgumentException).execute(
            inRange, byCol, uniqueness)

    def lopXLookup(
            self,
            oDoc: XPropertySet,
            criterion: Any,
            searchRange: DataArray,
            resultRange: DataArray,
            defaultValue: Any,
            matchMode: Any,
            searchMode: Any
    ) -> DataArray:
        oCollator = self._get_collator_from_doc(oDoc)
        return LopXMatch(oCollator, IllegalArgumentException).lookup(
            criterion, searchRange, resultRange, defaultValue, matchMode,
            searchMode)

    def lopXMatch(
            self,
            oDoc: XPropertySet,
            criterion: Any,
            searchRange: DataArray,
            matchMode: Any,
            searchMode: Any
    ):
        oCollator = self._get_collator_from_doc(oDoc)
        return LopXMatch(oCollator, IllegalArgumentException).match(
            criterion, searchRange, matchMode, searchMode)

    def lopChooseCols(
            self, array: DataArray, column1: int,
            column2: Any, column3: Any, column4: Any, column5: Any,
            column6: Any, column7: Any, column8: Any, column9: Any,
            column10: Any, column11: Any, column12: Any, column13: Any,
            column14: Any, column15: Any, column16: Any, column17: Any,
            column18: Any, column19: Any, column20: Any, column21: Any,
            column22: Any, column23: Any, column24: Any, column25: Any,
            column26: Any, column27: Any, column28: Any, column29: Any,
            column30: Any,
    ) -> List[List[Any]]:
        return LopArrayHandling(IllegalArgumentException).choose_cols(
            array, column1,
            column2, column3, column4, column5,
            column6, column7, column8, column9,
            column10, column11, column12, column13,
            column14, column15, column16, column17,
            column18, column19, column20, column21,
            column22, column23, column24, column25,
            column26, column27, column28, column29,
            column30)

    def lopChooseRows(
            self, array: DataArray, row1: int,
            row2: Any, row3: Any, row4: Any, row5: Any,
            row6: Any, row7: Any, row8: Any, row9: Any,
            row10: Any, row11: Any, row12: Any, row13: Any,
            row14: Any, row15: Any, row16: Any, row17: Any,
            row18: Any, row19: Any, row20: Any, row21: Any,
            row22: Any, row23: Any, row24: Any, row25: Any,
            row26: Any, row27: Any, row28: Any, row29: Any,
            row30: Any,
    ) -> List[List[Any]]:
        return LopArrayHandling(IllegalArgumentException).choose_rows(
            array, row1,
            row2, row3, row4, row5,
            row6, row7, row8, row9,
            row10, row11, row12, row13,
            row14, row15, row16, row17,
            row18, row19, row20, row21,
            row22, row23, row24, row25,
            row26, row27, row28, row29,
            row30)

    def lopDrop(
            self, array: DataArray, rows: Any,
            columns: Any,
    ) -> List[DataRow]:
        return LopArrayHandling(IllegalArgumentException).drop(
            array, rows, columns)

    def lopTake(
            self, array: DataArray, rows: Any,
            columns: Any,
    ) -> List[DataRow]:
        return LopArrayHandling(IllegalArgumentException).take(
            array, rows, columns)

    def lopExpand(
            self, array: DataArray, rows: Any,
            columns: Any, pad_with: Any
    ) -> List[DataRow]:
        return LopArrayHandling(IllegalArgumentException).expand(
            array, rows, columns, pad_with)

    def lopHStack(
            self, array: DataArray,
            array1: Any, array2: Any, array3: Any, array4: Any,
            array5: Any, array6: Any, array7: Any, array8: Any,
            array9: Any, array10: Any, array11: Any, array12: Any,
            array13: Any, array14: Any, array15: Any, array16: Any,
            array17: Any, array18: Any, array19: Any, array20: Any,
            array21: Any, array22: Any, array23: Any, array24: Any,
            array25: Any, array26: Any, array27: Any, array28: Any,
            array29: Any, array30: Any,
    ) -> List[List[Any]]:
        return LopArrayHandling(IllegalArgumentException).hstack(
            array, array1, array2, array3, array4, array5, array6, array7,
            array8, array9, array10, array11, array12, array13, array14,
            array15, array16, array17, array18, array19, array20, array21,
            array22, array23, array24, array25, array26, array27, array28,
            array29, array30,
        )

    def lopVStack(
            self, array: DataArray,
            array1: Any, array2: Any, array3: Any, array4: Any,
            array5: Any, array6: Any, array7: Any, array8: Any,
            array9: Any, array10: Any, array11: Any, array12: Any,
            array13: Any, array14: Any, array15: Any, array16: Any,
            array17: Any, array18: Any, array19: Any, array20: Any,
            array21: Any, array22: Any, array23: Any, array24: Any,
            array25: Any, array26: Any, array27: Any, array28: Any,
            array29: Any, array30: Any,
    ) -> List[DataRow]:
        return LopArrayHandling(IllegalArgumentException).vstack(
            array, array1, array2, array3, array4, array5, array6, array7,
            array8, array9, array10, array11, array12, array13, array14,
            array15, array16, array17, array18, array19, array20, array21,
            array22, array23, array24, array25, array26, array27, array28,
            array29, array30,
        )

    def lopToCol(
            self, array: DataArray, ignore: Any,
            by_column: Any
    ) -> List[DataRow]:
        return LopArrayHandling(IllegalArgumentException).to_col(
            array, ignore, by_column)

    def lopToRow(
            self, array: DataArray, ignore: Any,
            by_column: Any
    ) -> List[List[Any]]:
        return LopArrayHandling(IllegalArgumentException).to_row(
            array, ignore, by_column)

    def lopWrapCols(
            self, in_range: DataArray, wrap_count: int,
            pad_with: Any
    ) -> List[List[Any]]:
        return LopArrayHandling(IllegalArgumentException).wraps_cols(
            in_range, wrap_count, pad_with)

    def lopWrapRows(
            self, in_range: DataArray, wrap_count: int,
            pad_with: Any
    ) -> List[DataRow]:
        return LopArrayHandling(IllegalArgumentException).wraps_rows(
            in_range, wrap_count, pad_with)

    def _get_collator_from_doc(
            self, oDoc: XPropertySet, ignore_case: bool = True) -> XCollator:
        oServiceManager = self.ctxt.ServiceManager
        oCollator = oServiceManager.createInstance(
            "com.sun.star.i18n.Collator")

        oLocale = oDoc.CharLocale
        oCollator.loadDefaultCollator(oLocale, 1 if ignore_case else 0)

        return oCollator


def create_instance(ctxt):
    logging.basicConfig(filename=str(Path.home() / "lopolyfill.log"),
                        encoding='utf-8',
                        level=logging.DEBUG, filemode="w")
    logging.getLogger(__name__).debug("Instance creation")
    ret = LoPolyfillImpl(ctxt)
    return ret


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    create_instance, "com.github.jferard.lopolyfill.LoPolyfillImpl",
    ("com.sun.star.sheet.AddIn",),
)


class SimpleCollator:
    @staticmethod
    def compareString(s1: str, s2: str) -> int:
        s1 = s1.casefold()
        s2 = s2.casefold()
        if s1 < s2:
            return -1
        elif s1 > s2:
            return 1
        else:
            return 0
