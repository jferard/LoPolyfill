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
from typing import Any, Sequence

import unohelper
from com.github.jferard.lopolyfill import XLoPolyfill
from com.sun.star.lang import IllegalArgumentException

from lopolyfill_funcs import (
    LopFilter, LopRandarray, LopSequence, LopSort, LopUnique, LopXMatch)


class LoPolyfillImpl(unohelper.Base, XLoPolyfill):
    def __init__(self, ctx):
        self.ctxt = ctx

    # FILTER https://help.libreoffice.org/master/en-US/text/scalc/01/func_filter.html
    def lopFilter(
            self, inRange: Sequence[Sequence[Any]],
            criteria: Sequence[Sequence[Any]], defaultValue: Any):
        return LopFilter(IllegalArgumentException).execute(
            inRange, criteria, defaultValue)

    # RANDARRAY https://help.libreoffice.org/master/en-US/text/scalc/01/func_randarray.html
    def lopRandarray(self, rows: Any, columns: Any, minValue: Any,
                     maxValue: Any, integers: Any):
        return LopRandarray(IllegalArgumentException).execute(
            rows, columns, minValue, maxValue, integers)

    # SEQUENCE https://help.libreoffice.org/25.8/en-US/text/scalc/01/func_sequence.html
    def lopSequence(
            self, rows: int, columns: int, start: Any, step: Any):
        return LopSequence(IllegalArgumentException).execute(
            rows, columns, start, step)

    def lopSort(
            self, inRange: Sequence[Sequence[Any]], sortIndex: Any,
            sortOrder: Any, byCol: Any):
        oCollator = self._get_collator_from_doc()
        return LopSort(oCollator, IllegalArgumentException).sort(
            inRange, sortIndex, sortOrder, byCol)

    def _get_collator_from_doc(self):
        oDesktop = self.ctxt.getByName(
            "/singletons/com.sun.star.frame.theDesktop")
        oDoc = oDesktop.CurrentComponent
        ignore_case = oDoc.CharLocale
        oLocale = oDoc.CharLocale

        oServiceManager = self.ctxt.ServiceManager
        oCollator = oServiceManager.createInstance(
            "com.sun.star.i18n.Collator")
        oCollator.loadDefaultCollator(oLocale, 1 if ignore_case else 0)
        return oCollator

    def lopSortBy(
            self, inRange: Sequence[Sequence[Any]],
            sortByRange1: Sequence[Sequence[Any]], sortOrder1: int,
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
    ):
        oCollator = self._get_collator_from_doc()
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
            self, inRange: Sequence[Sequence[Any]], byCol: Any, uniqueness: Any
    ):
        return LopUnique(IllegalArgumentException).execute(
            inRange, byCol, uniqueness)

    def lopXLookup(
            self, criterion: Any,
            searchRange: Sequence[Sequence[Any]],
            resultRange: Sequence[Sequence[Any]],
            defaultValue: Any,
            matchMode: Any,
            searchMode: Any
    ):
        oCollator = self._get_collator_from_doc()
        return LopXMatch(oCollator, IllegalArgumentException).lookup(
            criterion, searchRange, resultRange,defaultValue, matchMode, searchMode)

    def lopXMatch(
            self, criterion: Any,
            searchRange: Sequence[Sequence[Any]],
            matchMode: Any,
            searchMode: Any
    ):
        oCollator = self._get_collator_from_doc()
        return LopXMatch(oCollator, IllegalArgumentException).match(
            criterion, searchRange, matchMode, searchMode)


def create_instance(ctx):
    logging.basicConfig(filename=str(Path.home() / "lopolyfill.log"),
                        encoding='utf-8',
                        level=logging.DEBUG, filemode="w")
    return LoPolyfillImpl(ctx)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    create_instance, "com.github.jferard.lopolyfill.LoPolyfillImpl",
    ("com.sun.star.sheet.AddIn",),
)
