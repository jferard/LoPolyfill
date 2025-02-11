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
# taken from the LibreOffice help pages (Mozilla Public License v2.0).
import collections
import enum
import functools
import itertools
import random
from typing import Sequence, Any, Callable, List, Tuple, Optional, Iterable


class LopFilter:
    def __init__(self, illegal_argument_exception: Any):
        self._illegal_argument_exception = illegal_argument_exception

    def execute(self, rows: Sequence[Sequence[Any]],
                criteria: Sequence[Sequence[Any]], default_value: Any):
        assert rows and rows[0]

        orientation = get_orientation(criteria, rows)

        if orientation == Orientation.BY_ROW:  # row filter
            ret = [
                row
                for c, row in zip(criteria, rows)
                if c[0]
            ]
        elif orientation == Orientation.BY_COL:  # col filter
            selected_columns = criteria[0]
            ret = [
                [v for s, v in zip(selected_columns, row) if s]
                for row in rows
            ]
            if not ret[0]:
                ret = []
        else:
            raise self._illegal_argument_exception("Bad criteria")

        if ret:
            return ret
        else:
            return [] if default_value is None else [[default_value]]


class LopRandarray:
    def __init__(self, illegal_argument_exception: Any):
        self._illegal_argument_exception = illegal_argument_exception

    def execute(
            self, row_count: Any, column_count: Any, min_value: Any,
            max_value: Any, integers: Any):
        if row_count is None:
            row_count = 1
        else:
            row_count = int(row_count)
            if row_count < 1:
                raise self._illegal_argument_exception("Number of row_count")

        if column_count is None:
            column_count = 1
        else:
            column_count = int(column_count)
            if column_count < 1:
                raise self._illegal_argument_exception("Number of column_count")

        if min_value is None:
            min_value = 0
        if max_value is None:
            max_value = 1
        if integers is None or integers:
            _minValue = int(min_value)
            _maxValue = int(max_value)

            def generate() -> int:
                return random.randint(_minValue, _maxValue)
        else:
            def generate() -> float:
                return min_value + random.random() * (max_value - min_value)

        return [
            [generate() for _ in range(column_count)]
            for _ in range(row_count)
        ]


class LopSequence:
    def __init__(self, illegal_argument_exception: Any):
        self._illegal_argument_exception = illegal_argument_exception

    def execute(
            self, rows: int, columns: int, start: Any, step: Any):
        if rows < 1:
            raise self._illegal_argument_exception("Number of row_count")
        rows = int(rows)

        if columns < 1:
            raise self._illegal_argument_exception("Number of column_count")
        columns = int(columns)

        if start is None:
            start = 1
        else:
            start = int(start)

        if step is None:
            step = 1
        else:
            step = int(step)

        values_iter = itertools.count(start, step)
        return [
            list(itertools.islice(values_iter, columns))
            for _ in range(rows)
        ]


class LopSort:
    def __init__(self, oCollator, illegal_argument_exception: Any):
        self._oCollator = oCollator
        self._illegal_argument_exception = illegal_argument_exception

    def sort(
            self, in_range: Sequence[Sequence[Any]], sort_index: Any,
            sort_order: Any, by_col: Any):
        # ascending: Number / String / Error
        # desc: Error / Number / String
        assert in_range and in_range[0]

        if sort_index is None:
            sort_index = 0
        else:
            sort_index = int(sort_index) - 1

        ascending = self._is_ascending(sort_order)

        if by_col:
            return self._by_col_lop_sort(in_range, sort_index, ascending)
        else:
            return self._by_row_lop_sort(in_range, sort_index, ascending)

    def _is_ascending(self, sortOrder: Any) -> bool:
        if sortOrder is None:
            return True

        sortOrder = int(sortOrder)
        if sortOrder == 1:
            ascending = True
        elif sortOrder == -1:
            ascending = False
        else:
            raise self._illegal_argument_exception("sort_order")
        return ascending

    def _by_col_lop_sort(
            self, inRange: Sequence[Sequence[Any]], sortIndex: int,
            ascending: bool):
        cols = list(zip(*inRange))
        sorted_cols = self._by_row_lop_sort(cols, sortIndex, ascending)
        return list(zip(*sorted_cols))

    def _by_row_lop_sort(
            self, rows: Sequence[Sequence[Any]], sort_index: int,
            ascending: bool):
        if sort_index < 0 or sort_index >= len(rows[0]):
            raise self._illegal_argument_exception("SortIndex col")

        cmp_values_with_collator = create_cmp_values_with_collator(
            self._oCollator)

        def cmp_rows(row1: Sequence[Any], row2: Sequence[Any]):
            return cmp_values_with_collator(row1[sort_index], row2[sort_index])

        return sorted(rows, key=functools.cmp_to_key(cmp_rows),
                      reverse=ascending is False)

    def sort_by(
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
        if not (inRange and inRange[0]):
            return inRange
        if not (sortByRange1 and sortByRange1[0]):
            return inRange

        h = len(inRange)
        w = len(inRange[0])
        byCol = self._is_by_col(sortByRange1, h, w)
        extract = self._create_extract(byCol, h, w)
        sortKeys = [
            (extract(sortByRange), self._is_ascending(sortOrder))
            for sortByRange, sortOrder in (
                (sortByRange1, sortOrder1),
                (sortByRange2, sortOrder2),
                (sortByRange3, sortOrder3),
                (sortByRange4, sortOrder4),
                (sortByRange5, sortOrder5),
                (sortByRange6, sortOrder6),
                (sortByRange7, sortOrder7),
                (sortByRange8, sortOrder8),
                (sortByRange9, sortOrder9),
                (sortByRange10, sortOrder10),
                (sortByRange11, sortOrder11),
                (sortByRange12, sortOrder12),
                (sortByRange13, sortOrder13),
                (sortByRange14, sortOrder14),
                (sortByRange15, sortOrder15),
            )
            if sortByRange is not None
        ]
        if byCol:
            columns = list(zip(*inRange))
            columns = self._by_row_lop_sortby(columns, sortKeys)
            return list(zip(*columns))
        else:
            return self._by_row_lop_sortby(inRange, sortKeys)

    def _is_by_col(
            self, sortByRange1: Sequence[Sequence[Any]], h: int, w: int
    ) -> bool:
        if len(sortByRange1) == h:
            byCol = False
        elif len(sortByRange1[0]) == w:
            byCol = True
        else:
            raise self._illegal_argument_exception("sortRange")
        return byCol

    def _create_extract(
            self, byCol: bool, h: int, w: int
    ) -> Callable[[Sequence[Sequence[Any]]], Sequence[Any]]:
        if byCol:
            def extract(sortByRange: Sequence[Sequence[Any]]) -> Sequence[Any]:
                if len(sortByRange) != 1 or len(sortByRange[0]) != w:
                    raise self._illegal_argument_exception("sortRange")
                return sortByRange[0]
        else:
            def extract(sortByRange: Sequence[Sequence[Any]]) -> Sequence[Any]:
                if len(sortByRange) != h or len(sortByRange[0]) != 1:
                    raise self._illegal_argument_exception("sortRange")
                return [row[0] for row in sortByRange]
        return extract

    def _by_row_lop_sortby(
            self, rows: Sequence[Sequence[Any]],
            sortKeys: List[Tuple[Any, bool]]
    ) -> Sequence[Sequence[Any]]:
        cmp_values_with_collator = create_cmp_values_with_collator(
            self._oCollator)

        def cmp_integers(i: int, j: int) -> int:
            for values, ascending in sortKeys:
                c = cmp_values_with_collator(values[i], values[j])
                if c != 0:
                    if ascending:
                        return c
                    else:
                        return -c
            return 0

        indices = range(len(rows))
        sorted_indices = sorted(indices, key=functools.cmp_to_key(cmp_integers))
        return [rows[i] for i in sorted_indices]


class LopUnique:
    def __init__(self, illegal_argument_exception: Any):
        self._illegal_argument_exception = illegal_argument_exception

    def execute(
            self, in_range: Sequence[Sequence[Any]], by_col: Any,
            uniqueness: Any
    ):
        if uniqueness is None:
            uniqueness = False
        else:
            uniqueness = bool(uniqueness)
        if by_col:
            cols = list(zip(*in_range))
            unique_cols = self._unique_by_row(cols, uniqueness)
            return list(zip(*unique_cols))
        else:
            return self._unique_by_row(in_range, uniqueness)

    def _unique_by_row(self, in_range, uniqueness):
        # TODO: collator?
        # class LoStr:
        # eq : collator.compareString == 0
        # hash = hash(unicodedata.normalize("NFD", egypt))

        rows = [tuple(row) for row in in_range]
        ret = []
        if uniqueness:
            c = collections.Counter(rows)
            for row in rows:
                if c[row] == 1:
                    ret.append(row)
                    c[row] = 0
        else:
            seen = set()
            for row in rows:
                if row not in seen:
                    seen.add(row)
                    ret.append(row)
        return ret


class XMatchMode(enum.IntEnum):
    EXACT = 0
    SMALLER = -1
    LARGER = 1
    WILDCARD = 2
    REGEX = 3


class XSearchMode(enum.IntEnum):
    FIRST = 1
    LAST = -1
    FIRST_BINARY = 2
    LAST_BINARY = -2


class LopXMatch:
    def __init__(self, oCollator, illegal_argument_exception: Any):
        self._oCollator = oCollator
        self._illegal_argument_exception = illegal_argument_exception

    def lookup(
            self, criterion: Any,
            search_range: Sequence[Sequence[Any]],
            result_range: Sequence[Sequence[Any]],
            default_value: Any,
            match_mode: Any,
            search_mode: Any
    ):
        assert result_range and result_range[0]

        orientation = get_orientation(search_range, result_range)
        if orientation == Orientation.BY_ROW:
            values = [row[0] for row in search_range]
        else:
            values = search_range[0]
        match_mode = self._get_match_mode(match_mode)
        search_mode = self._get_search_mode(search_mode)

        idx = self._match_value(criterion, values, match_mode, search_mode)
        if idx is None:
            return [[default_value]]
        elif orientation == Orientation.BY_ROW:
            return [result_range[idx]]
        else:
            return [[row[idx]] for row in result_range]

    def _get_match_mode(self, match_mode: int) -> XMatchMode:
        if match_mode is None:
            match_mode = XMatchMode.EXACT
        else:
            try:
                match_mode = XMatchMode(int(match_mode))
            except ValueError:
                raise self._illegal_argument_exception("MatchMode")
        return match_mode

    def _get_search_mode(self, search_mode: int) -> XSearchMode:
        if search_mode is None:
            search_mode = XSearchMode.FIRST
        else:
            try:
                search_mode = XSearchMode(int(search_mode))
            except ValueError:
                raise self._illegal_argument_exception("SearchModeMode")
        return search_mode

    def match(
            self, criterion: Any,
            search_range: Sequence[Sequence[Any]],
            match_mode: Any,
            search_mode: Any
    ) -> Sequence[Sequence[int]]:
        assert search_range and search_range[0]

        if len(search_range[0]) == 1:
            values = [row[0] for row in search_range]
        else:
            values = search_range[0]

        match_mode = self._get_match_mode(match_mode)
        search_mode = self._get_search_mode(search_mode)

        idx = self._match_value(criterion, values, match_mode, search_mode)
        ret = None if idx is None else idx + 1
        return [[ret]]

    def _match_value(self, criterion: Any, values: Sequence[Any],
                     match_mode: XMatchMode, search_mode: XSearchMode):
        if criterion is None:
            raise self._illegal_argument_exception("Criterion")

        if (
                match_mode == XMatchMode.WILDCARD
                or match_mode == XMatchMode.REGEX
        ) and (
                search_mode == XSearchMode.FIRST_BINARY
                or search_mode == XSearchMode.LAST_BINARY
        ):
            raise self._illegal_argument_exception(
                "Incompatible MatchMode/SearchMode")

        finder = IndexFinder(self._oCollator, self._illegal_argument_exception)

        if search_mode == XSearchMode.FIRST:
            return finder.find_index(
                criterion, values, match_mode, reverse=False)
        elif search_mode == XSearchMode.LAST:
            return finder.find_index(
                criterion, values, match_mode, reverse=True)
        elif search_mode == XSearchMode.FIRST_BINARY:
            return finder.binary_find_index(
                criterion, values, match_mode, reverse=False)
        elif search_mode == XSearchMode.LAST_BINARY:
            return finder.binary_find_index(
                criterion, values, match_mode, reverse=True)


class IndexFinder:
    def __init__(self, oCollator, illegal_argument_exception: Any):
        self._oCollator = oCollator
        self._illegal_argument_exception = illegal_argument_exception

    def find_index(
            self, criterion: Any,
            values: Sequence[Any],
            match_mode: Any,
            reverse: bool
    ) -> Optional[int]:
        if match_mode == XMatchMode.EXACT:
            eq_values = create_eq_values_with_collator(
                self._oCollator)
            ret = self._find_eq_value_index(
                eq_values, criterion, values, reverse)
            return ret
        elif match_mode == XMatchMode.SMALLER:
            cmp_values = create_cmp_values_with_collator(self._oCollator)
            return self._find_smaller_value_index(
                cmp_values, criterion, values, reverse)
        elif match_mode == XMatchMode.LARGER:
            cmp_values = create_cmp_values_with_collator(self._oCollator)
            return self._find_larger_value_index(
                cmp_values, criterion, values, reverse)
        elif match_mode == XMatchMode.WILDCARD:
            raise NotImplementedError()
        elif match_mode == XMatchMode.REGEX:
            raise NotImplementedError()

    def _find_eq_value_index(
            self, eq_values: Callable[[Any, Any], bool], criterion: Any,
            values: Sequence[Any], reverse: bool) -> Optional[int]:
        """Lookup for the first value using an eq_values function"""
        for i in self._get_indices(values, reverse):
            value = values[i]
            if eq_values(value, criterion):
                return i
        return None

    def _get_indices(self, values: Sequence[Any], reverse: bool) -> Iterable[
        int]:
        if reverse:
            indices = range(len(values) - 1, -1, -1)
        else:
            indices = range(len(values))
        return indices

    def _find_smaller_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any], reverse: bool) -> Optional[int]:
        """Lookup for the first value using an cmp_values function"""
        ABSOLUTE_NONE = object()

        cur = ABSOLUTE_NONE
        cur_idx = None
        for i in self._get_indices(values, reverse):
            search_value = values[i]
            cmp_to_criterion = cmp_values(search_value, criterion)
            if cmp_to_criterion == 0:
                return i
            elif cmp_to_criterion == -1:
                if cur is ABSOLUTE_NONE or cmp_values(search_value, cur) == 1:
                    cur = search_value
                    cur_idx = i

        return cur_idx

    def _find_larger_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any], reverse: bool) -> Optional[int]:
        """Lookup for the first value using an cmp_values function"""
        ABSOLUTE_NONE = object()

        cur = ABSOLUTE_NONE
        cur_idx = None

        for i in self._get_indices(values, reverse):
            search_value = values[i]
            cmp_to_criterion = cmp_values(search_value, criterion)
            if cmp_to_criterion == 0:
                return i
            elif cmp_to_criterion == 1:
                if cur is ABSOLUTE_NONE or cmp_values(search_value, cur) == -1:
                    cur = search_value
                    cur_idx = i

        return cur_idx

    # BINARY FIRST BY ROW *****************************************************
    def binary_find_index(
            self, criterion: Any,
            values: Sequence[Any],
            match_mode: Any,
            reverse: bool
    ) -> Optional[int]:
        if match_mode == XMatchMode.EXACT:
            cmp_values = create_cmp_values_with_collator(
                self._oCollator)
            if reverse:
                return self._find_binary_last_eq_value_index(
                    cmp_values, criterion, values)
            else:
                return self._find_binary_first_eq_value_index(
                    cmp_values, criterion, values)
        elif match_mode == XMatchMode.SMALLER:
            cmp_values = create_cmp_values_with_collator(self._oCollator)
            if reverse:
                return self._find_binary_last_smaller_value_index(
                    cmp_values, criterion, values)
            else:
                return self._find_binary_first_smaller_value_index(
                    cmp_values, criterion, values)
        elif match_mode == XMatchMode.LARGER:
            cmp_values = create_cmp_values_with_collator(self._oCollator)
            if reverse:
                return self._find_binary_last_larger_value_index(
                    cmp_values, criterion, values)
            else:
                return self._find_binary_first_larger_value_index(
                    cmp_values, criterion, values)
        elif match_mode == XMatchMode.WILDCARD:
            raise NotImplementedError()
        elif match_mode == XMatchMode.REGEX:
            raise NotImplementedError()

    def _find_binary_first_eq_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any]) -> Optional[int]:
        """Lookup for the first value using an eq_values function"""
        # values[idx - 1] < criterion <= values[idx]
        idx = bisect_left(values, criterion, cmp_values)
        if idx < len(values) and cmp_values(values[idx], criterion) == 0:
            return idx

        return None

    def _find_binary_last_eq_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any]) -> Optional[int]:
        """Lookup for the first value using an eq_values function"""
        # values[idx - 1] <= criterion < values[idx]
        idx = bisect_right(values, criterion, cmp_values)
        if idx > 0:
            new_idx = idx - 1
            if cmp_values(values[new_idx], criterion) == 0:
                return new_idx

        return None

    def _find_binary_first_smaller_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any]) -> Optional[int]:
        """Lookup for the first value using a cmp_values function"""
        # values[idx - 1] < criterion <= values[idx]
        idx = bisect_left(values, criterion, cmp_values)
        if idx < len(values) and cmp_values(values[idx], criterion) == 0:
            return idx
        elif idx > 0:  # criterion > values[idx-1]
            new_criterion = values[idx - 1]
            # ret < idx and values[ret] == new_criterion
            return bisect_left(values, new_criterion, cmp_values)

    def _find_binary_last_smaller_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any]) -> Optional[int]:
        """Lookup for the first value using a cmp_values function"""
        # values[idx - 1] <= criterion < values[idx]
        idx = bisect_right(values, criterion, cmp_values)
        if idx == 0:   # criterion < values[0]
            return None

        new_idx = idx - 1
        if cmp_values(values[new_idx], criterion) == 0:
            return new_idx
        else:
            new_criterion = values[new_idx]
            # other_idx - 1 < new_idx and values[other_idx] == new_criterion
            other_idx = bisect_right(values, new_criterion, cmp_values)
            assert other_idx > 0
            return other_idx - 1

    def _find_binary_first_larger_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any]) -> Optional[int]:
        # values[idx - 1] <= criterion < values[idx]
        idx = bisect_right(values, criterion, cmp_values)
        if idx == 0:  # criterion < values[0]
            return 0

        new_idx = idx - 1
        if cmp_values(values[new_idx], criterion) == 0:
            # ret < idx and values[ret] == criterion
            return bisect_left(values, criterion, cmp_values)
        elif idx < len(values):
            return idx

        return None

    def _find_binary_last_larger_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any]) -> Optional[int]:
        # values[idx - 1] <= criterion < values[idx]
        idx = bisect_right(values, criterion, cmp_values)
        if idx == 0:  # criterion < values[0]
            return None

        new_idx = idx - 1
        if cmp_values(values[new_idx], criterion) == 0:
            return new_idx
        elif idx == len(values):
            return None
        else:
            new_criterion = values[idx]
            # ret >= idx and values[ret + 1] == new_criterion
            ret = bisect_right(values, new_criterion, cmp_values) - 1
            return ret



def create_cmp_values_with_collator(oCollator) -> Callable[[Any, Any], int]:
    def cmp_values_with_collator(x: Any, y: Any) -> int:
        """
        A comparison function is any callable that accepts two arguments,
        compares them,
        and returns a negative number for less-than,
        zero for equality,
        or a positive number for greater-than.

        float < str < None
        """
        if isinstance(x, (int, float)):
            if isinstance(y, (int, float)):
                if x < y:
                    return -1
                elif x > y:
                    return 1
                else:
                    return 0
            else:  # float < str or None
                return -1
        elif isinstance(x, str):
            if isinstance(y, (int, float)):  # str > float
                return 1
            elif isinstance(y, str):
                return oCollator.compareString(x, y)
            else:  # str < None
                return -1
        elif x is None:
            if y is None:
                return 0
            else:  # None > str
                return 1
        else:
            return 0

    return cmp_values_with_collator


def create_eq_values_with_collator(oCollator) -> Callable[[Any, Any], bool]:
    def eq_values_with_collator(x: Any, y: Any) -> bool:
        if x is None:
            return False
        elif isinstance(x, str) and isinstance(y, str):
            return oCollator.compareString(x, y) == 0
        else:
            return x == y

    return eq_values_with_collator


class Orientation(enum.Enum):
    BY_COL = 1
    BY_ROW = 2


def get_orientation(
        row_or_col_range: Sequence[Sequence[Any]],
        base_range: Sequence[Sequence[Any]]
) -> Optional[Orientation]:
    if len(row_or_col_range) == len(base_range) and len(
            row_or_col_range[0]) == 1:  # it's a col
        orientation = Orientation.BY_ROW
    elif len(row_or_col_range) == 1 and len(row_or_col_range[0]) == len(
            base_range[0]):
        orientation = Orientation.BY_COL
    else:
        orientation = None
    return orientation


# from https://github.com/python/cpython/blob/main/Lib/bisect.py
def bisect_right(a, x, cmp):
    lo = 0
    hi = len(a)
    while lo < hi:
        mid = (lo + hi) // 2
        if cmp(x, a[mid]) < 0:
            hi = mid
        else:
            lo = mid + 1
    return lo


def bisect_left(a, x, cmp):
    lo = 0
    hi = len(a)
    while lo < hi:
        mid = (lo + hi) // 2
        if cmp(a[mid], x) < 0:
            lo = mid + 1
        else:
            hi = mid
    return lo
