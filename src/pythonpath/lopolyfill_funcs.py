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
import logging
import random
import re
from typing import Sequence, Any, Callable, List, Tuple, Optional, Iterable

DataRow = Tuple[Any, ...]
DataArray = Tuple[DataRow, ...]


def debug(func):
    logger = logging.getLogger(__name__)

    def aux(*args, **kwargs):
        logger.debug("%s(*%s, **%s)", func.__qualname__, args[1:], kwargs)
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.exception("LoPolyfill")
            raise
        finally:
            logger.debug("ok")

    return aux


class LopFilter:
    def __init__(self, illegal_argument_exception: Any):
        self._illegal_argument_exception = illegal_argument_exception

    @debug
    def execute(self, rows: DataArray,
                criteria: DataArray, default_value: Any):
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

    @debug
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
        if integers:
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

    @debug
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

    @debug
    def sort(
            self, in_range: DataArray, sort_index: Any,
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
            self, inRange: DataArray, sortIndex: int,
            ascending: bool):
        cols = list(zip(*inRange))
        sorted_cols = self._by_row_lop_sort(cols, sortIndex, ascending)
        return list(zip(*sorted_cols))

    def _by_row_lop_sort(
            self, rows: Sequence[DataRow], sort_index: int,
            ascending: bool) -> List[DataRow]:
        if sort_index < 0 or sort_index >= len(rows[0]):
            raise self._illegal_argument_exception("SortIndex col")

        cmp_values_with_collator = create_cmp_values_with_collator(
            self._oCollator)

        def cmp_rows(row1: Sequence[Any], row2: Sequence[Any]):
            return cmp_values_with_collator(row1[sort_index], row2[sort_index])

        return sorted(rows, key=functools.cmp_to_key(cmp_rows),
                      reverse=ascending is False)

    @debug
    def sort_by(
            self, inRange: DataArray,
            sortByRange1: DataArray, sortOrder1: int,
            *args: Any
    ):
        if not (inRange and inRange[0]):
            return inRange
        if not (sortByRange1 and sortByRange1[0]):
            return inRange

        h = len(inRange)
        w = len(inRange[0])
        byCol = self._is_by_col(sortByRange1, h, w)
        extract = self._create_extract(byCol, h, w)

        sortByRanges = (sortByRange1,) + args[0::2]
        sortOrders = (sortOrder1,) + args[1::2]

        sortKeys = [
            (extract(sortByRange), self._is_ascending(sortOrder))
            for sortByRange, sortOrder in zip(sortByRanges, sortOrders)
            if sortByRange is not None
        ]
        if byCol:
            columns = list(zip(*inRange))
            columns = self._by_row_lop_sortby(columns, sortKeys)
            return list(zip(*columns))
        else:
            return self._by_row_lop_sortby(inRange, sortKeys)

    def _is_by_col(
            self, sortByRange1: DataArray, h: int, w: int
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
    ) -> Callable[[DataArray], Sequence[Any]]:
        if byCol:
            def extract(sortByRange: DataArray) -> Sequence[Any]:
                if len(sortByRange) != 1 or len(sortByRange[0]) != w:
                    raise self._illegal_argument_exception("sortRange")
                return sortByRange[0]
        else:
            def extract(sortByRange: DataArray) -> Sequence[Any]:
                if len(sortByRange) != h or len(sortByRange[0]) != 1:
                    raise self._illegal_argument_exception("sortRange")
                return [row[0] for row in sortByRange]
        return extract

    def _by_row_lop_sortby(
            self, rows: Sequence[DataRow],
            sort_keys: List[Tuple[Any, bool]]
    ) -> List[DataRow]:
        cmp_values_with_collator = create_cmp_values_with_collator(
            self._oCollator)

        def cmp_integers(i: int, j: int) -> int:
            for values, ascending in sort_keys:
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

    @debug
    def execute(
            self, in_range: DataArray, by_col: Any,
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
    def __init__(
            self, oCollator, illegal_argument_exception: Any, whole_cell: bool):
        self._oCollator = oCollator
        self._illegal_argument_exception = illegal_argument_exception
        self._whole_cell = whole_cell

    @debug
    def lookup(
            self, criterion: Any,
            search_range: DataArray,
            result_range: DataArray,
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

    @debug
    def match(
            self, criterion: Any,
            search_range: DataArray,
            match_mode: Any,
            search_mode: Any
    ) -> int:
        assert search_range and search_range[0]

        values = self._extract_values(search_range)

        match_mode = self._get_match_mode(match_mode)
        search_mode = self._get_search_mode(search_mode)

        idx = self._match_value(criterion, values, match_mode, search_mode)
        ret = None if idx is None else idx + 1
        return ret

    def _extract_values(self, search_range):
        if len(search_range[0]) == 1:
            values = [row[0] for row in search_range]
        elif len(search_range) == 1:
            values = search_range[0]
        else:
            raise self._illegal_argument_exception("Search range")
        return values

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

        finder = IndexFinder(
            self._oCollator, self._illegal_argument_exception, self._whole_cell
        )

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
    def __init__(self, oCollator, illegal_argument_exception: Any,
                 whole_cell: bool):
        self._oCollator = oCollator
        self._illegal_argument_exception = illegal_argument_exception
        self._whole_cell = whole_cell

    def find_index(
            self, criterion: Any,
            values: Sequence[Any],
            match_mode: Any,
            reverse: bool
    ) -> Optional[int]:
        if match_mode == XMatchMode.EXACT:
            eq_criterion = create_eq_criterion_with_collator(
                self._oCollator, criterion, self._whole_cell)
            return self._find_eq_value_index(
                eq_criterion, values, reverse)
        elif match_mode == XMatchMode.SMALLER:
            cmp_values = create_cmp_values_with_collator(self._oCollator)
            return self._find_smaller_value_index(
                cmp_values, criterion, values, reverse)
        elif match_mode == XMatchMode.LARGER:
            cmp_values = create_cmp_values_with_collator(self._oCollator)
            return self._find_larger_value_index(
                cmp_values, criterion, values, reverse)
        elif match_mode == XMatchMode.WILDCARD:
            if isinstance(criterion, str):
                eq_criterion = create_eq_criterion_with_wildcard(
                    criterion, self._whole_cell)
            else:
                eq_criterion = lambda x: x == criterion
            return self._find_eq_value_index(
                eq_criterion, values, reverse)
        elif match_mode == XMatchMode.REGEX:
            if isinstance(criterion, str):
                eq_criterion = create_eq_criterion_with_regex(
                    criterion, self._whole_cell)
            else:
                eq_criterion = lambda x: x == criterion
            return self._find_eq_value_index(
                eq_criterion, values, reverse)

    def _find_eq_value_index(
            self, eq_criterion: Callable[[Any], bool],
            values: Sequence[Any], reverse: bool) -> Optional[int]:
        """Lookup for the first value using an eq_criterion function"""
        for i in self._get_indices(values, reverse):
            value = values[i]
            if eq_criterion(value):
                return i
        return None

    def _get_indices(
            self, values: Sequence[Any], reverse: bool
    ) -> Iterable[int]:
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
        """Lookup for the first value using an eq_criterion function"""
        # values[idx - 1] < criterion <= values[idx]
        idx = bisect_left(values, criterion, cmp_values)
        if idx < len(values) and cmp_values(values[idx], criterion) == 0:
            return idx

        return None

    def _find_binary_last_eq_value_index(
            self, cmp_values: Callable[[Any, Any], int], criterion: Any,
            values: Sequence[Any]) -> Optional[int]:
        """Lookup for the first value using an eq_criterion function"""
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
        if idx == 0:  # criterion < values[0]
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


class Ignore(enum.IntEnum):
    KEEP_ALL = 0
    IGNORE_BLANKS = 1
    IGNORE_ERRORS = 2
    IGNORE_BLANKS_AND_ERRORS = 3


class LopArrayHandling:
    """
    Future (25.8)
    """

    def __init__(self, illegal_argument_exception: Any):
        self._illegal_argument_exception = illegal_argument_exception

    def choose_cols(
            self, rows: DataArray, col_index1: int, *cols_indices: Any
    ) -> List[List[Any]]:
        assert rows and rows[0]

        cols_indices = [col_index1] + [
            int(col_index) for col_index in cols_indices
            if col_index is not None
        ]
        if any(i == 0 or i > len(rows[0]) or i <= -len(rows[0]) for i in
               cols_indices):
            raise self._illegal_argument_exception(
                "Indices {}".format(cols_indices))

        cols_indices = [i - 1 if i >= 1 else i for i in cols_indices]
        return [
            [row[i] for i in cols_indices]
            for row in rows
        ]

    @debug
    def choose_rows(
            self, rows: DataArray, row_index1: int, *rows_indices: Any
    ) -> List[List[Any]]:
        assert rows and rows[0]

        rows_indices = [row_index1] + [
            int(row_index) for row_index in rows_indices
            if row_index is not None
        ]
        if any(i == 0 or i > len(rows) or i <= -len(rows) for i in
               rows_indices):
            raise self._illegal_argument_exception(
                "Indices {}".format(rows_indices))

        rows_indices = [i - 1 if i >= 1 else i for i in rows_indices]
        return [
            rows[i] for i in rows_indices
        ]

    @debug
    def drop(
            self, rows: DataArray, row_count: int, col_count: Any
    ) -> List[DataRow]:
        if row_count is None:
            selected_rows = rows
        else:
            row_count = int(row_count)
            if row_count > 0:
                selected_rows = rows[row_count:]
            elif row_count < 0:
                selected_rows = rows[:row_count]
            else:
                raise self._illegal_argument_exception(
                    "Wrong row_count parameter")

        if col_count is None:
            return selected_rows

        col_count = int(col_count)
        if col_count > 0:
            return [row[col_count:] for row in selected_rows]
        elif col_count < 0:
            return [row[:col_count] for row in selected_rows]
        else:
            raise self._illegal_argument_exception("Wrong col_count parameter")

    @debug
    def take(
            self, rows: DataArray, row_count: int, col_count: int
    ) -> List[DataRow]:
        assert rows and rows[0]

        if row_count is None:
            selected_rows = rows
        else:
            row_count = int(row_count)
            if row_count > 0:
                selected_rows = rows[:row_count]
            elif row_count < 0:
                selected_rows = rows[row_count:]
            else:
                raise self._illegal_argument_exception(
                    "Wrong row_count parameter")

        if col_count is None:
            return selected_rows

        col_count = int(col_count)
        if col_count > 0:
            return [row[:col_count] for row in selected_rows]
        elif col_count < 0:
            return [row[col_count:] for row in selected_rows]
        else:
            raise self._illegal_argument_exception("Wrong col_count parameter")

    @debug
    def expand(
            self, rows: DataArray, row_count: Any, col_count: Any, pad_with: Any
    ) -> List[DataRow]:
        assert rows and rows[0]

        if row_count is None:
            missing_row_count = 0
        else:
            row_count = int(row_count)
            missing_row_count = row_count - len(rows)
            if missing_row_count < 0:
                raise self._illegal_argument_exception("Row count")

        if col_count is None:
            col_count = len(rows[0])
            missing_row_count = 0
            pad_col = tuple()
        else:
            col_count = int(col_count)
            missing_col_count = col_count - len(rows[0])
            if missing_col_count < 0:
                raise self._illegal_argument_exception("Row count")
            pad_col = (pad_with,) * missing_col_count

        padding_row = [(pad_with,) * col_count]
        return [
            row + pad_col
            for row in rows
        ] + padding_row * missing_row_count

    @debug
    def hstack(self, array: DataArray, *arrays: DataArray) -> List[List[Any]]:
        arrays = [array] + [array for array in arrays if array is not None]
        assert all(array and array[0] for array in arrays)

        h = max(len(rows) for rows in arrays)
        return [
            [
                v
                for rows in arrays
                for v in (rows[i] if i < len(rows) else (None,) * len(rows[0]))
            ]
            for i in range(h)
        ]

    @debug
    def vstack(self, array: DataArray, *arrays: DataArray) -> List[DataRow]:
        arrays = [array] + [array for array in arrays if array is not None]
        assert all(array and array[0] for array in arrays)

        w = max(len(rows[0]) for rows in arrays)
        return [
            row + (None,) * (w - len(row))
            for row in itertools.chain(*arrays)
        ]

    @debug
    def to_col(
            self, rows: DataArray, ignore: Any, scan_by_col: Any
    ) -> List[DataRow]:
        if ignore is None:
            ignore = Ignore.KEEP_ALL
        else:
            ignore = Ignore(int(ignore))
        dont_ignore = self._get_dont_ignore_func(ignore)

        if scan_by_col:
            return [
                (x,) for x in itertools.chain(*zip(*rows)) if dont_ignore(x)
            ]
        else:
            return [(x,) for x in itertools.chain(*rows) if dont_ignore(x)]

    @debug
    def to_row(
            self, rows: DataArray, ignore: Any, scan_by_col: Any
    ) -> List[List[Any]]:
        if ignore is None:
            ignore = Ignore.KEEP_ALL
        else:
            ignore = Ignore(int(ignore))
        dont_ignore = self._get_dont_ignore_func(ignore)

        if scan_by_col:
            return [
                [x for x in itertools.chain(*zip(*rows)) if dont_ignore(x)]]
        else:
            return [[x for x in itertools.chain(*rows) if dont_ignore(x)]]

    def _get_dont_ignore_func(self, ignore: Ignore) -> Callable[[Any], bool]:
        if ignore == Ignore.IGNORE_BLANKS:
            def dont_ignore(x):
                return x != ''
        elif ignore == Ignore.IGNORE_ERRORS:
            def dont_ignore(x):
                return x is not None
        elif ignore == Ignore.IGNORE_BLANKS_AND_ERRORS:
            def dont_ignore(x):
                return x != '' and x is not None
        elif ignore is None or ignore == Ignore.KEEP_ALL:
            def dont_ignore(x):
                return True
        else:
            raise self._illegal_argument_exception("Ignore")
        return dont_ignore

    @debug
    def wraps_cols(
            self, rows: DataArray, wrap_count: int, pad_with: Any
    ) -> List[List[Any]]:
        assert rows and rows[0]

        values = self._extract_values(rows)
        d, m = divmod(len(values), wrap_count)
        if m != 0:
            d += 1
            values = values + (pad_with,) * (wrap_count - m)

        return [
            [values[j] for j in range(i, len(values), wrap_count)]
            for i in range(wrap_count)
        ]

    @debug
    def wraps_rows(
            self, rows: DataArray, wrap_count: int, pad_with: Any
    ) -> List[DataRow]:
        assert rows and rows[0]

        values = self._extract_values(rows)
        d, m = divmod(len(values), wrap_count)
        if m != 0:
            d += 1
            values = values + (pad_with,) * (wrap_count - m)

        return [
            values[i:i + wrap_count]
            for i in range(0, len(values), wrap_count)
        ]

    def _extract_values(self, rows: DataArray) -> DataRow:
        if len(rows[0]) == 1:
            values = tuple([row[0] for row in rows])
        elif len(rows) == 1:
            values = rows[0]
        else:
            raise self._illegal_argument_exception("Expect a vector")
        return values


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


def create_eq_criterion_with_collator(
        oCollator, criterion: Any, whole_cell: bool
) -> Callable[[Any], bool]:
    # todo: use whole_cell
    if isinstance(criterion, str):
        def eq_criterion_with_collator(x: Any) -> bool:
            return isinstance(x, str) and oCollator.compareString(x,
                                                                  criterion) == 0
    else:
        def eq_criterion_with_collator(x: Any) -> bool:
            return x == criterion

    return eq_criterion_with_collator


WILDCARD_REGEX = re.compile(r"(~?[?*])")


def create_eq_criterion_with_wildcard(
        criterion: str, whole_cell: bool) -> Callable[[Any], bool]:
    parts = WILDCARD_REGEX.split(criterion)
    for i in range(len(parts)):
        if i % 2 == 0:  # non wildcard
            parts[i] = re.escape(parts[i])
        elif parts[i].startswith("~"):
            parts[i] = "\\" + parts[i][1:]
        elif parts[i] == "?":
            parts[i] = "."
        elif parts[i] == "*":
            parts[i] = ".*"
    criterion_pattern = "".join(parts)

    if whole_cell:
        criterion_pattern = "^" + criterion_pattern + "$"

    regex = re.compile(criterion_pattern, re.I)

    def eq_criterion_with_wildcard(x: Any) -> bool:
        return isinstance(x, str) and regex.match(x)

    return eq_criterion_with_wildcard


def create_eq_criterion_with_regex(
        criterion: str, whole_cell: bool) -> Callable[[Any], bool]:
    if whole_cell:
        criterion_pattern = "^" + criterion + "$"
    else:
        criterion_pattern = criterion

    regex = re.compile(criterion_pattern, re.I)

    def eq_criterion_with_regex(x: Any) -> bool:
        return isinstance(x, str) and regex.match(x)

    return eq_criterion_with_regex


class Orientation(enum.Enum):
    BY_COL = 1
    BY_ROW = 2


def get_orientation(
        row_or_col_range: DataArray,
        base_range: DataArray
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
