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

import itertools
import unittest

from lopolyfill_funcs import (
    XSearchMode, XMatchMode, IndexFinder, LopArrayHandling, Ignore,
    create_eq_criterion_with_regex, create_eq_criterion_with_wildcard)
from pythonpath.lopolyfill_funcs import (
    LopFilter, LopRandarray, LopSort, LopUnique, LopXMatch
)

DATA_1 = tuple([
    tuple(["{}{}".format(chr(j + ord("A")), i) for j in range(5)])
    for i in range(1, 5)
])

DATA_2 = tuple([
    tuple(["{}{}".format(chr(j + ord("A")), i) for j in range(3)])
    for i in range(11, 16)
])

COLUMN_1 = tuple([
    tuple(["{}{}".format(chr(j + ord("A")), i) for j in range(1)])
    for i in range(21, 41)
])

ROW_1 = tuple([
    tuple(["{}{}".format(chr(j + ord("A")), i) for j in range(20)])
    for i in range(1, 2)
])

SIMPLE_2_2_ARRAY = tuple([tuple([1, 2]), tuple([3, 4])])


class LopFilterTestCase(unittest.TestCase):
    def test_lo_filter_doc(self):
        f = LopFilter(ValueError).execute
        array = [
            ["Maths", "Physics", "Biology"],
            [47, 67, 33],
            [36, 68, 42],
            [40, 65, 44],
            [39, 64, 60],
            ['', 38, 43],
            [47, 84, 62],
            [29, 80, 51],
            [27, 49, 40],
            [57, 49, 12],
            [56, 33, 60],
            [57, '', ''],
            [26, '', ''],
        ]
        criteria = [[isinstance(row[0], int) and row[0] > 50] for row in array]
        self.assertEqual([
            [57, 49, 12],
            [56, 33, 60],
            [57, '', '']
        ], f(array, criteria, None))

    def test_rows(self):
        f = LopFilter(ValueError).execute
        self.assertEqual([(1, 2), (3, 4)],
                         f(SIMPLE_2_2_ARRAY, [[1], [1]], "default"))
        self.assertEqual([(1, 2)], f(SIMPLE_2_2_ARRAY, [[1], [0]], "default"))
        self.assertEqual([(3, 4)], f(SIMPLE_2_2_ARRAY, [[0], [1]], "default"))

    def test_cols(self):
        f = LopFilter(ValueError).execute
        self.assertEqual([[1, 2], [3, 4]],
                         f(SIMPLE_2_2_ARRAY, [[1, 1]], "default"))
        self.assertEqual([[1], [3]], f(SIMPLE_2_2_ARRAY, [[1, 0]], "default"))
        self.assertEqual([[2], [4]], f(SIMPLE_2_2_ARRAY, [[0, 1]], "default"))

    def test_empty(self):
        f = LopFilter(ValueError).execute
        self.assertEqual([['default']],
                         f(SIMPLE_2_2_ARRAY, [[0, 0]], "default"))
        self.assertEqual([['default']],
                         f(SIMPLE_2_2_ARRAY, [[0], [0]], "default"))

    def test_dimensions(self):
        f = LopFilter(ValueError).execute
        with self.assertRaises(ValueError) as err:
            f(SIMPLE_2_2_ARRAY, [[0, 0, 0]], "default")
        self.assertEqual("Bad criteria", err.exception.args[0])

        with self.assertRaises(ValueError) as err:
            f(SIMPLE_2_2_ARRAY, [[0], [0], [0]], "default")
        self.assertEqual("Bad criteria", err.exception.args[0])


class LopRandarrayTestCase(unittest.TestCase):
    def test_dimensions(self):
        f = LopRandarray(ValueError).execute
        data = f(5, 3, 10, 20, 1)
        self.assertEqual(5, len(data))
        self.assertEqual(3, len(data[0]))

    def test_integers(self):
        f = LopRandarray(ValueError).execute
        s = set()
        for _ in range(1000):
            data = f(5, 3, 10, 20, 1)
            values = list(itertools.chain(*data))
            s.update(values)
            for v in values:
                self.assertTrue(10 <= v <= 20)
                self.assertTrue(isinstance(v, int))

        self.assertEqual({10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20}, s)

    def test_floats(self):
        f = LopRandarray(ValueError).execute
        s = set()
        for _ in range(1000):
            data = f(5, 3, 10, 20, 0)
            values = list(itertools.chain(*data))
            s.update(values)
            for v in values:
                self.assertTrue(10 <= v <= 20)
                self.assertTrue(isinstance(v, float))

        self.assertAlmostEqual(10.0, min(s), 2)
        self.assertAlmostEqual(20.0, max(s), 2)


DOC_SORT_DATA_ARRAY = [
    ["Product Name", "Sales", "Revenue"],
    ["pencil", 20, 65],
    ["pen", 35, 85],
    ["notebook", 20, 190],
    ["book", 17, 180],
    ["pencil-case", "not", "not"],
]

OTHER_SORT_DATA_ARRAY = DOC_SORT_DATA_ARRAY + [["paper", None, None],
                                               ["z", 100, 2]]

SORTBY_DATA_ARRAY = [
    ["Product Name", "Sales", "Revenue"],
    ["pencil", 20, 65],
    ["pencil", 22, 70],
    ["pen", 35, 85],
    ["pen", 5, 15],
    ["pen", 5, 18],
    ["pen", 7, 19],
    ["notebook", 20, 190],
    ["notebook", 20, 191],
    ["notebook", 20, 187],
    ["book", 19, 180],
    ["book", 17, 180],
]


def _extract_cols(rows):
    return [[[v] for v in col] for col in zip(*rows)]


class LopSortTestCase(unittest.TestCase):
    def test_sort_doc(self):
        f = LopSort(SimpleCollator(), ValueError).sort
        self.assertEqual([
            ['book', 17, 180],
            ['pencil', 20, 65],
            ['notebook', 20, 190],
            ['pen', 35, 85],
            ['pencil-case', 'not', 'not'],
        ], f(DOC_SORT_DATA_ARRAY[1:], 2, 1, None)
        )

        self.assertEqual([
            ['pencil-case', 'not', 'not'],
            ['pen', 35, 85],
            ['pencil', 20, 65],
            ['notebook', 20, 190],
            ['book', 17, 180],
        ], f(DOC_SORT_DATA_ARRAY[1:], 2, -1, None)
        )

    def test_sort_none(self):
        f = LopSort(SimpleCollator(), ValueError).sort
        self.assertEqual([
            ['book', 17, 180],
            ['pencil', 20, 65],
            ['notebook', 20, 190],
            ['pen', 35, 85],
            ['z', 100, 2],
            ['pencil-case', 'not', 'not'],
            ['paper', None, None],
        ], f(OTHER_SORT_DATA_ARRAY[1:], 2, 1, None)
        )

        self.assertEqual([
            ['paper', None, None],
            ['pencil-case', 'not', 'not'],
            ['z', 100, 2],
            ['pen', 35, 85],
            ['pencil', 20, 65],
            ['notebook', 20, 190],
            ['book', 17, 180],
        ], f(OTHER_SORT_DATA_ARRAY[1:], 2, -1, None)
        )

    def test_sort_alpha(self):
        f = LopSort(SimpleCollator(), ValueError).sort
        self.assertEqual([
            ['book', 17, 180],
            ['notebook', 20, 190],
            ['paper', None, None],
            ['pen', 35, 85],
            ['pencil', 20, 65],
            ['pencil-case', 'not', 'not'],
            ['Product Name', 'Sales', 'Revenue'],
            ['z', 100, 2],
        ], f(OTHER_SORT_DATA_ARRAY, 1, 1, None)
        )

    def test_sort_col(self):
        f = LopSort(SimpleCollator(), ValueError).sort
        self.assertEqual([
            ('Revenue', 'Sales', 'Product Name'),
            (65, 20, 'pencil'),
            (85, 35, 'pen'),
            (190, 20, 'notebook'),
            (180, 17, 'book'),
            ('not', 'not', 'pencil-case'),
            (None, None, 'paper'),
            (2, 100, 'z'),
        ], f(OTHER_SORT_DATA_ARRAY, 8, 1, 1)
        )

    def test_sortby(self):
        lop_sort = LopSort(SimpleCollator(), ValueError)

        def f(*args):
            return lop_sort.sort_by(
                *(args + tuple([None for _ in range(31 - len(args))]))
            )

        rows = SORTBY_DATA_ARRAY[1:]
        sort_by_cols = _extract_cols(rows)
        self.assertEqual([
            ['pen', 35, 85],
            ['pencil', 22, 70],
            ['pencil', 20, 65],
            ['notebook', 20, 187],
            ['notebook', 20, 190],
            ['notebook', 20, 191],
            ['book', 19, 180],
            ['book', 17, 180],
            ['pen', 7, 19],
            ['pen', 5, 15],
            ['pen', 5, 18],
        ], f(
            rows,
            sort_by_cols[1], -1,
            sort_by_cols[2], 1,
        ))

    def test_sortby_reverse(self):
        lop_sort = LopSort(SimpleCollator(), ValueError)

        def f(*args):
            return lop_sort.sort_by(
                *(args + tuple([None for _ in range(31 - len(args))]))
            )

        rows = SORTBY_DATA_ARRAY[1:]
        sort_by_cols = _extract_cols(rows)
        self.assertEqual([
            ['pen', 5, 18],
            ['pen', 5, 15],
            ['pen', 7, 19],
            ['book', 17, 180],
            ['book', 19, 180],
            ['notebook', 20, 191],
            ['notebook', 20, 190],
            ['notebook', 20, 187],
            ['pencil', 20, 65],
            ['pencil', 22, 70],
            ['pen', 35, 85]
        ], f(
            rows,
            sort_by_cols[1], 1,
            sort_by_cols[2], -1,
        ))

    def test_sortby_cols(self):
        lop_sort = LopSort(SimpleCollator(), ValueError)

        def f(*args):
            return lop_sort.sort_by(
                *(args + tuple([None for _ in range(31 - len(args))]))
            )

        rows = list(zip(*SORTBY_DATA_ARRAY[1:]))
        self.assertEqual([
            ('pencil', 'pencil', 'pen', 'pen', 'pen', 'pen',
             'notebook', 'notebook', 'notebook', 'book', 'book'),
            (20, 22, 35, 5, 5, 7, 20, 20, 20, 19, 17),
            (65, 70, 85, 15, 18, 19, 190, 191, 187, 180, 180)
        ], rows)
        self.assertEqual([
            ('pen', 'pen', 'pen', 'book', 'book', 'notebook', 'notebook',
             'notebook', 'pencil', 'pencil', 'pen'),
            (5, 5, 7, 17, 19, 20, 20, 20, 20, 22, 35),
            (18, 15, 19, 180, 180, 191, 190, 187, 65, 70, 85),
        ], f(
            rows,
            [rows[1]], 1,
            [rows[2]], -1,
        ))


UNIQUE_DATA_ARRAY = [
    ["Name", "Grade", "Age", "Distance", "Weight"],
    ["Andy", 3, 9, 150, 40],
    ["Betty", 4, 10, 1000, 42],
    ["Charles", 3, 10, 300, 51],
    ["Daniel", 5, 11, 1200, 48],
    ["Eva", 2, 8, 650, 33],
    ["Frank", 2, 7, 300, 42],
    ["Greta", 1, 7, 200, 36],
    ["Harry", 3, 9, 1200, 44],
    ["Irene", 2, 8, 1000, 42],
]


class LopUniqueTestCase(unittest.TestCase):
    def test_doc(self):
        f = LopUnique(ValueError).execute
        rows = [row[1:3] for row in UNIQUE_DATA_ARRAY]
        self.assertEqual([
            ("Grade", "Age"),
            (3, 9),
            (4, 10),
            (3, 10),
            (5, 11),
            (2, 8),
            (2, 7),
            (1, 7),
        ], f(rows, False, False))
        self.assertEqual([
            ("Grade", "Age"),
            (4, 10),
            (3, 10),
            (5, 11),
            (2, 7),
            (1, 7),
        ], f(rows, False, True))

    def test_bycol(self):
        f = LopUnique(ValueError).execute
        rows = list(zip(*UNIQUE_DATA_ARRAY))[1:3]
        self.assertEqual(
            [
                ('Grade', 3, 4, 3, 5, 2, 2, 1, 3, 2),
                ('Age', 9, 10, 10, 11, 8, 7, 7, 9, 8)
            ], rows
        )
        self.assertEqual([
            ('Grade', 3, 4, 3, 5, 2, 2, 1),
            ('Age', 9, 10, 10, 11, 8, 7, 7)
        ], f(rows, True, False))
        self.assertEqual([
            ('Grade', 4, 3, 5, 2, 1),
            ('Age', 10, 10, 11, 7, 7)
        ], f(rows, True, True))


XLOOKUP_DATA_ARRAY = [
    ["Element", "Hydrogen", "Helium", "Lithium", "...", "Oganesson"],
    ["Symbol", "H", "He", "Li", "...", "Og"],
    ["Atomic Number", 1, 2, 3, "...", 118],
    ["Relative Atomic Mass", 1.008, 4.0026, 6.94, "...", 294],
]

XBINARY_LOOKUP_DATA_ARRAY = [
    ["Hydrogen", "Helium", "Lithium", "Oganesson"],
    ["H", "He", "Li", "Og"],
    [1, 2, 3, 118],
    [1.008, 4.0026, 6.94, 294],
]


class LopXMatchTestCase(unittest.TestCase):
    def test_doc(self):
        f = LopXMatch(SimpleCollator(), ValueError).lookup
        search_range = [[row[0]] for row in XLOOKUP_DATA_ARRAY[1:]]
        self.assertEqual(
            [['Symbol'], ['Atomic Number'], ['Relative Atomic Mass']],
            search_range)
        self.assertEqual(
            [["Atomic Number", 1, 2, 3, "...", 118]],
            f("Atomic Number", search_range,
              XLOOKUP_DATA_ARRAY[1:], None, None, None)
        )

        search_range = [XLOOKUP_DATA_ARRAY[0]]
        self.assertEqual(
            [['Element', 'Hydrogen', 'Helium', 'Lithium', '...', 'Oganesson']],
            search_range)
        self.assertEqual(
            [
                ['Helium'],
                ['He'],
                [2],
                [4.0026],
            ],
            f("Helium", search_range, XLOOKUP_DATA_ARRAY, None,
              None, None)
        )

        self.assertEqual(
            [['Unknown element']],
            f("Kryptonite", [XLOOKUP_DATA_ARRAY[0]], XLOOKUP_DATA_ARRAY,
              "Unknown element", None, None)
        )

    def test_exact_binary(self):
        f = LopXMatch(SimpleCollator(), ValueError).lookup
        search_range = [XBINARY_LOOKUP_DATA_ARRAY[2]]
        self.assertEqual(
            [[1, 2, 3, 118]],
            search_range)
        self.assertEqual(
            [
                ['Helium'], ['He'], [2], [4.0026]
            ],
            f(2, search_range, XBINARY_LOOKUP_DATA_ARRAY, None,
              XMatchMode.EXACT, XSearchMode.FIRST_BINARY)
        )
        self.assertEqual(
            [
                ['Helium'], ['He'], [2], [4.0026]
            ],
            f(2, search_range, XBINARY_LOOKUP_DATA_ARRAY, None,
              XMatchMode.EXACT, XSearchMode.LAST_BINARY)
        )

    def test_smaller(self):
        f = LopXMatch(SimpleCollator(), ValueError).lookup
        search_range = [XLOOKUP_DATA_ARRAY[2]]
        self.assertEqual(
            [['Atomic Number', 1, 2, 3, '...', 118]],
            search_range)
        self.assertEqual(
            [
                ['Helium'], ['He'], [2], [4.0026]
            ],
            f(2.5, search_range, XLOOKUP_DATA_ARRAY, None,
              XMatchMode.SMALLER, None)
        )
        self.assertEqual(
            [
                ['Helium'], ['He'], [2], [4.0026]
            ],
            f(2.5, search_range, XLOOKUP_DATA_ARRAY, None,
              XMatchMode.SMALLER, XSearchMode.LAST)
        )

    def test_binary_smaller(self):
        f = LopXMatch(SimpleCollator(), ValueError).lookup
        search_range = [XBINARY_LOOKUP_DATA_ARRAY[2]]
        self.assertEqual(
            [[1, 2, 3, 118]],
            search_range)
        self.assertEqual(
            [
                ['Helium'], ['He'], [2], [4.0026]
            ],
            f(2.5, search_range, XBINARY_LOOKUP_DATA_ARRAY, None,
              XMatchMode.SMALLER, XSearchMode.FIRST_BINARY)
        )
        self.assertEqual(
            [
                ['Helium'], ['He'], [2], [4.0026]
            ],
            f(2.5, search_range, XBINARY_LOOKUP_DATA_ARRAY, None,
              XMatchMode.SMALLER, XSearchMode.LAST_BINARY)
        )

    def test_larger(self):
        f = LopXMatch(SimpleCollator(), ValueError).lookup
        search_range = [XLOOKUP_DATA_ARRAY[2]]
        self.assertEqual(
            [['Atomic Number', 1, 2, 3, '...', 118]],
            search_range)
        self.assertEqual(
            [
                ['Lithium'], ['Li'], [3], [6.94]
            ],
            f(2.5, search_range, XLOOKUP_DATA_ARRAY, None,
              XMatchMode.LARGER, None)
        )
        self.assertEqual(
            [
                ['Lithium'], ['Li'], [3], [6.94]
            ],
            f(2.5, search_range, XLOOKUP_DATA_ARRAY, None,
              XMatchMode.LARGER, XSearchMode.LAST)
        )

    def test_binary_larger(self):
        f = LopXMatch(SimpleCollator(), ValueError).lookup
        search_range = [XBINARY_LOOKUP_DATA_ARRAY[2]]
        self.assertEqual(
            [[1, 2, 3, 118]],
            search_range)
        self.assertEqual(
            [
                ['Lithium'], ['Li'], [3], [6.94]
            ],
            f(2.5, search_range, XBINARY_LOOKUP_DATA_ARRAY, None,
              XMatchMode.LARGER, XSearchMode.FIRST_BINARY)
        )
        self.assertEqual(
            [
                ['Lithium'], ['Li'], [3], [6.94]
            ],
            f(2.5, search_range, XBINARY_LOOKUP_DATA_ARRAY, None,
              XMatchMode.LARGER, XSearchMode.LAST_BINARY)
        )

    def test_match_doc(self):
        f = LopXMatch(SimpleCollator(), ValueError).match
        search_range = [[row[0]] for row in XLOOKUP_DATA_ARRAY]
        self.assertEqual(
            [['Element'], ['Symbol'], ['Atomic Number'],
             ['Relative Atomic Mass']],
            search_range)
        self.assertEqual(
            3,
            f("Atomic Number", search_range, None, None)
        )

        search_range = [XLOOKUP_DATA_ARRAY[1]]
        self.assertEqual(
            [['Symbol', 'H', 'He', 'Li', '...', 'Og']],
            search_range)
        self.assertEqual(
            4,
            f("Li", search_range, None, None)
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


class IndexFinderTestCase(unittest.TestCase):
    def test_exact(self):
        finder = IndexFinder(SimpleCollator(), ValueError)
        values = [1, 1, 1, 2, 2, 3, 3, 3, 4, 4, 6, 8, 11, 11]

        for f in (finder.find_index, finder.binary_find_index):
            self.assertEqual(0, f(1, values, XMatchMode.EXACT, reverse=False))
            self.assertEqual(2, f(1, values, XMatchMode.EXACT, reverse=True))
            self.assertEqual(12, f(11, values, XMatchMode.EXACT, reverse=False))
            self.assertEqual(13, f(11, values, XMatchMode.EXACT, reverse=True))

            self.assertEqual(5, f(3, values, XMatchMode.EXACT, reverse=False))
            self.assertEqual(7, f(3, values, XMatchMode.EXACT, reverse=True))
            self.assertIsNone(f(2.5, values, XMatchMode.EXACT, reverse=False))
            self.assertIsNone(f(2.5, values, XMatchMode.EXACT, reverse=True))

            self.assertIsNone(f(0, values, XMatchMode.EXACT, reverse=False))
            self.assertIsNone(f(0, values, XMatchMode.EXACT, reverse=True))
            self.assertIsNone(f(20, values, XMatchMode.EXACT, reverse=False))
            self.assertIsNone(f(20, values, XMatchMode.EXACT, reverse=True))

    def test_first_smaller(self):
        finder = IndexFinder(SimpleCollator(), ValueError)
        values = [1, 1, 1, 2, 2, 3, 3, 3, 4, 4, 6, 8, 11, 11]

        for f in (finder.find_index, finder.binary_find_index):
            self.assertEqual(0, f(1, values, XMatchMode.SMALLER, reverse=False))
            self.assertEqual(2, f(1, values, XMatchMode.SMALLER, reverse=True))
            self.assertEqual(12,
                             f(11, values, XMatchMode.SMALLER, reverse=False))
            self.assertEqual(13,
                             f(11, values, XMatchMode.SMALLER, reverse=True))

            self.assertEqual(
                5, f(3, values, XMatchMode.SMALLER, reverse=False))
            self.assertEqual(
                7, f(3, values, XMatchMode.SMALLER, reverse=True))
            self.assertEqual(
                3, f(2.5, values, XMatchMode.SMALLER, reverse=False))
            self.assertEqual(
                4, f(2.5, values, XMatchMode.SMALLER, reverse=True))

            self.assertIsNone(
                f(0, values, XMatchMode.SMALLER, reverse=False))
            self.assertIsNone(
                f(0, values, XMatchMode.SMALLER, reverse=True))

    def test_first_larger(self):
        finder = IndexFinder(SimpleCollator(), ValueError)
        values = [1, 1, 1, 2, 2, 3, 3, 3, 4, 4, 6, 8, 11, 11]

        for f in (finder.find_index, finder.binary_find_index):
            self.assertEqual(0, f(1, values, XMatchMode.LARGER, reverse=False))
            self.assertEqual(2, f(1, values, XMatchMode.LARGER, reverse=True))
            self.assertEqual(12,
                             f(11, values, XMatchMode.LARGER, reverse=False))
            self.assertEqual(13, f(11, values, XMatchMode.LARGER, reverse=True))

            self.assertEqual(
                5, f(3, values, XMatchMode.LARGER, reverse=False))
            self.assertEqual(
                7, f(3, values, XMatchMode.LARGER, reverse=True))
            self.assertEqual(
                8, f(3.5, values, XMatchMode.LARGER, reverse=False))
            self.assertEqual(
                9, f(3.5, values, XMatchMode.LARGER, reverse=True))

            self.assertIsNone(
                f(20, values, XMatchMode.LARGER, reverse=False))
            self.assertIsNone(
                f(20, values, XMatchMode.LARGER, reverse=True))

    def test_last_larger(self):
        finder = IndexFinder(SimpleCollator(), ValueError)
        values = [1, 2, 3, 118]

        self.assertEqual(2, finder.binary_find_index(2.5, values,
                                                     XMatchMode.LARGER,
                                                     reverse=False))
        self.assertEqual(2, finder.binary_find_index(2.5, values,
                                                     XMatchMode.LARGER,
                                                     reverse=True))

    def test_eq_with_wildcard_question_mark(self):
        criterion = "c?s"
        eq = create_eq_criterion_with_wildcard(criterion, True)
        self.assertTrue(eq('cas'))
        self.assertTrue(eq('cis'))
        self.assertFalse(eq('cars'))
        self.assertFalse(eq('bas'))
        self.assertFalse(eq('cs'))

    def test_eq_with_wildcard_star(self):
        criterion = "*cast"
        eq = create_eq_criterion_with_wildcard(criterion, True)
        self.assertTrue(eq("cast"))
        self.assertTrue(eq("forecast"))
        self.assertTrue(eq("outcast"))
        # If the option Search criteria = and <> must apply to whole cells is
        # disabled in Tools - Options - LibreOffice Calc - Calculate, then
        # "forecaster" will be a match using the "*cast" search string.
        self.assertFalse(eq('forecaster'))

    def test_eq_with_wildcard_tilde(self):
        criterion = "why~?"
        eq = create_eq_criterion_with_wildcard(criterion, True)
        self.assertTrue(eq("why?"))
        self.assertFalse(eq("whys"))
        self.assertFalse(eq("why~s"))

    def test_eq_with_wildcard_regex(self):
        criterion = "w.*y?"
        eq = create_eq_criterion_with_wildcard(criterion, True)
        self.assertTrue(eq("w.*y?"))
        self.assertFalse(eq("why"))

    def test_eq_with_regex(self):
        criterion = "a.+b"
        eq = create_eq_criterion_with_regex(criterion, True)
        self.assertTrue(eq('a__b'))
        self.assertTrue(eq('a b'))
        self.assertFalse(eq('a__bc'))
        self.assertFalse(eq('ab'))

    def test_eq_with_regex_part(self):
        criterion = "a.+b"
        eq = create_eq_criterion_with_regex(criterion, False)
        self.assertTrue(eq('a__b'))
        self.assertTrue(eq('a__bc'))
        self.assertTrue(eq('a b'))
        self.assertFalse(eq('ab'))

class LopArrayHandlingTestCase(unittest.TestCase):
    def test_choose_cols(self):
        self.assertEqual([
            ['A1', 'C1', 'A1', 'E1'],
            ['A2', 'C2', 'A2', 'E2'],
            ['A3', 'C3', 'A3', 'E3'],
            ['A4', 'C4', 'A4', 'E4']
        ], LopArrayHandling(ValueError).choose_cols(DATA_1, 1, 3, 1, -1))

    def test_choose_rows(self):
        self.assertEqual([
            ('A1', 'B1', 'C1', 'D1', 'E1'),
            ('A3', 'B3', 'C3', 'D3', 'E3'),
            ('A1', 'B1', 'C1', 'D1', 'E1'),
            ('A4', 'B4', 'C4', 'D4', 'E4')
        ], LopArrayHandling(ValueError).choose_rows(DATA_1, 1, 3, 1, -1))

    def test_vstack(self):
        self.assertEqual([
            ('A1', 'B1', 'C1', 'D1', 'E1'),
            ('A2', 'B2', 'C2', 'D2', 'E2'),
            ('A3', 'B3', 'C3', 'D3', 'E3'),
            ('A4', 'B4', 'C4', 'D4', 'E4'),
            ('A11', 'B11', 'C11', None, None),
            ('A12', 'B12', 'C12', None, None),
            ('A13', 'B13', 'C13', None, None),
            ('A14', 'B14', 'C14', None, None),
            ('A15', 'B15', 'C15', None, None)
        ], LopArrayHandling(ValueError).vstack(DATA_1, DATA_2))

    def test_htack_doc(self):
        ARRAY = (
            ("AAA", "BBB", "CCC", "DDD", "EEE"),
            ("FFF", "", "", "III", "JJJ"),
            ("KKK", "LLL", "MMM", "NNN", "    OOO"),
        )
        ARRAY1 = (
            ("PPP", "QQQ"),
            ("RRR", "SSS"),
            ("TTT", "UUU"),
            ("VVV", "WWW"),
            ("XXX", "YYY"),
        )

        self.assertEqual([
            ['AAA', 'BBB', 'CCC', 'DDD', 'EEE', 'PPP', 'QQQ'],
            ['FFF', '', '', 'III', 'JJJ', 'RRR', 'SSS'],
            ['KKK', 'LLL', 'MMM', 'NNN', '    OOO', 'TTT', 'UUU'],
            [None, None, None, None, None, 'VVV', 'WWW'],
            [None, None, None, None, None, 'XXX', 'YYY']
        ],
            LopArrayHandling(ValueError).hstack(ARRAY, ARRAY1))

    def test_hstack(self):
        self.assertEqual([
            ['A1', 'B1', 'C1', 'D1', 'E1', 'A11', 'B11', 'C11'],
            ['A2', 'B2', 'C2', 'D2', 'E2', 'A12', 'B12', 'C12'],
            ['A3', 'B3', 'C3', 'D3', 'E3', 'A13', 'B13', 'C13'],
            ['A4', 'B4', 'C4', 'D4', 'E4', 'A14', 'B14', 'C14'],
            [None, None, None, None, None, 'A15', 'B15', 'C15']
        ],
            LopArrayHandling(ValueError).hstack(DATA_1, DATA_2))

    def test_drop(self):
        self.assertEqual([
            ('A2', 'B2', 'C2', 'D2'),
            ('A3', 'B3', 'C3', 'D3'),
            ('A4', 'B4', 'C4', 'D4')
        ], LopArrayHandling(ValueError).drop(DATA_1, 1, -1))

    def test_take(self):
        self.assertEqual([
            ('D1', 'E1'), ('D2', 'E2')
        ], LopArrayHandling(ValueError).take(DATA_1, 2, -2))

    def test_expand(self):
        self.assertEqual([
            ('A11', 'B11', 'C11', None),
            ('A12', 'B12', 'C12', None),
            ('A13', 'B13', 'C13', None),
            ('A14', 'B14', 'C14', None),
            ('A15', 'B15', 'C15', None),
            (None, None, None, None)
        ], LopArrayHandling(ValueError).expand(DATA_2, 6, 4, None))

    def test_wrap_cols(self):
        self.assertEqual([
            ['A21', 'A28', 'A35'],
            ['A22', 'A29', 'A36'],
            ['A23', 'A30', 'A37'],
            ['A24', 'A31', 'A38'],
            ['A25', 'A32', 'A39'],
            ['A26', 'A33', 'A40'],
            ['A27', 'A34', 'foo']
        ], LopArrayHandling(ValueError).wraps_cols(
            COLUMN_1, 7, "foo"))
        self.assertEqual([
            ['A21', 'A25', 'A29', 'A33', 'A37'],
            ['A22', 'A26', 'A30', 'A34', 'A38'],
            ['A23', 'A27', 'A31', 'A35', 'A39'],
            ['A24', 'A28', 'A32', 'A36', 'A40']
        ], LopArrayHandling(ValueError).wraps_cols(
            COLUMN_1, 4, "foo"))

    def test_wrap_rows(self):
        self.assertEqual([
            ('A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1'),
            ('H1', 'I1', 'J1', 'K1', 'L1', 'M1', 'N1'),
            ('O1', 'P1', 'Q1', 'R1', 'S1', 'T1', 'foo')
        ], LopArrayHandling(ValueError).wraps_rows(
            ROW_1, 7, "foo"))

    def test_to_col(self):
        self.assertEqual([
            ['A1'],
            ['B1'],
            ['C1'],
            ['D1'],
            ['E1'],
            ['A2'],
            ['B2'],
            ['C2'],
            ['D2'],
            ['E2'],
            ['A3'],
            ['B3'],
            ['C3'],
            ['D3'],
            ['E3'],
            ['A4'],
            ['B4'],
            ['C4'],
            ['D4'],
            ['E4']
        ], LopArrayHandling(ValueError).to_col(DATA_1, Ignore.KEEP_ALL, False))
        self.assertEqual([
            ['A1'],
            ['A2'],
            ['A3'],
            ['A4'],
            ['B1'],
            ['B2'],
            ['B3'],
            ['B4'],
            ['C1'],
            ['C2'],
            ['C3'],
            ['C4'],
            ['D1'],
            ['D2'],
            ['D3'],
            ['D4'],
            ['E1'],
            ['E2'],
            ['E3'],
            ['E4']
        ], LopArrayHandling(ValueError).to_col(DATA_1, Ignore.KEEP_ALL, True))
        self.assertEqual([
            [1], [''], [''], [3], [2], [None]
        ], LopArrayHandling(ValueError).to_col(
            ((1, '', 2), ('', 3, None)), Ignore.KEEP_ALL, True))
        self.assertEqual([
            [1], [3], [2], [None]
        ], LopArrayHandling(ValueError).to_col(
            ((1, '', 2), ('', 3, None)), Ignore.IGNORE_BLANKS, True))
        self.assertEqual([
            [1], [''], [''], [3], [2]
        ], LopArrayHandling(ValueError).to_col(
            ((1, '', 2), ('', 3, None)), Ignore.IGNORE_ERRORS, True))

    def test_to_row(self):
        self.assertEqual([
            [1, 3, 2]
        ], LopArrayHandling(ValueError).to_row(
            ((1, '', 2), ('', 3, None)), Ignore.IGNORE_BLANKS_AND_ERRORS, True))


if __name__ == "__main__":
    unittest.main()
