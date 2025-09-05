import unittest

from xlcalculator.xlfunctions import lookup, xlerrors, func_xltypes


class LookupModuleTest(unittest.TestCase):

    def test_CHOOSE(self):
        self.assertEqual(lookup.CHOOSE('2', 2, 4, 6), 4)

    def test_CHOOSE_with_negative_index(self):
        self.assertIsInstance(
            lookup.CHOOSE(-1, 1, 2, 3), xlerrors.ValueExcelError)

    def test_CHOOSE_with_too_large_index(self):
        self.assertIsInstance(
            lookup.CHOOSE(5, 1, 2, 3), xlerrors.ValueExcelError)

    def test_VLOOOKUP(self):
        # Excel Doc example.
        range1 = func_xltypes.Array([
            [101, 'Davis', 'Sara'],
            [102, 'Fortana', 'Olivier'],
            [103, 'Leal', 'Karina'],
            [104, 'Patten', 'Michael'],
            [105, 'Burke', 'Brian'],
            [106, 'Sousa', 'Luis'],
        ])
        self.assertEqual(lookup.VLOOKUP(102, range1, 2, False), 'Fortana')

    def test_VLOOOKUP_with_range_lookup(self):
        with self.assertRaises(NotImplementedError):
            lookup.VLOOKUP(1, func_xltypes.Array([[]]), 2, True)

    def test_VLOOOKUP_with_oversized_col_index_num(self):
        # Excel Doc example.
        range1 = func_xltypes.Array([
            [101, 'Davis', 'Sara'],
        ])
        self.assertIsInstance(
            lookup.VLOOKUP(102, range1, 4, False), xlerrors.ValueExcelError)

    def test_VLOOOKUP_with_unknown_lookup_value(self):
        range1 = func_xltypes.Array([
            [101, 'Davis', 'Sara'],
        ])
        self.assertIsInstance(
            lookup.VLOOKUP(102, range1, 2, False), xlerrors.NaExcelError)

    def test_MATCH(self):
        range1 = [25, 28, 40, 41]
        self.assertEqual(lookup.MATCH(39, range1), 2)
        self.assertEqual(lookup.MATCH(39, range1, 1), 2)
        self.assertEqual(lookup.MATCH(40, range1, 1), 3)
        self.assertEqual(lookup.MATCH(40, range1, 0), 3)
        self.assertIsInstance(
            lookup.MATCH(40, range1, -1), xlerrors.NaExcelError
        )
        range2 = list(reversed(range1))
        self.assertEqual(lookup.MATCH(40, range2, -1), 2)
        self.assertIsInstance(
            lookup.MATCH(100, range2, -1), xlerrors.NaExcelError
        )
        self.assertIsInstance(
            lookup.MATCH(40, range2, 1), xlerrors.NaExcelError
        )
        self.assertIsInstance(
            lookup.MATCH(0, range1, 0), xlerrors.NaExcelError
        )

    def test_XLOOKUP_basic_exact_match(self):
        """Test basic XLOOKUP exact match functionality."""
        lookup_array = func_xltypes.Array([['Apple'], ['Banana'], ['Cherry']])
        return_array = func_xltypes.Array([[10], [20], [30]])
        
        self.assertEqual(lookup.XLOOKUP('Apple', lookup_array, return_array), 10)
        self.assertEqual(lookup.XLOOKUP('Banana', lookup_array, return_array), 20)
        self.assertEqual(lookup.XLOOKUP('Cherry', lookup_array, return_array), 30)

    def test_XLOOKUP_with_if_not_found(self):
        """Test XLOOKUP with custom if_not_found value."""
        lookup_array = func_xltypes.Array([['A'], ['B'], ['C']])
        return_array = func_xltypes.Array([[1], [2], [3]])
        
        self.assertEqual(lookup.XLOOKUP('D', lookup_array, return_array, 'Not Found'), 'Not Found')
        self.assertEqual(lookup.XLOOKUP('D', lookup_array, return_array, 0), 0)

    def test_XLOOKUP_not_found_without_default(self):
        """Test XLOOKUP returns error when value not found and no default provided."""
        lookup_array = func_xltypes.Array([['A'], ['B'], ['C']])
        return_array = func_xltypes.Array([[1], [2], [3]])
        
        result = lookup.XLOOKUP('D', lookup_array, return_array)
        self.assertIsInstance(result, xlerrors.NaExcelError)

    def test_XLOOKUP_horizontal_arrays(self):
        """Test XLOOKUP with horizontal (single row) arrays."""
        lookup_array = func_xltypes.Array([['Apple', 'Banana', 'Cherry']])
        return_array = func_xltypes.Array([[10, 20, 30]])
        
        self.assertEqual(lookup.XLOOKUP('Apple', lookup_array, return_array), 10)
        self.assertEqual(lookup.XLOOKUP('Banana', lookup_array, return_array), 20)

    def test_XLOOKUP_approximate_match_next_smallest(self):
        """Test XLOOKUP with match_mode=-1 (exact or next smallest)."""
        lookup_array = func_xltypes.Array([[10], [20], [30], [40]])
        return_array = func_xltypes.Array([['A'], ['B'], ['C'], ['D']])
        
        # Exact matches
        self.assertEqual(lookup.XLOOKUP(20, lookup_array, return_array, None, -1), 'B')
        # Next smallest
        self.assertEqual(lookup.XLOOKUP(25, lookup_array, return_array, None, -1), 'B')
        self.assertEqual(lookup.XLOOKUP(35, lookup_array, return_array, None, -1), 'C')

    def test_XLOOKUP_approximate_match_next_largest(self):
        """Test XLOOKUP with match_mode=1 (exact or next largest)."""
        lookup_array = func_xltypes.Array([[10], [20], [30], [40]])
        return_array = func_xltypes.Array([['A'], ['B'], ['C'], ['D']])
        
        # Exact matches
        self.assertEqual(lookup.XLOOKUP(20, lookup_array, return_array, None, 1), 'B')
        # Next largest
        self.assertEqual(lookup.XLOOKUP(15, lookup_array, return_array, None, 1), 'B')
        self.assertEqual(lookup.XLOOKUP(25, lookup_array, return_array, None, 1), 'C')

    def test_XLOOKUP_wildcard_match(self):
        """Test XLOOKUP with match_mode=2 (wildcard matching)."""
        lookup_array = func_xltypes.Array([['Apple'], ['Banana'], ['Cherry']])
        return_array = func_xltypes.Array([[10], [20], [30]])
        
        # Wildcard patterns
        self.assertEqual(lookup.XLOOKUP('App*', lookup_array, return_array, None, 2), 10)
        self.assertEqual(lookup.XLOOKUP('Ban?na', lookup_array, return_array, None, 2), 20)
        self.assertEqual(lookup.XLOOKUP('*erry', lookup_array, return_array, None, 2), 30)

    def test_XLOOKUP_reverse_search(self):
        """Test XLOOKUP with search_mode=-1 (reverse search)."""
        lookup_array = func_xltypes.Array([['A'], ['B'], ['A'], ['C']])
        return_array = func_xltypes.Array([[1], [2], [3], [4]])
        
        # Forward search finds first occurrence
        self.assertEqual(lookup.XLOOKUP('A', lookup_array, return_array, None, 0, 1), 1)
        # Reverse search finds last occurrence
        self.assertEqual(lookup.XLOOKUP('A', lookup_array, return_array, None, 0, -1), 3)

    def test_XLOOKUP_binary_search_ascending(self):
        """Test XLOOKUP with search_mode=2 (binary search ascending)."""
        lookup_array = func_xltypes.Array([[10], [20], [30], [40]])
        return_array = func_xltypes.Array([['A'], ['B'], ['C'], ['D']])
        
        self.assertEqual(lookup.XLOOKUP(20, lookup_array, return_array, None, 0, 2), 'B')
        self.assertEqual(lookup.XLOOKUP(30, lookup_array, return_array, None, 0, 2), 'C')

    def test_XLOOKUP_binary_search_descending(self):
        """Test XLOOKUP with search_mode=-2 (binary search descending)."""
        lookup_array = func_xltypes.Array([[40], [30], [20], [10]])
        return_array = func_xltypes.Array([['D'], ['C'], ['B'], ['A']])
        
        self.assertEqual(lookup.XLOOKUP(20, lookup_array, return_array, None, 0, -2), 'B')
        self.assertEqual(lookup.XLOOKUP(30, lookup_array, return_array, None, 0, -2), 'C')

    def test_XLOOKUP_binary_search_unsorted_array(self):
        """Test XLOOKUP binary search with unsorted array returns None."""
        lookup_array = func_xltypes.Array([[30], [10], [20], [40]])  # Unsorted
        return_array = func_xltypes.Array([['C'], ['A'], ['B'], ['D']])
        
        # Binary search should fail on unsorted array
        self.assertEqual(lookup.XLOOKUP(20, lookup_array, return_array, 'Not Found', 0, 2), 'Not Found')

    def test_XLOOKUP_array_dimension_mismatch(self):
        """Test XLOOKUP with mismatched array dimensions."""
        lookup_array = func_xltypes.Array([['A'], ['B']])  # 2 rows
        return_array = func_xltypes.Array([['1'], ['2'], ['3']])  # 3 rows
        
        result = lookup.XLOOKUP('A', lookup_array, return_array)
        self.assertIsInstance(result, xlerrors.ValueExcelError)

    def test_XLOOKUP_multi_dimensional_arrays(self):
        """Test XLOOKUP with multi-dimensional arrays (should error)."""
        lookup_array = func_xltypes.Array([['A', 'B'], ['C', 'D']])  # 2x2 array
        return_array = func_xltypes.Array([[1, 2], [3, 4]])
        
        result = lookup.XLOOKUP('A', lookup_array, return_array)
        self.assertIsInstance(result, xlerrors.ValueExcelError)

    def test_XLOOKUP_numeric_values(self):
        """Test XLOOKUP with numeric lookup values."""
        lookup_array = func_xltypes.Array([[100], [200], [300]])
        return_array = func_xltypes.Array([['Low'], ['Medium'], ['High']])
        
        self.assertEqual(lookup.XLOOKUP(200, lookup_array, return_array), 'Medium')
        self.assertEqual(lookup.XLOOKUP(250, lookup_array, return_array, 'Unknown'), 'Unknown')

    def test_XLOOKUP_mixed_data_types(self):
        """Test XLOOKUP with mixed data types."""
        lookup_array = func_xltypes.Array([[1], ['Text'], [3.14]])
        return_array = func_xltypes.Array([['One'], ['String'], ['Pi']])
        
        self.assertEqual(lookup.XLOOKUP(1, lookup_array, return_array), 'One')
        self.assertEqual(lookup.XLOOKUP('Text', lookup_array, return_array), 'String')
        self.assertEqual(lookup.XLOOKUP(3.14, lookup_array, return_array), 'Pi')

    def test_XLOOKUP_default_parameters(self):
        """Test XLOOKUP with default parameter values."""
        lookup_array = func_xltypes.Array([['A'], ['B'], ['C']])
        return_array = func_xltypes.Array([[1], [2], [3]])
        
        # Test with minimal parameters (should use defaults)
        self.assertEqual(lookup.XLOOKUP('B', lookup_array, return_array), 2)
        
        # Test explicit defaults
        self.assertEqual(lookup.XLOOKUP('B', lookup_array, return_array, None, 0, 1), 2)
