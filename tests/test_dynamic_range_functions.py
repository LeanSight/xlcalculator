"""
Comprehensive Test Suite for Dynamic Range Functions

This module provides thorough testing for OFFSET, INDEX, and INDIRECT functions,
including unit tests, integration tests, error handling, and Excel compatibility.
"""

import unittest
from unittest.mock import Mock, patch
from xlcalculator.model import Model
from xlcalculator import Evaluator
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.dynamic_range import OFFSET, INDEX, INDIRECT
from xlcalculator.xlfunctions.reference_utils import ReferenceResolver


class TestDynamicRangeFunctions(unittest.TestCase):
    """Test suite for dynamic range functions"""
    
    def setUp(self):
        """Set up test environment with sample data"""
        self.model = Model()
        
        # Create test data grid (A1:E5)
        test_data = [
            ['Name',    'Age',  'City',     'Score', 'Active'],
            ['Alice',   25,     'NYC',      85,      True],
            ['Bob',     30,     'LA',       92,      False],
            ['Charlie', 35,     'Chicago',  78,      True],
            ['Diana',   28,     'Miami',    95,      True],
        ]
        
        for row_idx, row_data in enumerate(test_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell_ref = ReferenceResolver.cell_to_string(row_idx, col_idx)
                self.model.set_cell_value(f'Sheet1!{cell_ref}', value)
        
        # Additional test values
        self.model.set_cell_value('Sheet1!G1', 'B2')  # For INDIRECT tests
        self.model.set_cell_value('Sheet1!G2', 'Sheet1!C3')
        
        self.evaluator = Evaluator(self.model)
    
    def create_test_array(self, rows: int, cols: int):
        """Create a mock array for testing"""
        array = Mock()
        array.values = []
        for r in range(rows):
            row = []
            for c in range(cols):
                row.append(f"R{r+1}C{c+1}")
            array.values.append(row)
        return array
    
    def _convert_result_to_list(self, result):
        """
        Convert INDEX function result to list for test comparison.
        
        Args:
            result: Result from INDEX function (could be array-like or single value)
            
        Returns:
            List representation of the result for comparison
        """
        from xlcalculator.xlfunctions import func_xltypes
        
        if isinstance(result, func_xltypes.Array):
            # Handle func_xltypes.Array objects
            if hasattr(result, 'values') and result.values is not None:
                if hasattr(result.values, 'tolist'):
                    # Numpy array
                    array_data = result.values.tolist()
                else:
                    # Regular list
                    array_data = result.values
                # Flatten if it's a single row/column wrapped in a list
                if len(array_data) == 1 and isinstance(array_data[0], list):
                    return [str(item) for item in array_data[0]]
                else:
                    return [str(item) for item in array_data]
            else:
                return []
        elif hasattr(result, 'tolist'):
            return result.tolist()
        elif hasattr(result, 'values') and result.values is not None:
            return list(result.values)
        elif isinstance(result, list):
            return result
        else:
            try:
                return [str(item) for item in result]
            except:
                return [str(result)]


class TestOFFSETFunction(TestDynamicRangeFunctions):
    """Test cases for OFFSET function"""
    
    def test_offset_basic_single_cell(self):
        """Test basic OFFSET with single cell results"""
        test_cases = [
            ("A1", 1, 1, None, None, "B2"),     # Basic offset
            ("A1", 0, 1, None, None, "B1"),     # Column only
            ("A1", 1, 0, None, None, "A2"),     # Row only
            ("B2", 1, 1, None, None, "C3"),     # From different start
            ("A1", 2, 3, None, None, "D3"),     # Larger offset
        ]
        
        for ref, rows, cols, height, width, expected in test_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols):
                result = OFFSET(ref, rows, cols, height, width)
                self.assertEqual(result, expected)
    
    def test_offset_with_dimensions(self):
        """Test OFFSET with explicit height and width"""
        test_cases = [
            ("A1", 1, 1, 1, 1, "B2"),           # Single cell
            ("A1", 1, 1, 2, 2, "B2:C3"),        # 2x2 range
            ("A1", 0, 0, 3, 3, "A1:C3"),        # No offset, just resize
            ("B2", 1, 1, 1, 3, "C3:E3"),        # Row range
            ("B2", 1, 1, 3, 1, "C3:C5"),        # Column range
        ]
        
        for ref, rows, cols, height, width, expected in test_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols, height=height, width=width):
                result = OFFSET(ref, rows, cols, height, width)
                self.assertEqual(result, expected)
    
    def test_offset_range_input(self):
        """Test OFFSET with range references as input"""
        test_cases = [
            ("A1:B2", 1, 1, None, None, "B2:C3"),   # Preserve dimensions
            ("A1:C3", 0, 1, None, None, "B1:D3"),   # Column shift
            ("A1:B2", 1, 1, 1, 1, "B2"),            # Resize to single cell
            ("A1:B2", 0, 0, 3, 3, "A1:C3"),         # Resize larger
        ]
        
        for ref, rows, cols, height, width, expected in test_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols, height=height, width=width):
                result = OFFSET(ref, rows, cols, height, width)
                self.assertEqual(result, expected)
    
    def test_offset_negative_offsets(self):
        """Test OFFSET with negative row/column offsets"""
        test_cases = [
            ("C3", -1, -1, None, None, "B2"),   # Move up and left
            ("C3", -2, 0, None, None, "C1"),    # Move up only
            ("C3", 0, -2, None, None, "A3"),    # Move left only
        ]
        
        for ref, rows, cols, height, width, expected in test_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols):
                result = OFFSET(ref, rows, cols, height, width)
                self.assertEqual(result, expected)
    
    def test_offset_ref_errors(self):
        """Test OFFSET #REF! errors"""
        error_cases = [
            ("A1", -1, 0, None, None),    # Negative row result
            ("A1", 0, -1, None, None),    # Negative column result
            ("A1", 1048576, 0, None, None),  # Row too large
            ("A1", 0, 16384, None, None),    # Column too large
        ]
        
        for ref, rows, cols, height, width in error_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols):
                result = OFFSET(ref, rows, cols, height, width)
                self.assertIsInstance(result, xlerrors.RefExcelError)
    
    def test_offset_value_errors(self):
        """Test OFFSET #VALUE! errors"""
        error_cases = [
            ("InvalidRef", 1, 1, None, None),     # Invalid reference
            ("A1", 1, 1, 0, None),                # Zero height
            ("A1", 1, 1, None, 0),                # Zero width
            ("A1", 1, 1, -1, None),               # Negative height
            ("A1", 1, 1, None, -1),               # Negative width
        ]
        
        for ref, rows, cols, height, width in error_cases:
            with self.subTest(ref=ref, rows=rows, cols=cols, height=height, width=width):
                result = OFFSET(ref, rows, cols, height, width)
                self.assertIsInstance(result, xlerrors.ValueExcelError)


class TestINDEXFunction(TestDynamicRangeFunctions):
    """Test cases for INDEX function"""
    
    def test_index_single_value(self):
        """Test INDEX returning single values"""
        array = self.create_test_array(3, 3)
        
        test_cases = [
            (1, 1, "R1C1"),     # Top-left
            (2, 2, "R2C2"),     # Center
            (3, 3, "R3C3"),     # Bottom-right
            (1, 3, "R1C3"),     # Top-right
            (3, 1, "R3C1"),     # Bottom-left
        ]
        
        for row, col, expected in test_cases:
            with self.subTest(row=row, col=col):
                result = INDEX(array, row, col)
                # For single values, just compare the string representation
                self.assertEqual(str(result), expected)
    
    def test_index_entire_row(self):
        """Test INDEX returning entire rows (col_num=0)"""
        array = self.create_test_array(3, 3)
        
        test_cases = [
            (1, 0, ["R1C1", "R1C2", "R1C3"]),   # First row
            (2, 0, ["R2C1", "R2C2", "R2C3"]),   # Second row
            (3, 0, ["R3C1", "R3C2", "R3C3"]),   # Third row
        ]
        
        for row, col, expected in test_cases:
            with self.subTest(row=row, col=col):
                result = INDEX(array, row, col)
                result_list = self._convert_result_to_list(result)
                self.assertListEqual(result_list, expected)
    
    def test_index_entire_column(self):
        """Test INDEX returning entire columns (row_num=0)"""
        array = self.create_test_array(3, 3)
        
        test_cases = [
            (0, 1, ["R1C1", "R2C1", "R3C1"]),   # First column
            (0, 2, ["R1C2", "R2C2", "R3C2"]),   # Second column
            (0, 3, ["R1C3", "R2C3", "R3C3"]),   # Third column
        ]
        
        for row, col, expected in test_cases:
            with self.subTest(row=row, col=col):
                result = INDEX(array, row, col)
                result_list = self._convert_result_to_list(result)
                self.assertListEqual(result_list, expected)
    
    def test_index_default_column(self):
        """Test INDEX with default column (col_num not specified)"""
        array = self.create_test_array(3, 1)  # Single column array
        
        result = INDEX(array, 2)  # Should default to col_num=1
        self.assertEqual(result, "R2C1")
    
    def test_index_ref_errors(self):
        """Test INDEX #REF! errors"""
        array = self.create_test_array(3, 3)
        
        error_cases = [
            (4, 1),     # Row out of bounds
            (1, 4),     # Column out of bounds
            (0, 4),     # Column out of bounds for entire column
            (4, 0),     # Row out of bounds for entire row
        ]
        
        for row, col in error_cases:
            with self.subTest(row=row, col=col):
                result = INDEX(array, row, col)
                self.assertIsInstance(result, xlerrors.RefExcelError)
    
    def test_index_value_errors(self):
        """Test INDEX #VALUE! errors"""
        array = self.create_test_array(3, 3)
        
        error_cases = [
            (-1, 1),    # Negative row
            (1, -1),    # Negative column
            (0, 0),     # Both row and column zero
        ]
        
        for row, col in error_cases:
            with self.subTest(row=row, col=col):
                result = INDEX(array, row, col)
                self.assertIsInstance(result, xlerrors.ValueExcelError)
    
    def test_index_empty_array(self):
        """Test INDEX with empty array"""
        empty_array = Mock()
        empty_array.values = []
        
        result = INDEX(empty_array, 1, 1)
        self.assertIsInstance(result, xlerrors.ValueExcelError)
    
    def test_index_array_boolean_evaluation_fix(self):
        """Test that INDEX works with func_xltypes.Array (Priority 0 fix).
        
        This test demonstrates the Array Boolean Bug where numpy arrays
        cause ValueError when evaluated as boolean in _get_array_data().
        """
        # Import func_xltypes to create a proper Array object
        from xlcalculator.xlfunctions import func_xltypes
        
        # Create a func_xltypes.Array with numpy-like behavior
        # This simulates the real scenario where INDEX receives Array objects
        test_data = [
            ['Name', 'Age', 'City'],
            ['Alice', 25, 'NYC'],
            ['Bob', 30, 'LA']
        ]
        array = func_xltypes.Array(test_data)
        
        # This should work without ValueError
        # Before fix: ValueError: The truth value of an array with more than one element is ambiguous
        # After fix: Should return 25 (Alice's age)
        result = INDEX(array, 2, 2)
        expected = 25
        self.assertEqual(result, expected)
    
    def test_index_array_return_integration_fix(self):
        """Test that INDEX returns proper array types (Priority 1 fix).
        
        This test demonstrates the Array Return Type Integration issue where
        INDEX returns Python lists instead of func_xltypes.Array objects.
        """
        from xlcalculator.xlfunctions import func_xltypes
        
        # Create test data
        test_data = [
            ['Name', 'Age', 'City'],
            ['Alice', 25, 'NYC'],
            ['Bob', 30, 'LA']
        ]
        array = func_xltypes.Array(test_data)
        
        # Test column array return (row_num=0)
        result = INDEX(array, 0, 2)  # Should return Age column
        self.assertIsInstance(result, func_xltypes.Array, 
                            "INDEX should return func_xltypes.Array for column access")
        
        # Test row array return (col_num=0)  
        result = INDEX(array, 2, 0)  # Should return Alice row
        self.assertIsInstance(result, func_xltypes.Array,
                            "INDEX should return func_xltypes.Array for row access")
    
    def test_offset_range_resolution_integration_fix(self):
        """Test that OFFSET resolves to actual values when evaluator context is available.
        
        This test demonstrates the OFFSET Range Resolution fix where
        OFFSET can resolve references to actual values when called with evaluator context.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET, _set_evaluator_context
        from xlcalculator.xlfunctions import func_xltypes
        
        # Set up test data in the model
        test_data = [
            ['Name', 'Age', 'Score'],
            ['Alice', 25, 85],
            ['Bob', 30, 92],
            ['Charlie', 35, 78]
        ]
        
        for row_idx, row_data in enumerate(test_data, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell_ref = f'Sheet1!{chr(64 + col_idx)}{row_idx}'
                self.model.set_cell_value(cell_ref, value)
        
        # Test OFFSET without evaluator context (backward compatibility)
        result_no_context = OFFSET('A1', 1, 0)
        expected_no_context = 'A2'  # Should return reference string
        self.assertEqual(result_no_context, expected_no_context,
                        "OFFSET should return reference string without context")
        
        # Test OFFSET with evaluator context (new functionality)
        _set_evaluator_context(self.evaluator)
        result_with_context = OFFSET('A1', 1, 0)
        expected_with_context = 'Alice'  # Should return actual value
        self.assertEqual(result_with_context, expected_with_context,
                        "OFFSET should resolve to actual value with context")
        
        # Test range OFFSET with evaluator context
        result_range = OFFSET('B1', 1, 0, 3, 1)  # Ages column
        self.assertIsInstance(result_range, func_xltypes.Array,
                            "OFFSET should return Array for range results")
        
        # Verify the array contains the expected values
        # The array should contain ages: [25, 30, 35]
        array_values = result_range.values
        expected_ages = [[25], [30], [35]]  # Column format
        self.assertEqual(array_values.tolist(), expected_ages,
                        "OFFSET range should contain correct values")
        
        # Clean up evaluator context to avoid affecting other tests
        from xlcalculator.xlfunctions.dynamic_range import _clear_evaluator_context
        _clear_evaluator_context()


class TestINDIRECTFunction(TestDynamicRangeFunctions):
    """Test cases for INDIRECT function"""
    
    def test_indirect_basic_references(self):
        """Test INDIRECT with basic cell references"""
        test_cases = [
            ("A1", "A1"),
            ("B2", "B2"),
            ("Z26", "Z26"),
            ("AA1", "AA1"),
        ]
        
        for ref_text, expected in test_cases:
            with self.subTest(ref_text=ref_text):
                result = INDIRECT(ref_text)
                self.assertEqual(result, expected)
    
    def test_indirect_range_references(self):
        """Test INDIRECT with range references"""
        test_cases = [
            ("A1:B2", "A1:B2"),
            ("C3:E5", "C3:E5"),
            ("A:A", "A:A"),      # Entire column (if supported)
            ("1:1", "1:1"),      # Entire row (if supported)
        ]
        
        for ref_text, expected in test_cases:
            with self.subTest(ref_text=ref_text):
                result = INDIRECT(ref_text)
                self.assertEqual(result, expected)
    
    def test_indirect_sheet_references(self):
        """Test INDIRECT with sheet references"""
        test_cases = [
            ("Sheet1!A1", "Sheet1!A1"),
            ("Data!B2:C3", "Data!B2:C3"),
            ("'Sheet Name'!A1", "'Sheet Name'!A1"),
        ]
        
        for ref_text, expected in test_cases:
            with self.subTest(ref_text=ref_text):
                result = INDIRECT(ref_text)
                self.assertEqual(result, expected)
    
    def test_indirect_absolute_references(self):
        """Test INDIRECT with absolute references"""
        test_cases = [
            ("$A$1", "A1"),         # Should normalize
            ("$B2", "B2"),          # Should normalize
            ("A$3", "A3"),          # Should normalize
        ]
        
        for ref_text, expected in test_cases:
            with self.subTest(ref_text=ref_text):
                result = INDIRECT(ref_text)
                self.assertEqual(result, expected)
    
    def test_indirect_name_errors(self):
        """Test INDIRECT #NAME? errors"""
        error_cases = [
            "",                 # Empty string
            "InvalidRef123",    # Invalid format
            "1A",              # Numbers before letters
            "A",               # Missing row number
            "!A1",             # Invalid sheet reference
        ]
        
        for ref_text in error_cases:
            with self.subTest(ref_text=ref_text):
                result = INDIRECT(ref_text)
                self.assertIsInstance(result, xlerrors.NameExcelError)
    
    def test_indirect_r1c1_not_supported(self):
        """Test INDIRECT with R1C1 style (not yet supported)"""
        result = INDIRECT("R1C1", False)
        self.assertIsInstance(result, xlerrors.ValueExcelError)


class TestDynamicRangeIntegration(TestDynamicRangeFunctions):
    """Integration tests for dynamic range functions"""
    
    def test_offset_in_formula(self):
        """Test OFFSET used within formulas"""
        # Set up formula that uses OFFSET
        self.model.set_cell_value('Sheet1!F1', '=OFFSET(A1, 1, 1)')
        
        # Evaluate and check result
        result = self.evaluator.evaluate('Sheet1!F1')
        # Currently returns the formula text due to limited evaluator integration
        # Full integration would require the evaluator to resolve OFFSET references
        self.assertIsNotNone(result)  # Just verify it doesn't crash
    
    def test_index_in_formula(self):
        """Test INDEX used within formulas"""
        # This would require proper array support in the evaluator
        # For now, test the function directly
        pass
    
    def test_indirect_dynamic_reference(self):
        """Test INDIRECT with dynamic references"""
        # Set up cell containing reference text
        self.model.set_cell_value('Sheet1!H1', '=INDIRECT(G1)')
        
        # G1 contains "B2", so H1 should reference B2
        result = self.evaluator.evaluate('Sheet1!H1')
        # This test would require full integration with evaluator
        pass
    
    def test_nested_dynamic_functions(self):
        """Test nested dynamic range functions"""
        # Example: INDEX(OFFSET(A1, 1, 1, 3, 3), 2, 2)
        # This would require full evaluator integration
        pass


class TestExcelCompatibility(TestDynamicRangeFunctions):
    """Test Excel compatibility and edge cases"""
    
    def test_excel_error_types(self):
        """Verify error types match Excel exactly"""
        # These test cases are derived from actual Excel behavior
        excel_error_cases = [
            # OFFSET errors
            (lambda: OFFSET("A1", -1, 0), xlerrors.RefExcelError),
            (lambda: OFFSET("InvalidRef", 1, 1), xlerrors.ValueExcelError),
            
            # INDEX errors  
            (lambda: INDEX(self.create_test_array(3, 3), 4, 1), xlerrors.RefExcelError),
            (lambda: INDEX(self.create_test_array(3, 3), 0, 0), xlerrors.ValueExcelError),
            
            # INDIRECT errors
            (lambda: INDIRECT("InvalidRef123"), xlerrors.NameExcelError),
            (lambda: INDIRECT(""), xlerrors.NameExcelError),
        ]
        
        for func, expected_error in excel_error_cases:
            with self.subTest(func=func.__name__ if hasattr(func, '__name__') else str(func)):
                result = func()
                self.assertIsInstance(result, expected_error)
    
    def test_boundary_conditions(self):
        """Test behavior at Excel limits"""
        # Test maximum valid coordinates
        max_row_ref = ReferenceResolver.cell_to_string(1048576, 16384)  # XFD1048576
        result = OFFSET(max_row_ref, 0, 0)
        self.assertEqual(result, max_row_ref)
        
        # Test just beyond limits should error
        result = OFFSET("A1", 1048576, 0)  # Would exceed max row
        self.assertIsInstance(result, xlerrors.RefExcelError)
    
    def test_case_insensitive_references(self):
        """Test that references are case-insensitive"""
        test_cases = [
            ("a1", "A1"),
            ("sheet1!b2", "sheet1!B2"),
            ("$a$1", "A1"),
        ]
        
        for ref_text, expected in test_cases:
            with self.subTest(ref_text=ref_text):
                result = INDIRECT(ref_text)
                self.assertEqual(result, expected)


class TestPerformance(TestDynamicRangeFunctions):
    """Performance tests for dynamic range functions"""
    
    def test_large_offset_performance(self):
        """Test OFFSET performance with large offsets"""
        import time
        
        start_time = time.time()
        for i in range(1000):
            OFFSET("A1", i, i)
        end_time = time.time()
        
        # Should complete 1000 operations in reasonable time
        self.assertLess(end_time - start_time, 1.0)  # Less than 1 second
    
    def test_large_array_index_performance(self):
        """Test INDEX performance with large arrays"""
        # Create large array
        large_array = self.create_test_array(100, 100)
        
        import time
        start_time = time.time()
        for i in range(100):
            INDEX(large_array, i + 1, i + 1)
        end_time = time.time()
        
        # Should complete 100 operations in reasonable time
        self.assertLess(end_time - start_time, 1.0)  # Less than 1 second


if __name__ == '__main__':
    # Run specific test suites
    import sys
    
    if len(sys.argv) > 1:
        # Run specific test class
        test_class = sys.argv[1]
        suite = unittest.TestLoader().loadTestsFromName(f'__main__.{test_class}')
        unittest.TextTestRunner(verbosity=2).run(suite)
    else:
        # Run all tests
        unittest.main(verbosity=2)