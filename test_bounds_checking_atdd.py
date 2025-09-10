"""
ATDD Test Suite for Category 2: Bounds Checking Issues

This test suite focuses on the specific bounds checking failures identified in Category 2,
particularly the INDEX function incorrectly reporting "Row index out of range" when
accessing valid array positions.

Key Issue: INDEX function bounds validation is incorrectly rejecting valid array indices
"""

import unittest
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text
from xlcalculator.utils.validation import validate_array_bounds


class BoundsCheckingATDDTest(unittest.TestCase):
    """ATDD tests for bounds checking issues in INDEX function."""
    
    def test_array_bounds_validation_logic(self):
        """
        FAILING TEST: Array bounds validation should accept valid indices.
        
        This tests the core bounds checking logic that's causing INDEX to fail.
        The issue is likely in validate_array_bounds() function.
        """
        # Create test array data that matches the failing pattern
        # 6 rows x 5 columns array (Data!A1:E6)
        array_data = [
            ['Name', 'Age', 'City', 'Score', 'Active'],      # Row 1
            ['Alice', 25, 'NYC', 85, True],                  # Row 2  
            ['Bob', 30, 'LA', 92, False],                    # Row 3
            ['Charlie', 35, 'Chicago', 78, True],            # Row 4
            ['Diana', 28, 'Miami', 95, True],                # Row 5
            ['Eve', 22, 'Boston', 88, False]                 # Row 6
        ]
        
        # Test case that should work: INDEX(array, 2, 2) = 25
        # This means row_idx=1, col_idx=1 in 0-based indexing
        row_idx = 1  # Row 2 in 1-based = index 1 in 0-based
        col_idx = 1  # Column 2 in 1-based = index 1 in 0-based
        
        # This should NOT raise an exception - the bounds are valid
        try:
            validate_array_bounds(array_data, row_idx, col_idx)
            # If we get here, bounds validation passed correctly
            actual_value = array_data[row_idx][col_idx]
            self.assertEqual(25, actual_value, "Array access should return correct value")
        except xlerrors.RefExcelError as e:
            # This is the bug - bounds validation is incorrectly rejecting valid indices
            self.fail(f"Bounds validation incorrectly rejected valid indices (1,1): {e}")
    
    def test_bounds_validation_edge_cases(self):
        """
        Test edge cases for bounds validation to identify the specific issue.
        """
        # Same test array: 6 rows x 5 columns
        array_data = [
            ['Name', 'Age', 'City', 'Score', 'Active'],
            ['Alice', 25, 'NYC', 85, True],
            ['Bob', 30, 'LA', 92, False],
            ['Charlie', 35, 'Chicago', 78, True],
            ['Diana', 28, 'Miami', 95, True],
            ['Eve', 22, 'Boston', 88, False]
        ]
        
        # Test valid bounds - these should all pass
        valid_cases = [
            (0, 0, "First cell"),           # array_data[0][0] = 'Name'
            (1, 1, "Target cell"),          # array_data[1][1] = 25
            (5, 4, "Last cell"),            # array_data[5][4] = False
            (0, 4, "First row, last col"),  # array_data[0][4] = 'Active'
            (5, 0, "Last row, first col")   # array_data[5][0] = 'Eve'
        ]
        
        for row_idx, col_idx, description in valid_cases:
            with self.subTest(case=description, row=row_idx, col=col_idx):
                try:
                    validate_array_bounds(array_data, row_idx, col_idx)
                    # Validation passed - verify we can access the data
                    value = array_data[row_idx][col_idx]
                    self.assertIsNotNone(value, f"Should be able to access {description}")
                except xlerrors.RefExcelError as e:
                    self.fail(f"Valid bounds rejected for {description}: {e}")
        
        # Test invalid bounds - these should fail
        invalid_cases = [
            (-1, 0, "Negative row"),
            (0, -1, "Negative column"),
            (6, 0, "Row beyond array"),      # Only 6 rows (0-5)
            (0, 5, "Column beyond array"),   # Only 5 columns (0-4)
            (10, 10, "Both beyond array")
        ]
        
        for row_idx, col_idx, description in invalid_cases:
            with self.subTest(case=description, row=row_idx, col=col_idx):
                with self.assertRaises(xlerrors.RefExcelError, 
                                     msg=f"Invalid bounds should be rejected for {description}"):
                    validate_array_bounds(array_data, row_idx, col_idx)
    
    def test_array_processor_extract_array_data_bounds(self):
        """
        Test that ArrayProcessor.extract_array_data returns correct dimensions.
        
        This tests if the issue is in how array data is extracted and passed to bounds validation.
        """
        from xlcalculator.utils.arrays import ArrayProcessor
        
        # Test with a mock evaluator that returns known array data
        class MockEvaluator:
            def get_range_values(self, range_ref):
                # Return the same test data structure
                return [
                    ['Name', 'Age', 'City', 'Score', 'Active'],
                    ['Alice', 25, 'NYC', 85, True],
                    ['Bob', 30, 'LA', 92, False],
                    ['Charlie', 35, 'Chicago', 78, True],
                    ['Diana', 28, 'Miami', 95, True],
                    ['Eve', 22, 'Boston', 88, False]
                ]
        
        mock_evaluator = MockEvaluator()
        
        # Test extracting array data from string range reference
        array_data = ArrayProcessor.extract_array_data("Data!A1:E6", mock_evaluator)
        
        # Verify dimensions
        self.assertEqual(6, len(array_data), "Should have 6 rows")
        self.assertEqual(5, len(array_data[0]), "Should have 5 columns")
        
        # Verify specific values
        self.assertEqual('Name', array_data[0][0], "First cell should be 'Name'")
        self.assertEqual(25, array_data[1][1], "Target cell should be 25")
        self.assertEqual(False, array_data[5][4], "Last cell should be False")
        
        # Test bounds validation on this extracted data
        # This should work: row_idx=1, col_idx=1 (0-based for row 2, col 2)
        try:
            validate_array_bounds(array_data, 1, 1)
        except xlerrors.RefExcelError as e:
            self.fail(f"Bounds validation failed on extracted array data: {e}")
    
    def test_indirect_function_root_cause(self):
        """
        FAILING TEST: Identify why INDIRECT returns BLANK instead of range data.
        
        This is the actual root cause of the "bounds checking" issue.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context with evaluator
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [
                    ['Name', 'Age', 'City', 'Score', 'Active'],
                    ['Alice', 25, 'NYC', 85, True],
                    ['Bob', 30, 'LA', 92, False],
                    ['Charlie', 35, 'Chicago', 78, True],
                    ['Diana', 28, 'Miami', 95, True],
                    ['Eve', 22, 'Boston', 88, False]
                ]
            
            def evaluate(self, reference):
                # Mock evaluate method
                return "mock_result"
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"  # Current sheet context
        
        mock_context = MockContext()
        
        # Test INDIRECT function directly
        # INDIRECT("Data!A1:E6") should return range data, not BLANK
        try:
            result = INDIRECT("Data!A1:E6", _context=mock_context)
            
            # This should NOT be BLANK
            self.assertIsNotNone(result, "INDIRECT should not return None")
            
            # Check if it's the expected type
            if hasattr(result, 'values'):
                # It's an Array type
                array_data = result.values
                self.assertEqual(6, len(array_data), "Should have 6 rows")
                self.assertEqual(5, len(array_data[0]), "Should have 5 columns")
                self.assertEqual(25, array_data[1][1], "Should contain expected value")
            else:
                # It might be raw array data
                self.assertIsInstance(result, list, "Should return array data")
                self.assertEqual(6, len(result), "Should have 6 rows")
                self.assertEqual(25, result[1][1], "Should contain expected value")
                
        except Exception as e:
            # This helps us identify what's actually failing in INDIRECT
            self.fail(f"INDIRECT function failed: {e}")
    
    def test_index_with_indirect_integration(self):
        """
        Test the full INDEX(INDIRECT(...)) integration to reproduce the exact failure.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDEX, INDIRECT
        
        # Create a mock context with evaluator
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [
                    ['Name', 'Age', 'City', 'Score', 'Active'],
                    ['Alice', 25, 'NYC', 85, True],
                    ['Bob', 30, 'LA', 92, False],
                    ['Charlie', 35, 'Chicago', 78, True],
                    ['Diana', 28, 'Miami', 95, True],
                    ['Eve', 22, 'Boston', 88, False]
                ]
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Step 1: Test INDIRECT alone
        indirect_result = INDIRECT("Data!A1:E6", _context=mock_context)
        
        # Step 2: Test INDEX with INDIRECT result
        # This should reproduce the "Row index out of range" error
        try:
            final_result = INDEX(indirect_result, 2, 2, _context=mock_context)
            self.assertEqual(25, final_result, "INDEX(INDIRECT(...), 2, 2) should return 25")
        except xlerrors.RefExcelError as e:
            # This is the actual bug - INDEX receiving wrong data from INDIRECT
            self.fail(f"INDEX failed with INDIRECT result: {e}. INDIRECT returned: {indirect_result}")


    def test_indirect_edge_cases(self):
        """
        Test INDIRECT function edge cases to ensure robust bounds checking.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        from xlcalculator.xlfunctions import xlerrors
        
        # Create a mock context with evaluator
        class MockEvaluator:
            def get_range_values(self, range_ref):
                if range_ref == "Data!A1:E6":
                    return [
                        ['Name', 'Age', 'City', 'Score', 'Active'],
                        ['Alice', 25, 'NYC', 85, True],
                        ['Bob', 30, 'LA', 92, False],
                        ['Charlie', 35, 'Chicago', 78, True],
                        ['Diana', 28, 'Miami', 95, True],
                        ['Eve', 22, 'Boston', 88, False]
                    ]
                elif range_ref == "Data!A1:B2":
                    return [['Name', 'Age'], ['Alice', 25]]
                else:
                    raise Exception(f"Unknown range: {range_ref}")
            
            def evaluate(self, reference):
                # Mock single cell evaluation
                if reference == "Data!A1":
                    return "Name"
                return "mock_cell_value"
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test different range sizes
        test_cases = [
            ("Data!A1:E6", 6, 5, "6x5 range"),
            ("Data!A1:B2", 2, 2, "2x2 range"),
        ]
        
        for range_ref, expected_rows, expected_cols, description in test_cases:
            with self.subTest(range_ref=range_ref):
                result = INDIRECT(range_ref, _context=mock_context)
                self.assertIsInstance(result, Array, f"{description} should return Array")
                
                array_data = result.values
                self.assertEqual(expected_rows, len(array_data), f"{description} should have {expected_rows} rows")
                self.assertEqual(expected_cols, len(array_data[0]), f"{description} should have {expected_cols} columns")
        
        # Test invalid references
        with self.assertRaises(xlerrors.RefExcelError):
            INDIRECT("InvalidRange!A1:Z99", _context=mock_context)
    
    def test_index_with_indirect_edge_cases(self):
        """
        Test INDEX with INDIRECT for various edge cases and boundary conditions.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDEX, INDIRECT
        from xlcalculator.xlfunctions import xlerrors
        
        # Create a mock context with evaluator
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [
                    ['Name', 'Age', 'City', 'Score', 'Active'],
                    ['Alice', 25, 'NYC', 85, True],
                    ['Bob', 30, 'LA', 92, False],
                    ['Charlie', 35, 'Chicago', 78, True],
                    ['Diana', 28, 'Miami', 95, True],
                    ['Eve', 22, 'Boston', 88, False]
                ]
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test valid boundary cases
        boundary_cases = [
            (1, 1, 'Name', "First cell"),
            (2, 2, 25, "Target cell"),
            (6, 5, False, "Last cell"),
            (1, 5, 'Active', "First row, last column"),
            (6, 1, 'Eve', "Last row, first column"),
        ]
        
        for row, col, expected, description in boundary_cases:
            with self.subTest(row=row, col=col, description=description):
                indirect_result = INDIRECT("Data!A1:E6", _context=mock_context)
                result = INDEX(indirect_result, row, col, _context=mock_context)
                self.assertEqual(expected, result, f"{description} should return {expected}")
        
        # Test special cases: row=0 or col=0 (should return arrays)
        special_cases = [
            (0, 2, Array, "Row 0 should return entire column"),
            (2, 0, Array, "Column 0 should return entire row"),
        ]
        
        for row, col, expected_type, description in special_cases:
            with self.subTest(row=row, col=col, description=description):
                indirect_result = INDIRECT("Data!A1:E6", _context=mock_context)
                result = INDEX(indirect_result, row, col, _context=mock_context)
                self.assertIsInstance(result, expected_type, f"{description}")
        
        # Test invalid boundary cases - these should raise errors
        invalid_cases = [
            (0, 0, "Both row and column 0"),  # ValueExcelError
            (7, 1, "Row beyond array"),       # RefExcelError
            (1, 6, "Column beyond array"),    # RefExcelError
            (-1, 1, "Negative row"),          # ValueExcelError
            (1, -1, "Negative column"),       # ValueExcelError
        ]
        
        for row, col, description in invalid_cases:
            with self.subTest(row=row, col=col, description=description):
                indirect_result = INDIRECT("Data!A1:E6", _context=mock_context)
                with self.assertRaises((xlerrors.RefExcelError, xlerrors.ValueExcelError), 
                                     msg=f"{description} should raise an error"):
                    INDEX(indirect_result, row, col, _context=mock_context)
    
    def test_array_processor_consistency(self):
        """
        Test that ArrayProcessor.extract_array_data works consistently with INDIRECT results.
        """
        from xlcalculator.utils.arrays import ArrayProcessor
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context with evaluator
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [
                    ['Name', 'Age', 'City'],
                    ['Alice', 25, 'NYC'],
                    ['Bob', 30, 'LA']
                ]
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        mock_evaluator = MockEvaluator()
        
        # Test that INDIRECT result can be processed by ArrayProcessor
        indirect_result = INDIRECT("Data!A1:C3", _context=mock_context)
        
        # Extract array data from INDIRECT result
        array_data = ArrayProcessor.extract_array_data(indirect_result, mock_evaluator)
        
        # Verify consistency
        self.assertEqual(3, len(array_data), "Should have 3 rows")
        self.assertEqual(3, len(array_data[0]), "Should have 3 columns")
        self.assertEqual('Name', array_data[0][0], "First cell should be 'Name'")
        self.assertEqual(25, array_data[1][1], "Should contain expected value")


if __name__ == '__main__':
    unittest.main()