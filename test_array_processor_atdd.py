#!/usr/bin/env python3
"""
ATDD Test for ArrayProcessor.extract_array_data() Bug Fix

RED PHASE: Create failing acceptance tests that demonstrate the expected behavior
for array data extraction that is currently broken.

Test Cases Based on Failure Analysis:
1. String range references should extract actual 2D array data
2. pandas DataFrames should be converted to 2D lists
3. Direct arrays should be preserved as 2D format
4. Range objects should be evaluated to 2D arrays

Expected Behavior (from original working code):
- 'Data!A1:E6' → [[1,2,3,4,5], [6,7,8,9,10], ...] (actual cell values)
- pandas DataFrame → 2D list via .values.tolist()
- [1,2,3] → [[1,2,3]] (ensure 2D format)
- Range object → evaluated 2D array
"""

import unittest
from unittest.mock import Mock, MagicMock
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xlcalculator.utils.arrays import ArrayProcessor


class ArrayProcessorATDDTest(unittest.TestCase):
    """ATDD tests for ArrayProcessor.extract_array_data() bug fix."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.mock_evaluator = Mock()
        
        # Mock evaluator.get_range_values() to return realistic 2D array
        self.mock_evaluator.get_range_values.return_value = [
            [1, 2, 3, 4, 5],
            [6, 7, 8, 9, 10],
            [11, 12, 13, 14, 15],
            [16, 17, 18, 19, 20],
            [21, 22, 23, 24, 25],
            [26, 27, 28, 29, 30]
        ]
        
        # Mock evaluator.evaluate() for range objects
        self.mock_evaluator.evaluate.return_value = [
            ['A', 'B', 'C'],
            ['D', 'E', 'F']
        ]
    
    def test_string_range_reference_extraction(self):
        """
        ACCEPTANCE TEST 1: String range references should extract actual 2D array data
        
        GIVEN: A string range reference like 'Data!A1:E6'
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should return actual 2D array of cell values, not wrapped string
        """
        # GIVEN
        range_string = 'Data!A1:E6'
        
        # WHEN
        result = ArrayProcessor.extract_array_data(range_string, self.mock_evaluator)
        
        # THEN
        self.assertIsInstance(result, list, "Should return a list")
        self.assertIsInstance(result[0], list, "Should return 2D array (list of lists)")
        self.assertEqual(len(result), 6, "Should have 6 rows")
        self.assertEqual(len(result[0]), 5, "Should have 5 columns")
        self.assertEqual(result[0][0], 1, "First cell should be actual value, not string")
        self.assertNotEqual(result, [['Data!A1:E6']], "Should NOT wrap string as single value")
        
        # Verify correct evaluator method was called
        self.mock_evaluator.get_range_values.assert_called_once_with('Data!A1:E6')
    
    def test_pandas_dataframe_extraction(self):
        """
        ACCEPTANCE TEST 2: pandas DataFrames should be converted to 2D lists
        
        GIVEN: A pandas DataFrame-like object with .values attribute
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should convert to 2D list using .values.tolist()
        """
        # GIVEN - Mock pandas DataFrame
        mock_dataframe = Mock()
        mock_dataframe.values.tolist.return_value = [
            ['Alice', 25, 'Engineer'],
            ['Bob', 30, 'Manager'],
            ['Carol', 28, 'Designer']
        ]
        
        # WHEN
        result = ArrayProcessor.extract_array_data(mock_dataframe, self.mock_evaluator)
        
        # THEN
        expected = [
            ['Alice', 25, 'Engineer'],
            ['Bob', 30, 'Manager'],
            ['Carol', 28, 'Designer']
        ]
        self.assertEqual(result, expected, "Should convert DataFrame to 2D list")
        mock_dataframe.values.tolist.assert_called_once()
    
    def test_direct_array_preservation(self):
        """
        ACCEPTANCE TEST 3: Direct arrays should be preserved in 2D format
        
        GIVEN: A direct list or tuple array
        WHEN: ArrayProcessor.extract_array_data() is called  
        THEN: Should ensure 2D format without data loss
        """
        # GIVEN - 1D array should become 2D
        array_1d = [1, 2, 3, 4, 5]
        
        # WHEN
        result = ArrayProcessor.extract_array_data(array_1d, self.mock_evaluator)
        
        # THEN
        self.assertEqual(result, [[1, 2, 3, 4, 5]], "1D array should become 2D")
        
        # GIVEN - 2D array should be preserved
        array_2d = [[1, 2], [3, 4], [5, 6]]
        
        # WHEN
        result = ArrayProcessor.extract_array_data(array_2d, self.mock_evaluator)
        
        # THEN
        self.assertEqual(result, [[1, 2], [3, 4], [5, 6]], "2D array should be preserved")
    
    def test_range_object_evaluation(self):
        """
        ACCEPTANCE TEST 4: Range objects should be evaluated to 2D arrays
        
        GIVEN: A range object with .address attribute (but NOT .values like pandas)
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should evaluate the range object to get 2D array
        """
        # GIVEN - Range object with address attribute only (realistic)
        mock_range = Mock()
        mock_range.address = 'Sheet1!A1:C2'
        # Ensure it doesn't have 'values' attribute like pandas
        del mock_range.values
        
        # WHEN
        result = ArrayProcessor.extract_array_data(mock_range, self.mock_evaluator)
        
        # THEN
        expected = [['A', 'B', 'C'], ['D', 'E', 'F']]
        self.assertEqual(result, expected, "Range object should be evaluated")
        self.mock_evaluator.evaluate.assert_called_with(mock_range)
    
    def test_single_value_handling(self):
        """
        ACCEPTANCE TEST 5: Single values should be wrapped as 2D arrays
        
        GIVEN: A single scalar value
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should wrap as 2D array [[value]]
        """
        # GIVEN
        single_value = 42
        
        # WHEN
        result = ArrayProcessor.extract_array_data(single_value, self.mock_evaluator)
        
        # THEN
        self.assertEqual(result, [[42]], "Single value should be wrapped as 2D array")
    
    def test_integration_with_index_function(self):
        """
        ACCEPTANCE TEST 6: Integration test with INDEX function pattern
        
        GIVEN: The exact pattern used by INDEX function
        WHEN: ArrayProcessor.extract_array_data() is called with string range
        THEN: INDEX should be able to access array_data[row][col] successfully
        """
        # GIVEN - Simulate INDEX function call pattern
        range_reference = 'Data!A1:E6'
        
        # WHEN
        array_data = ArrayProcessor.extract_array_data(range_reference, self.mock_evaluator)
        
        # THEN - Should be able to access like INDEX function does
        try:
            # INDEX(Data!A1:E6, 2, 2) should access row 1, col 1 (0-based)
            value = array_data[1][1]  # Row 2, Column 2 in Excel (1-based)
            self.assertEqual(value, 7, "Should access correct cell value")
            
            # INDEX(Data!A1:E6, 1, 5) should access row 0, col 4 (0-based)  
            value = array_data[0][4]  # Row 1, Column 5 in Excel (1-based)
            self.assertEqual(value, 5, "Should access correct cell value")
            
        except (IndexError, TypeError) as e:
            self.fail(f"INDEX function pattern should work without errors: {e}")
    
    def test_empty_string_handling(self):
        """
        EDGE CASE TEST 1: Empty strings should be handled gracefully
        
        GIVEN: An empty string reference
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should handle gracefully (likely return empty array or raise appropriate error)
        """
        # GIVEN
        empty_string = ''
        self.mock_evaluator.get_range_values.return_value = []
        
        # WHEN
        result = ArrayProcessor.extract_array_data(empty_string, self.mock_evaluator)
        
        # THEN
        self.assertEqual(result, [], "Empty string should return empty array")
        self.mock_evaluator.get_range_values.assert_called_with('')
    
    def test_none_value_handling(self):
        """
        EDGE CASE TEST 2: None values should be handled gracefully
        
        GIVEN: A None reference
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should wrap None as single value [[None]]
        """
        # GIVEN
        none_value = None
        
        # WHEN
        result = ArrayProcessor.extract_array_data(none_value, self.mock_evaluator)
        
        # THEN
        self.assertEqual(result, [[None]], "None should be wrapped as single value")
    
    def test_empty_list_handling(self):
        """
        EDGE CASE TEST 3: Empty lists should be handled gracefully
        
        GIVEN: An empty list
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should return empty 2D array [[]]
        """
        # GIVEN
        empty_list = []
        
        # WHEN
        result = ArrayProcessor.extract_array_data(empty_list, self.mock_evaluator)
        
        # THEN
        self.assertEqual(result, [[]], "Empty list should return [[]]")
    
    def test_mixed_type_array_handling(self):
        """
        EDGE CASE TEST 4: Arrays with mixed types should be preserved
        
        GIVEN: An array with mixed data types
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should preserve all data types in 2D format
        """
        # GIVEN
        mixed_array = [1, 'text', 3.14, True, None]
        
        # WHEN
        result = ArrayProcessor.extract_array_data(mixed_array, self.mock_evaluator)
        
        # THEN
        expected = [[1, 'text', 3.14, True, None]]
        self.assertEqual(result, expected, "Mixed types should be preserved")
    
    def test_nested_list_preservation(self):
        """
        EDGE CASE TEST 5: Nested lists should be preserved as-is
        
        GIVEN: A properly formatted 2D array
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should preserve the exact structure
        """
        # GIVEN
        nested_array = [
            [1, 2, 3],
            ['a', 'b', 'c'],
            [True, False, None]
        ]
        
        # WHEN
        result = ArrayProcessor.extract_array_data(nested_array, self.mock_evaluator)
        
        # THEN
        self.assertEqual(result, nested_array, "Nested structure should be preserved")
    
    def test_evaluator_error_propagation(self):
        """
        EDGE CASE TEST 6: Evaluator errors should be propagated
        
        GIVEN: An evaluator that raises an exception
        WHEN: ArrayProcessor.extract_array_data() is called with string
        THEN: Should propagate the evaluator exception
        """
        # GIVEN
        range_string = 'Invalid!Range'
        self.mock_evaluator.get_range_values.side_effect = Exception("Invalid range")
        
        # WHEN/THEN
        with self.assertRaises(Exception) as context:
            ArrayProcessor.extract_array_data(range_string, self.mock_evaluator)
        
        self.assertIn("Invalid range", str(context.exception))
    
    def test_numeric_types_handling(self):
        """
        EDGE CASE TEST 7: Various numeric types should be handled
        
        GIVEN: Different numeric types (int, float, complex)
        WHEN: ArrayProcessor.extract_array_data() is called
        THEN: Should wrap each as single value 2D array
        """
        # Test cases
        test_cases = [
            (42, [[42]]),
            (3.14159, [[3.14159]]),
            (1+2j, [[1+2j]]),
            (0, [[0]]),
            (-5, [[-5]])
        ]
        
        for input_value, expected in test_cases:
            with self.subTest(input_value=input_value):
                result = ArrayProcessor.extract_array_data(input_value, self.mock_evaluator)
                self.assertEqual(result, expected, f"Numeric type {type(input_value)} should be wrapped")


if __name__ == '__main__':
    print("🟢 GREEN PHASE: Running acceptance tests for ArrayProcessor.extract_array_data()")
    print("Expected: All tests should PASS after implementing fixes")
    print()
    
    unittest.main(verbosity=2)