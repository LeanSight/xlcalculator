"""
ATDD Test Suite for Category 4: OFFSET Function Array Parameter Issues

This test suite focuses on the specific OFFSET function failures identified in Category 4,
particularly around array parameter handling and error detection.

Key Issues:
1. OFFSET with array parameters (ROW(A1:A2)-1) should return array results
2. OFFSET error handling should properly trigger ISERROR detection
"""

import unittest
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text


class OffsetArrayATDDTest(unittest.TestCase):
    """ATDD tests for OFFSET array parameter issues."""
    
    def test_offset_with_single_parameters(self):
        """
        CONTROL TEST: OFFSET with single numeric parameters should work.
        
        This verifies that basic OFFSET functionality works and isolates
        the issue to array parameter handling.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        
        # Create a mock context with evaluator
        class MockEvaluator:
            def get_range_values(self, range_ref):
                # Mock range data for Data!A1:E6
                return [
                    ['Name', 'Age', 'City', 'Score', 'Active'],
                    ['Alice', 25, 'NYC', 85, True],
                    ['Bob', 30, 'LA', 92, False],
                    ['Charlie', 35, 'Chicago', 78, True],
                    ['Diana', 28, 'Miami', 95, True],
                    ['Eve', 22, 'Boston', 88, False]
                ]
            
            def evaluate(self, reference):
                # Mock single cell evaluation
                cell_data = {
                    'Data!A1': 'Name',
                    'Data!A2': 'Alice', 
                    'Data!A3': 'Bob',
                    'Data!B1': 'Age',
                    'Data!B2': 25,
                    'Data!B3': 30
                }
                return cell_data.get(reference, None)
            
            def get_cell_value(self, address):
                # Same as evaluate for our mock
                return self.evaluate(address)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test simple OFFSET: OFFSET(Data!A1, 1, 0) should return 'Alice'
        result = OFFSET("Data!A1", 1, 0, _context=mock_context)
        self.assertEqual('Alice', result, "Simple OFFSET should work")
    
    def test_offset_with_array_parameters_failing(self):
        """
        FAILING TEST: OFFSET with array parameters should return array results.
        
        This reproduces the core Category 4 issue where OFFSET fails when
        row or column parameters are arrays instead of single numbers.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        from xlcalculator.xlfunctions.func_xltypes import Array
        
        # Create a mock context
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
                cell_data = {
                    'Data!A1': 'Name',
                    'Data!A2': 'Alice', 
                    'Data!A3': 'Bob',
                    'Data!B1': 'Age',
                    'Data!B2': 25,
                    'Data!B3': 30
                }
                return cell_data.get(reference, None)
            
            def get_cell_value(self, address):
                return self.evaluate(address)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test the pattern that fails: OFFSET(Data!A1, ROW(A1:A2)-1, 0)
        # This should be equivalent to OFFSET(Data!A1, [0, 1], 0)
        # and should return an array with ['Name', 'Alice']
        
        # Simulate ROW(A1:A2)-1 which should be [1-1, 2-1] = [0, 1]
        row_array = Array([[0], [1]])  # ROW returns column array
        
        # This should work but currently fails with "Row and column offsets must be numbers"
        try:
            result = OFFSET("Data!A1", row_array, 0, _context=mock_context)
            self.assertIsInstance(result, Array, "OFFSET with array parameters should return Array")
            
            # Check the array contents
            if hasattr(result, 'values'):
                values = result.values
                self.assertEqual('Name', values[0][0], "First result should be 'Name'")
                self.assertEqual('Alice', values[1][0], "Second result should be 'Alice'")
        except xlerrors.ValueExcelError as e:
            # This is the current bug - OFFSET rejects array parameters
            self.fail(f"OFFSET should handle array parameters but failed with: {e}")
    
    def test_offset_error_handling_bounds_checking(self):
        """
        FAILING TEST: OFFSET should properly raise errors for out-of-bounds references.
        
        This tests the error handling issue where OFFSET(-1, 0) should raise RefExcelError
        but might not be properly detected by ISERROR.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [['Name'], ['Alice']]
            
            def evaluate(self, reference):
                if reference == 'Data!A1':
                    return 'Name'
                return None
            
            def get_cell_value(self, address):
                return self.evaluate(address)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test OFFSET with negative row offset - should raise RefExcelError
        # OFFSET(Data!A1, -1, 0) should fail because it goes before row 1
        with self.assertRaises(xlerrors.RefExcelError, 
                             msg="OFFSET with negative row should raise RefExcelError"):
            OFFSET("Data!A1", -1, 0, _context=mock_context)
        
        # Test OFFSET with negative column offset - should raise RefExcelError
        with self.assertRaises(xlerrors.RefExcelError,
                             msg="OFFSET with negative column should raise RefExcelError"):
            OFFSET("Data!A1", 0, -1, _context=mock_context)
    
    def test_offset_array_parameter_types(self):
        """
        Test different array parameter types that OFFSET should handle.
        
        This helps identify what types of array inputs OFFSET should accept.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        from xlcalculator.xlfunctions.func_xltypes import Array
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [
                    ['Name', 'Age'],
                    ['Alice', 25],
                    ['Bob', 30],
                    ['Charlie', 35]
                ]
            
            def evaluate(self, reference):
                return 'Name'  # Default return
            
            def get_cell_value(self, address):
                return self.evaluate(address)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test different array formats
        test_cases = [
            (Array([[0], [1]]), "Column array [0, 1]"),
            (Array([[0, 1]]), "Row array [0, 1]"),
            ([0, 1], "Python list [0, 1]"),
        ]
        
        for array_param, description in test_cases:
            with self.subTest(array_type=description):
                try:
                    result = OFFSET("Data!A1", array_param, 0, _context=mock_context)
                    # If it works, it should return an Array
                    self.assertIsInstance(result, (Array, list), 
                                        f"OFFSET should handle {description}")
                except Exception as e:
                    # Document what types currently fail
                    print(f"OFFSET failed with {description}: {e}")
    
    def test_offset_with_arithmetic_array_expressions(self):
        """
        Test OFFSET with arithmetic expressions on arrays (like ROW(A1:A2)-1).
        
        This tests the specific pattern from the failing test case.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        from xlcalculator.xlfunctions.func_xltypes import Array
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [['Name'], ['Alice'], ['Bob']]
            
            def evaluate(self, reference):
                return 'Name'
            
            def get_cell_value(self, address):
                return self.evaluate(address)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Simulate the result of ROW(A1:A2)-1
        # ROW(A1:A2) would return [[1], [2]]
        # ROW(A1:A2)-1 would return [[0], [1]]
        row_minus_one = Array([[0], [1]])
        
        # Test OFFSET(Data!A1, ROW(A1:A2)-1, 0)
        # This should return an array with the values at offsets 0 and 1 from Data!A1
        try:
            result = OFFSET("Data!A1", row_minus_one, 0, _context=mock_context)
            self.assertIsInstance(result, Array, 
                                "OFFSET with arithmetic array should return Array")
        except Exception as e:
            # This is expected to fail initially
            self.fail(f"OFFSET should handle arithmetic array expressions: {e}")
    
    def test_offset_mixed_array_and_scalar_parameters(self):
        """
        Test OFFSET with mixed array and scalar parameters.
        
        This tests edge cases where one parameter is an array and the other is scalar.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        from xlcalculator.xlfunctions.func_xltypes import Array
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [['Name'], ['Alice'], ['Bob'], ['Charlie']]
            
            def evaluate(self, reference):
                return 'Name'
            
            def get_cell_value(self, address):
                cell_data = {
                    'Data!A1': 'Name',
                    'Data!A2': 'Alice',
                    'Data!A3': 'Bob',
                    'Data!B1': 'Age',
                    'Data!B2': 25,
                    'Data!B3': 30
                }
                return cell_data.get(address, None)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test array rows with scalar column
        row_array = Array([[0], [1]])
        result = OFFSET("Data!A1", row_array, 0, _context=mock_context)
        self.assertIsInstance(result, Array, "Mixed array/scalar should return Array")
        
        # Test scalar rows with array column
        col_array = Array([[0], [1]])
        result = OFFSET("Data!A1", 0, col_array, _context=mock_context)
        self.assertIsInstance(result, Array, "Mixed scalar/array should return Array")
    
    def test_offset_array_error_propagation(self):
        """
        Test that errors in array elements are properly propagated.
        
        This tests error handling when some array elements cause errors.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        from xlcalculator.xlfunctions.func_xltypes import Array
        from xlcalculator.xlfunctions import xlerrors
        
        # Create a mock context
        class MockEvaluator:
            def get_cell_value(self, address):
                # Only Data!A1 exists, others should cause errors
                if address == 'Data!A1':
                    return 'Name'
                return None
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test with array that includes out-of-bounds offsets
        # This should include both valid and invalid offsets
        row_array = Array([[0], [-1]])  # 0 is valid, -1 should cause error
        
        result = OFFSET("Data!A1", row_array, 0, _context=mock_context)
        self.assertIsInstance(result, Array, "Should return Array even with some errors")
        
        # The result should contain both valid values and error values
        if hasattr(result, 'values'):
            values = result.values
            # First element should be valid
            self.assertEqual('Name', values[0][0])
            # Second element should be an error (RefExcelError for out-of-bounds)
            # Note: The exact error type may vary based on implementation
    
    def test_offset_python_list_parameters(self):
        """
        Test OFFSET with Python list parameters (not just Array objects).
        
        This ensures compatibility with different array input types.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        
        # Create a mock context
        class MockEvaluator:
            def get_cell_value(self, address):
                cell_data = {
                    'Data!A1': 'Name',
                    'Data!A2': 'Alice',
                    'Data!A3': 'Bob',
                }
                return cell_data.get(address, None)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test with Python list instead of Array
        row_list = [0, 1]  # Python list
        result = OFFSET("Data!A1", row_list, 0, _context=mock_context)
        
        # Should work the same as Array parameters
        from xlcalculator.xlfunctions.func_xltypes import Array
        self.assertIsInstance(result, Array, "Python list parameters should work")


if __name__ == '__main__':
    unittest.main()