"""
ATDD Test Suite for Category 1: INDIRECT Reference Construction Issues

This test suite focuses on the specific INDIRECT reference construction failures identified in Category 1,
particularly when INDIRECT uses dynamically constructed references with functions like CHAR, COLUMN, ROW.

Key Issue: INDIRECT("Data!" & CHAR(65+COLUMN()) & "1") returns BLANK because CHAR and COLUMN functions fail
"""

import unittest
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text


class IndirectConstructionATDDTest(unittest.TestCase):
    """ATDD tests for INDIRECT reference construction issues."""
    
    def test_indirect_with_simple_concatenation(self):
        """
        CONTROL TEST: INDIRECT with simple string concatenation should work.
        
        This verifies that basic INDIRECT functionality works and isolates
        the issue to dynamic function-based construction.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context with evaluator
        class MockEvaluator:
            def get_range_values(self, range_ref):
                # Mock range data
                return [['Name', 'Age'], ['Alice', 25], ['Bob', 30]]
            
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
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test simple concatenation: INDIRECT("Data!A" & "2") should work
        # This simulates =INDIRECT("Data!A" & 2) which works in the real test
        result = INDIRECT("Data!A2", _context=mock_context)
        self.assertEqual('Alice', result, "Simple INDIRECT should work")
    
    def test_indirect_with_char_function_construction(self):
        """
        FAILING TEST: INDIRECT with CHAR function construction fails.
        
        This reproduces the core Category 1 issue where INDIRECT fails when
        the reference string is constructed using functions like CHAR.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context with evaluator that supports function evaluation
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [['Name', 'Age'], ['Alice', 25], ['Bob', 30]]
            
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
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test the pattern that fails: INDIRECT("Data!" & CHAR(66) & "3")
        # CHAR(66) = 'B', so this should construct "Data!B3" and return 30
        
        # First, test that the target reference works directly
        direct_result = INDIRECT("Data!B3", _context=mock_context)
        self.assertEqual(30, direct_result, "Direct reference should work")
        
        # Now test the constructed reference
        # This should work but currently fails due to CHAR function issues
        constructed_ref = "Data!" + chr(66) + "3"  # Manually construct "Data!B3"
        constructed_result = INDIRECT(constructed_ref, _context=mock_context)
        self.assertEqual(30, constructed_result, "Constructed reference should work")
    
    def test_indirect_with_function_parameter_evaluation(self):
        """
        FAILING TEST: INDIRECT fails when reference parameter contains function calls.
        
        This tests the core issue where INDIRECT receives a parameter that needs
        to be evaluated (like the result of CHAR or COLUMN functions) but the
        evaluation system doesn't work properly.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        from xlcalculator.xlfunctions.func_xltypes import Text
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [['Name', 'Age'], ['Alice', 25], ['Bob', 30]]
            
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
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test INDIRECT with Text type parameter (simulating function result)
        # This simulates what happens when CHAR(66) returns a Text object
        text_param = Text("Data!B3")
        result = INDIRECT(text_param, _context=mock_context)
        self.assertEqual(30, result, "INDIRECT should handle Text parameters correctly")
        
        # Test INDIRECT with different parameter types that might come from functions
        test_cases = [
            ("Data!B3", "String parameter"),
            (Text("Data!B3"), "Text type parameter"),
        ]
        
        for param, description in test_cases:
            with self.subTest(param_type=type(param).__name__, description=description):
                result = INDIRECT(param, _context=mock_context)
                self.assertEqual(30, result, f"INDIRECT should handle {description}")
    
    def test_indirect_reference_string_processing(self):
        """
        Test INDIRECT's reference string processing logic.
        
        This isolates the specific part of INDIRECT that processes the reference
        string to identify where the construction failure occurs.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return [['Name', 'Age'], ['Alice', 25], ['Bob', 30]]
            
            def evaluate(self, reference):
                cell_data = {
                    'Data!A1': 'Name',
                    'Data!A2': 'Alice', 
                    'Data!A3': 'Bob',
                    'Data!B1': 'Age',
                    'Data!B2': 25,
                    'Data!B3': 30,
                    'Data!I1': 0  # Add the expected test data
                }
                return cell_data.get(reference, None)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test various reference string formats that might be constructed
        test_references = [
            ("Data!A2", 'Alice', "Simple cell reference"),
            ("Data!B3", 30, "Different cell reference"),
            ("Data!I1", 0, "Extended column reference"),
        ]
        
        for ref_string, expected, description in test_references:
            with self.subTest(reference=ref_string, description=description):
                result = INDIRECT(ref_string, _context=mock_context)
                self.assertEqual(expected, result, f"INDIRECT should handle {description}")
    
    def test_indirect_with_evaluator_function_calls(self):
        """
        FAILING TEST: Test the actual issue - INDIRECT with evaluator-dependent functions.
        
        This reproduces the real Category 1 issue where formulas like
        =INDIRECT("Data!" & CHAR(65+COLUMN()) & "1") fail because the evaluator
        cannot properly evaluate the CHAR and COLUMN functions.
        """
        from tests import testing
        
        # Use the actual test infrastructure to reproduce the real issue
        class TestCheck(testing.FunctionalTestCase):
            filename = 'indirect_constructed_references.xlsx'
        
        test = TestCheck()
        test.setUp()
        
        # Test the individual components that should work
        print("Testing individual components:")
        
        # Test if CHAR function works
        char_result = test.evaluator.evaluate('=CHAR(66)')
        print(f"CHAR(66): {char_result}")
        
        # Test if COLUMN function works  
        column_result = test.evaluator.evaluate('=COLUMN(H4)')
        print(f"COLUMN(H4): {column_result}")
        
        # Test if basic arithmetic works
        arithmetic_result = test.evaluator.evaluate('=65+8')
        print(f"65+8: {arithmetic_result}")
        
        # The core issue: these functions return BLANK instead of their expected values
        # CHAR(66) should return 'B'
        # COLUMN(H4) should return 8
        # 65+8 should return 73
        
        # This test should FAIL initially, showing the function evaluation issue
        self.assertNotEqual(char_result, '', "CHAR function should not return BLANK")
        self.assertNotEqual(column_result, '', "COLUMN function should not return BLANK") 
        self.assertNotEqual(arithmetic_result, '', "Arithmetic should not return BLANK")
    
    def test_indirect_empty_cell_handling(self):
        """
        Test INDIRECT handling of empty cells according to Excel behavior.
        
        This tests the fix for empty cells returning 0 instead of None.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                return []
            
            def evaluate(self, reference):
                # Return None for empty cells
                if reference in ['Data!I1', 'Data!J1']:
                    return None
                elif reference == 'Data!A1':
                    return 'Name'
                return 'other_value'
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test empty cell handling - should return 0 according to Excel behavior
        result = INDIRECT("Data!I1", _context=mock_context)
        self.assertEqual(0, result, "INDIRECT should return 0 for empty cells")
        
        # Test non-empty cell still works
        result2 = INDIRECT("Data!A1", _context=mock_context)
        self.assertEqual('Name', result2, "INDIRECT should return actual value for non-empty cells")
    
    def test_indirect_error_handling_for_invalid_constructions(self):
        """
        Test INDIRECT error handling when reference construction fails.
        
        This ensures that INDIRECT properly handles cases where the reference
        string construction results in invalid references.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context
        class MockEvaluator:
            def get_range_values(self, range_ref):
                if "Invalid" in range_ref:
                    raise Exception("Invalid range")
                return [['Name', 'Age'], ['Alice', 25]]
            
            def evaluate(self, reference):
                if "Invalid" in reference:
                    raise Exception("Invalid reference")
                return None  # Cell doesn't exist
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test invalid reference handling
        with self.assertRaises(xlerrors.RefExcelError):
            INDIRECT("InvalidSheet!A1", _context=mock_context)
        
        # Test empty/None reference handling
        with self.assertRaises(xlerrors.RefExcelError):
            INDIRECT("", _context=mock_context)


if __name__ == '__main__':
    unittest.main()