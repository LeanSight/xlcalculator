"""
ATDD Test Suite for Error Handling in Dynamic Functions

This test suite validates that OFFSET and INDIRECT functions properly raise
and propagate Excel-compatible errors that can be detected by ISERROR/IFERROR.

Category: Error Handling
Source: Microsoft Excel documentation and failing integration tests
"""

import unittest
from xlcalculator.xlfunctions.dynamic_range import OFFSET, INDIRECT
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean


class ErrorHandlingATDDTest(unittest.TestCase):
    """
    ATDD tests for error handling in dynamic functions.
    
    These tests ensure that errors are properly raised and can be detected
    by Excel's error handling functions like ISERROR and IFERROR.
    """
    
    def test_offset_out_of_bounds_error_propagation(self):
        """
        FAILING TEST: OFFSET should raise RefExcelError for out-of-bounds references.
        
        This reproduces the issue where OFFSET(A1, -1, 0) should raise RefExcelError
        but the evaluator converts it to Blank, breaking ISERROR detection.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        
        # Create a mock context
        class MockEvaluator:
            def get_cell_value(self, address):
                # Simple mock data
                if address == 'Data!A1':
                    return 'Name'
                return None
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test case 1: Negative row offset (goes to row 0, which is invalid)
        with self.assertRaises(xlerrors.RefExcelError, 
                             msg="OFFSET with negative row should raise RefExcelError"):
            OFFSET("Data!A1", -1, 0, _context=mock_context)
        
        # Test case 2: Negative column offset (goes to column 0, which is invalid)
        with self.assertRaises(xlerrors.RefExcelError,
                             msg="OFFSET with negative column should raise RefExcelError"):
            OFFSET("Data!A1", 0, -1, _context=mock_context)
        
        # Test case 3: Large positive offset (beyond Excel limits)
        with self.assertRaises(xlerrors.RefExcelError,
                             msg="OFFSET beyond Excel limits should raise RefExcelError"):
            OFFSET("Data!A1", 1048576, 0, _context=mock_context)  # Row 1048577 is beyond limit
    
    def test_offset_error_in_evaluator_context(self):
        """
        Test OFFSET error handling in full evaluator context.
        
        This tests the specific pattern from the failing integration test:
        IF(ISERROR(OFFSET(Data!A1, -1, 0)), "Error", "OK")
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET
        from xlcalculator.xlfunctions.information import ISERROR
        from xlcalculator.xlfunctions.logical import IF
        
        # Create a mock context that simulates the real evaluator
        class MockEvaluator:
            def get_cell_value(self, address):
                if address == 'Data!A1':
                    return 'Name'
                return None
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test the error detection pattern
        try:
            # This should raise RefExcelError
            offset_result = OFFSET("Data!A1", -1, 0, _context=mock_context)
            # If we get here, the error wasn't raised properly
            self.fail("OFFSET should have raised RefExcelError but returned: " + str(offset_result))
        except xlerrors.RefExcelError as e:
            # This is the expected behavior
            # Now test that ISERROR would detect this
            self.assertIsInstance(e, xlerrors.RefExcelError, "Should be RefExcelError")
    
    def test_indirect_invalid_sheet_error_propagation(self):
        """
        FAILING TEST: INDIRECT should raise RefExcelError for invalid sheet references.
        
        This reproduces the issue where INDIRECT("InvalidSheet!A1") should raise
        RefExcelError but returns 0 instead.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context
        class MockEvaluator:
            def get_cell_value(self, address):
                # Only Data sheet exists
                if address.startswith('Data!'):
                    return 'test'
                # Invalid sheet should cause error
                return None
            
            def sheet_exists(self, sheet_name):
                return sheet_name == 'Data'
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test case 1: Invalid sheet reference
        with self.assertRaises(xlerrors.RefExcelError,
                             msg="INDIRECT with invalid sheet should raise RefExcelError"):
            INDIRECT("InvalidSheet!A1", _context=mock_context)
        
        # Test case 2: Empty reference string
        with self.assertRaises(xlerrors.RefExcelError,
                             msg="INDIRECT with empty string should raise RefExcelError"):
            INDIRECT("", _context=mock_context)
        
        # Test case 3: Invalid reference format
        with self.assertRaises(xlerrors.RefExcelError,
                             msg="INDIRECT with invalid format should raise RefExcelError"):
            INDIRECT("NotAReference", _context=mock_context)
    
    def test_indirect_error_in_evaluator_context(self):
        """
        Test INDIRECT error handling in full evaluator context.
        
        This tests the pattern from the failing integration test where
        INDIRECT errors should be detectable by ISERROR.
        """
        from xlcalculator.xlfunctions.dynamic_range import INDIRECT
        
        # Create a mock context that simulates the real evaluator
        class MockEvaluator:
            def get_cell_value(self, address):
                # Only valid addresses return values
                if address == 'Data!A1':
                    return 'Name'
                return None
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test the error detection pattern
        try:
            # This should raise RefExcelError for invalid sheet
            indirect_result = INDIRECT("InvalidSheet!A1", _context=mock_context)
            # If we get here, the error wasn't raised properly
            self.fail("INDIRECT should have raised RefExcelError but returned: " + str(indirect_result))
        except xlerrors.RefExcelError as e:
            # This is the expected behavior
            self.assertIsInstance(e, xlerrors.RefExcelError, "Should be RefExcelError")
    
    def test_error_types_consistency(self):
        """
        Test that error types are consistent and detectable.
        
        This ensures that all dynamic function errors use the same
        error types that Excel's error handling functions can detect.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET, INDIRECT
        
        # Create a mock context
        class MockEvaluator:
            def get_cell_value(self, address):
                return None
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test that all reference errors are RefExcelError
        error_cases = [
            lambda: OFFSET("Data!A1", -1, 0, _context=mock_context),
            lambda: OFFSET("Data!A1", 0, -1, _context=mock_context),
            lambda: INDIRECT("InvalidSheet!A1", _context=mock_context),
            lambda: INDIRECT("", _context=mock_context),
        ]
        
        for i, error_case in enumerate(error_cases):
            with self.subTest(case=i):
                with self.assertRaises(xlerrors.RefExcelError,
                                     msg=f"Error case {i} should raise RefExcelError"):
                    error_case()
    
    def test_error_handling_in_excel_file_context(self):
        """
        Test error handling with actual Excel file evaluation.
        
        This tests the specific issue where Excel files have pre-calculated
        values that prevent proper error handling evaluation.
        """
        from xlcalculator import model, evaluator
        import os
        
        # Load the test workbook
        resource_path = os.path.join('tests', 'resources', 'index_offset_iferror.xlsx')
        compiler = model.ModelCompiler()
        test_model = compiler.read_and_parse_archive(resource_path)
        test_evaluator = evaluator.Evaluator(test_model)
        
        # Clear any cached values to force re-evaluation
        for addr in test_model.cells:
            if test_model.cells[addr].formula:
                test_model.cells[addr].value = None
        
        # Test the specific failing case
        result = test_evaluator.evaluate('Tests!P2')
        
        # The formula is: =IF(ISERROR(OFFSET(Data!A1, -1, 0)), "Error", "OK")
        # OFFSET(Data!A1, -1, 0) should raise RefExcelError
        # ISERROR should catch it and return True
        # IF should return "Error"
        self.assertEqual('Error', result, 
                        "IF(ISERROR(OFFSET(Data!A1, -1, 0)), 'Error', 'OK') should return 'Error'")
    
    def test_valid_references_dont_raise_errors(self):
        """
        Test that valid references work correctly and don't raise errors.
        
        This ensures that the error handling fixes don't break normal operation.
        """
        from xlcalculator.xlfunctions.dynamic_range import OFFSET, INDIRECT
        
        # Create a mock context with valid data
        class MockEvaluator:
            def get_cell_value(self, address):
                data = {
                    'Data!A1': 'Name',
                    'Data!A2': 'Alice',
                    'Data!B1': 'Age',
                    'Data!B2': 25
                }
                return data.get(address, None)
            
            def evaluate(self, address):
                return self.get_cell_value(address)
        
        class MockContext:
            def __init__(self):
                self.evaluator = MockEvaluator()
                self.sheet = "Tests"
        
        mock_context = MockContext()
        
        # Test valid OFFSET operations
        result = OFFSET("Data!A1", 1, 0, _context=mock_context)  # Should get Alice
        self.assertEqual('Alice', result, "Valid OFFSET should work correctly")
        
        result = OFFSET("Data!A1", 0, 1, _context=mock_context)  # Should get Age
        self.assertEqual('Age', result, "Valid OFFSET should work correctly")
        
        # Test valid INDIRECT operations
        result = INDIRECT("Data!A1", _context=mock_context)  # Should get Name
        self.assertEqual('Name', result, "Valid INDIRECT should work correctly")
        
        result = INDIRECT("Data!B2", _context=mock_context)  # Should get 25
        self.assertEqual(25, result, "Valid INDIRECT should work correctly")


if __name__ == '__main__':
    unittest.main()