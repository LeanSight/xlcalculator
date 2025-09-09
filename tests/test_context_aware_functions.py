"""
ATDD Acceptance Tests for Context-Aware Function Execution

These tests define the expected behavior for ROW() and COLUMN() functions
when they need to access actual cell coordinates instead of hardcoded values.

RED PHASE: These tests should FAIL with current implementation
GREEN PHASE: Implementation will make these tests pass
REFACTOR PHASE: Improve implementation while maintaining test passage
"""

import unittest
from xlcalculator import ModelCompiler, Evaluator
from xlcalculator.xlfunctions import xlerrors


class ContextAwareFunctionsAcceptanceTest(unittest.TestCase):
    """
    Acceptance tests for context-aware function execution.
    
    These tests verify that ROW() and COLUMN() functions can access
    actual cell coordinates through proper context injection.
    """
    
    def setUp(self):
        """Set up test model with known cell structure."""
        # Create a simple model programmatically for predictable testing
        from xlcalculator.model import Model
        from xlcalculator.xltypes import XLCell, XLFormula
        
        self.model = Model()
        
        # Create test data in known positions
        # Sheet1!A1 = "Test"
        self.model.set_cell_value("Sheet1!A1", "Test")
        
        # Sheet1!B2 = ROW() formula (should return 2)
        self.model.cells["Sheet1!B2"] = XLCell(
            "Sheet1!B2", 
            None,
            XLFormula("=ROW()", "Sheet1", "B2")
        )
        
        # Sheet1!C3 = COLUMN() formula (should return 3)  
        self.model.cells["Sheet1!C3"] = XLCell(
            "Sheet1!C3",
            None, 
            XLFormula("=COLUMN()", "Sheet1", "C3")
        )
        
        # Sheet1!D4 = ROW() formula (should return 4)
        self.model.cells["Sheet1!D4"] = XLCell(
            "Sheet1!D4",
            None,
            XLFormula("=ROW()", "Sheet1", "D4") 
        )
        
        # Sheet1!E5 = COLUMN() formula (should return 5)
        self.model.cells["Sheet1!E5"] = XLCell(
            "Sheet1!E5", 
            None,
            XLFormula("=COLUMN()", "Sheet1", "E5")
        )
        
        # Build the AST for formulas
        self.model.build_code()
        
        self.evaluator = Evaluator(self.model)
    
    def test_row_function_returns_actual_row_number(self):
        """
        ACCEPTANCE TEST: ROW() function returns actual row number of calling cell
        
        Expected behavior:
        - ROW() called from B2 should return 2
        - ROW() called from D4 should return 4
        
        Current implementation: Returns hardcoded values or fails
        Target implementation: Returns actual cell row coordinate
        """
        # Test ROW() from row 2
        result_b2 = self.evaluator.evaluate("Sheet1!B2")
        self.assertEqual(2, result_b2, 
                        "ROW() called from B2 should return row number 2")
        
        # Test ROW() from row 4  
        result_d4 = self.evaluator.evaluate("Sheet1!D4")
        self.assertEqual(4, result_d4,
                        "ROW() called from D4 should return row number 4")
    
    def test_column_function_returns_actual_column_number(self):
        """
        ACCEPTANCE TEST: COLUMN() function returns actual column number of calling cell
        
        Expected behavior:
        - COLUMN() called from C3 should return 3 (column C)
        - COLUMN() called from E5 should return 5 (column E)
        
        Current implementation: Returns hardcoded value 3
        Target implementation: Returns actual cell column coordinate
        """
        # Test COLUMN() from column C (3)
        result_c3 = self.evaluator.evaluate("Sheet1!C3") 
        self.assertEqual(3, result_c3,
                        "COLUMN() called from C3 should return column number 3")
        
        # Test COLUMN() from column E (5)
        result_e5 = self.evaluator.evaluate("Sheet1!E5")
        self.assertEqual(5, result_e5,
                        "COLUMN() called from E5 should return column number 5")
    
    def test_row_column_with_explicit_reference(self):
        """
        ACCEPTANCE TEST: ROW() and COLUMN() with explicit reference parameter
        
        Expected behavior:
        - ROW("A1") should return 1
        - COLUMN("A1") should return 1  
        - ROW("C5") should return 5
        - COLUMN("C5") should return 3
        
        This tests the reference parameter path, not context injection
        """
        # Test explicit references - these should work with current implementation
        self.assertEqual(1, self.evaluator.evaluate('ROW("A1")'))
        self.assertEqual(1, self.evaluator.evaluate('COLUMN("A1")'))
        self.assertEqual(5, self.evaluator.evaluate('ROW("C5")'))
        self.assertEqual(3, self.evaluator.evaluate('COLUMN("C5")'))


if __name__ == '__main__':
    unittest.main()