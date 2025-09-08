"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class ErrorHandlingTest(testing.FunctionalTestCase):
    """
    Manejo de Errores integration tests.
    
    Tests: 3 cases across 1 levels
    Category: error_handling
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "error_handling.xlsx"

    def test_4p(self):
        """Manejo de Errores: Manejo de errores con IFERROR/ISERROR"""
        
        # Manejo errores IFERROR+INDEX
        value = self.evaluator.evaluate('Tests!P1')
        self.assertEqual('Not Found', value, "=IFERROR(INDEX(Data!A1:E6, 10, 1), \"Not Found\") should return 'Not Found'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Detección errores IF+ISERROR+OFFSET
        value = self.evaluator.evaluate('Tests!P2')
        self.assertEqual('Error', value, "=IF(ISERROR(OFFSET(Data!A1, -1, 0)), \"Error\", \"OK\") should return 'Error'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # Manejo errores IFERROR+INDIRECT
        value = self.evaluator.evaluate('Tests!P3')
        self.assertEqual('Sheet Error', value, "=IFERROR(INDIRECT(\"InvalidSheet!A1\"), \"Sheet Error\") should return 'Sheet Error'")
        self.assertIsInstance(value, (str, Text), "Should be text")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # text validation
        text_value = self.evaluator.evaluate('Tests!P1')
        self.assertIsInstance(text_value, (str, Text))
