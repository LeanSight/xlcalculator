"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndirectFundamentalsTest(testing.FunctionalTestCase):
    """
    INDIRECT integration tests.
    
    Tests: 4 cases across 1 levels
    Category: indirect_fundamentals
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "indirect_fundamentals.xlsx"

    def test_2g(self):
        """INDIRECT - Casos Fundamentales: Referencias directas"""
        
        # INDIRECT básico - valor numérico
        value = self.evaluator.evaluate('Tests!G1')
        self.assertEqual(25, value, "=INDIRECT(\"Data!B2\") should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDIRECT básico - texto
        value = self.evaluator.evaluate('Tests!G2')
        self.assertEqual('LA', value, "=INDIRECT(\"Data!C3\") should return 'LA'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT básico - boolean
        value = self.evaluator.evaluate('Tests!G3')
        self.assertEqual(True, value, "=INDIRECT(\"Data!E4\") should return True")
        self.assertIsInstance(value, (bool, Boolean), "Should be boolean")

        # INDIRECT básico - desde celda
        value = self.evaluator.evaluate('Tests!G4')
        self.assertEqual(25, value, "=INDIRECT(P1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

    def test_data_integrity(self):
        """Verify test data integrity."""
        # Auto-generated data validation
        self.assertEqual('Alice', self.evaluator.evaluate('Data!A2'))
        self.assertEqual(25, self.evaluator.evaluate('Data!B2'))
        self.assertEqual('NYC', self.evaluator.evaluate('Data!C2'))

    def test_type_consistency(self):
        """Verify data type consistency across test cases."""
        # Auto-generated type validation
        # number validation
        number_value = self.evaluator.evaluate('Tests!G1')
        self.assertIsInstance(number_value, (int, float, Number))

        # text validation
        text_value = self.evaluator.evaluate('Tests!G2')
        self.assertIsInstance(text_value, (str, Text))

        # boolean validation
        boolean_value = self.evaluator.evaluate('Tests!G3')
        self.assertIsInstance(boolean_value, (bool, Boolean))
