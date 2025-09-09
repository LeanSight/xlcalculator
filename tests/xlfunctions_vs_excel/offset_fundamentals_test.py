"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class OffsetFundamentalsTest(testing.FunctionalTestCase):
    """
    OFFSET integration tests.
    
    Tests: 5 cases across 1 levels
    Category: offset_fundamentals
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "offset_fundamentals.xlsx"

    def test_2d(self):
        """OFFSET - Casos Fundamentales: Valores individuales básicos"""
        
        # OFFSET básico - B2
        value = self.evaluator.evaluate('Tests!D1')
        self.assertEqual(25, value, "=OFFSET(Data!A1, 1, 1) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # OFFSET básico - desde B2
        value = self.evaluator.evaluate('Tests!D2')
        self.assertEqual('LA', value, "=OFFSET(Data!B2, 1, 1) should return 'LA'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET básico - horizontal
        value = self.evaluator.evaluate('Tests!D3')
        self.assertEqual('City', value, "=OFFSET(Data!A1, 0, 2) should return 'City'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # OFFSET básico - esquina
        value = self.evaluator.evaluate('Tests!D4')
        self.assertEqual(False, value, "=OFFSET(Data!A1, 5, 4) should return False")
        self.assertIsInstance(value, (bool, Boolean), "Should be boolean")

        # OFFSET básico - offset negativo
        value = self.evaluator.evaluate('Tests!D5')
        self.assertEqual(85, value, "=OFFSET(Data!C3, -1, 1) should return 85")
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
        number_value = self.evaluator.evaluate('Tests!D1')
        self.assertIsInstance(number_value, (int, float, Number))

        # text validation
        text_value = self.evaluator.evaluate('Tests!D2')
        self.assertIsInstance(text_value, (str, Text))

        # boolean validation
        boolean_value = self.evaluator.evaluate('Tests!D4')
        self.assertIsInstance(boolean_value, (bool, Boolean))
