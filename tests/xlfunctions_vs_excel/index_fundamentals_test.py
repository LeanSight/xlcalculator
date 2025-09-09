"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndexFundamentalsTest(testing.FunctionalTestCase):
    """
    INDEX integration tests.
    
    Tests: 5 cases across 1 levels
    Category: index_fundamentals
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "index_fundamentals.xlsx"

    def test_1a(self):
        """INDEX - Casos Fundamentales: Valores individuales básicos"""
        
        # INDEX básico - valor numérico
        value = self.evaluator.evaluate('Tests!A1')
        self.assertEqual(25, value, "=INDEX(Data!A1:E6, 2, 2) should return 25")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDEX básico - texto
        value = self.evaluator.evaluate('Tests!A2')
        self.assertEqual('Bob', value, "=INDEX(Data!A1:E6, 3, 1) should return 'Bob'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX básico - boolean
        value = self.evaluator.evaluate('Tests!A3')
        self.assertEqual(True, value, "=INDEX(Data!A1:E6, 4, 5) should return True")
        self.assertIsInstance(value, (bool, Boolean), "Should be boolean")

        # INDEX básico - última fila
        value = self.evaluator.evaluate('Tests!A4')
        self.assertEqual('Eve', value, "=INDEX(Data!A1:E6, 6, 1) should return 'Eve'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDEX básico - primera fila
        value = self.evaluator.evaluate('Tests!A5')
        self.assertEqual('Active', value, "=INDEX(Data!A1:E6, 1, 5) should return 'Active'")
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
        # number validation
        number_value = self.evaluator.evaluate('Tests!A1')
        self.assertIsInstance(number_value, (int, float, Number))

        # text validation
        text_value = self.evaluator.evaluate('Tests!A2')
        self.assertIsInstance(text_value, (str, Text))

        # boolean validation
        boolean_value = self.evaluator.evaluate('Tests!A3')
        self.assertIsInstance(boolean_value, (bool, Boolean))
