"""
Comprehensive integration tests for dynamic ranges.
These tests validate FAITHFUL Excel behavior of INDEX, OFFSET, and INDIRECT functions.

Generated from JSON test configuration.
"""

from .. import testing
from xlcalculator.xlfunctions import xlerrors
from xlcalculator.xlfunctions.func_xltypes import Array, Number, Text, Boolean

class IndirectDynamicTest(testing.FunctionalTestCase):
    """
    INDIRECT integration tests.
    
    Tests: 4 cases across 1 levels
    Category: indirect_dynamic
    Source: DYNAMIC_RANGES_DESIGN.md
    """
    filename = "indirect_dynamic.xlsx"

    def test_2h(self):
        """INDIRECT - Referencias Dinámicas: Referencias construidas"""
        
        # INDIRECT dinámico - concatenación
        value = self.evaluator.evaluate('Tests!H1')
        self.assertEqual('Alice', value, "=INDIRECT(\"Data!A\" & 2) should return 'Alice'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT dinámico - CHAR
        value = self.evaluator.evaluate('Tests!H2')
        self.assertEqual(30, value, "=INDIRECT(\"Data!\" & CHAR(66) & \"3\") should return 30")
        self.assertIsInstance(value, (int, float, Number), "Should be number")

        # INDIRECT dinámico - ROW
        value = self.evaluator.evaluate('Tests!H3')
        self.assertEqual('Charlie', value, "=INDIRECT(\"Data!A\" & ROW()) should return 'Charlie'")
        self.assertIsInstance(value, (str, Text), "Should be text")

        # INDIRECT dinámico - COLUMN
        value = self.evaluator.evaluate('Tests!H4')
        self.assertEqual('Score', value, "=INDIRECT(\"Data!\" & CHAR(65+COLUMN()) & \"1\") should return 'Score'")
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
        text_value = self.evaluator.evaluate('Tests!H1')
        self.assertIsInstance(text_value, (str, Text))

        # number validation
        number_value = self.evaluator.evaluate('Tests!H2')
        self.assertIsInstance(number_value, (int, float, Number))
